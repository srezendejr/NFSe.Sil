using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace GInfes.Sil
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class GInfes
    {
        public string EnvioLoteRPS(string xml, string certName, string ambiente = "2")
        {
            string cRetorno = string.Empty;
            try
            {
                X509Certificate2 Certificado = LocalizaCertificadoValido(certName);
                if (Certificado == null)
                    throw new Exception(StatusIntegracao.NaoEncontrado.ObterDescricao());
                var xmlIntegracao = AssinaXmlDigitalmente(xml, Certificado);
                dynamic serviceGinfes;
                if (ambiente == "1")
                    serviceGinfes = new HomologacaoGinfes.ServiceGinfesImplService();
                else
                    serviceGinfes = new ProducaoGinfes.ServiceGinfesImplService();
                //HomologacaoGinfes.ServiceGinfesImplService serviceGinfes = new HomologacaoGinfes.ServiceGinfesImplService();
                serviceGinfes.Timeout = 5000;
                serviceGinfes.ClientCertificates.Add(Certificado);
                serviceGinfes.Proxy = CriarProxy();
                cRetorno = serviceGinfes.RecepcionarLoteRpsV3(CriarCabecalho(), xmlIntegracao);
                if (cRetorno.Contains("DataRecebimento"))
                    cRetorno = $"{StatusIntegracao.Sucesso.ObterDescricao()}-{cRetorno}";
                else
                    cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()}-{cRetorno}";
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains(StatusIntegracao.NaoEncontrado.ObterDescricao()))
                    cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()} {ex.Message.ToString()}";
                else
                    cRetorno = ex.Message.ToString();
            }

            return cRetorno;

        }

        public string ConsultaLoteRPS(string cLote, string cCNPJ, string cIM, string certName, string ambiente = "2")
        {
            string cRetorno = string.Empty;
            try
            {
                X509Certificate2 Certificado = LocalizaCertificadoValido(certName);
                if (Certificado == null)
                    throw new Exception(StatusIntegracao.NaoEncontrado.ObterDescricao());
                string xml = CriarXmlConsultaLote(cLote, cCNPJ, cIM);
                var xmlIntegracao = AssinaXmlDigitalmente(xml, Certificado);
                dynamic serviceGinfes;
                if (ambiente == "1")
                    serviceGinfes =  new HomologacaoGinfes.ServiceGinfesImplService();
                else
                    serviceGinfes= new ProducaoGinfes.ServiceGinfesImplService();
                //HomologacaoGinfes.ServiceGinfesImplService serviceGinfes = new HomologacaoGinfes.ServiceGinfesImplService();
                serviceGinfes.Timeout = 5000;
                serviceGinfes.ClientCertificates.Add(Certificado);
                serviceGinfes.Proxy = CriarProxy();
                cRetorno = serviceGinfes.ConsultarLoteRpsV3(CriarCabecalho(), xmlIntegracao);
                if (!cRetorno.Contains("Correcao"))
                    cRetorno = $"{StatusIntegracao.Sucesso.ObterDescricao()}-{cRetorno}";
                else
                    cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()}-{cRetorno}";

            }
            catch (Exception ex)
            {
                cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()} {ex.Message.ToString()}";
            }

            return cRetorno;
        }

        public string CancelaNFSE(string cNFSe, string CNPJ, string InscrMunicipal, string certName, string ambiente = "2")
        {
            string cRetorno = string.Empty;
            try
            {
                X509Certificate2 Certificado = LocalizaCertificadoValido(certName);
                if (Certificado == null)
                    throw new Exception(StatusIntegracao.NaoEncontrado.ObterDescricao());
                string xml = CriarXmlCancelarLoteV2(cNFSe, CNPJ, InscrMunicipal);
                var xmlIntegracao = AssinaXmlDigitalmente(xml, Certificado);
                dynamic serviceGinfes;
                if (ambiente == "1")
                    serviceGinfes = new HomologacaoGinfes.ServiceGinfesImplService();
                else
                    serviceGinfes = new ProducaoGinfes.ServiceGinfesImplService();
                //HomologacaoGinfes.ServiceGinfesImplService serviceGinfes = new HomologacaoGinfes.ServiceGinfesImplService();
                serviceGinfes.Timeout = 5000;
                serviceGinfes.ClientCertificates.Add(Certificado);
                serviceGinfes.Proxy = CriarProxy();
                cRetorno = serviceGinfes.CancelarNfse(xmlIntegracao);
                if (cRetorno.Contains("<ns5:Sucesso>true</ns5:Sucesso>"))
                    cRetorno = $"{StatusIntegracao.Sucesso.ObterDescricao()}-{cRetorno}";
                else
                    cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()}-{cRetorno}";
            }
            catch (Exception ex)
            {
                cRetorno = $"{StatusIntegracao.Erro.ObterDescricao()} {ex.Message.ToString()}";
            }

            return cRetorno;
        }

        private IWebProxy CriarProxy()
        {
            IWebProxy iw = WebRequest.GetSystemWebProxy();
            NetworkCredential nc = CredentialCache.DefaultNetworkCredentials;
            iw.Credentials = nc;
            return iw;
        }

        private X509Certificate2 LocalizaCertificadoValido(string certName)
        {
            // Get the certificate store for the current user.
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);

                // Place all certificates in an X509Certificate2Collection object.
                X509Certificate2Collection certCollection = store.Certificates;

                // If using a certificate with a trusted root you do not need to FindByTimeValid, instead:
                // currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certName, true);
                // crea una collezione di tutti i certificati "Validi" ossia con data non scaduta
                X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);

                // prende dalla collezione il certificato di cui ho passato l'Assunto certName
                X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certName, false);

                if (signingCert.Count == 0)
                    return null;

                // Return the first certificate in the collection, has the right name and is current.
                return signingCert[0];
            }
            finally
            {
                store.Close();
            }
        }

        [Obsolete]
        private string AssinaLoteRpsDigitalmente(string xml, X509Certificate2 Certificado)
        {
            //Cria um novo ArquivoXml para manusea-lo
            string cRetorno = string.Empty;
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = false;
            doc.LoadXml(xml);

            //Pega o certificado válido para usar na assinatura do xml
            X509Certificate2 Cert = Certificado;
            if (Cert == null)
                throw new Exception(StatusIntegracao.NaoEncontrado.ObterDescricao());
            string x = Cert.GetKeyAlgorithm();

            //Cria um objeto de assinatura do arquivo
            SignedXml signedXml = new SignedXml(doc);

            //Adiciona a chave do certificado no objeto de assinatura do arquivo
            signedXml.SigningKey = Cert.PrivateKey;
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            //Cria um objeto de referencia para assinatura
            Reference reference = new Reference();
            //Pega os atributos da tag LoteRps
            XmlAttributeCollection _Uri = doc.GetElementsByTagName("ns3:LoteRps").Item(0).Attributes;
            foreach (XmlAttribute _atributo in _Uri)
            {
                if (_atributo.Name == "Id")
                {
                    reference.Uri = "#" + _atributo.InnerText; //Comentada para realizar a integração.
                    //reference.Uri = "";
                }
            }

            //Adiciona o envelope de transformação da referencia
            XmlDsigEnvelopedSignatureTransform env = new XmlDsigEnvelopedSignatureTransform();
            reference.AddTransform(env);
            XmlDsigC14NTransform c14 = new XmlDsigC14NTransform();
            reference.AddTransform(c14);

            //Adiciona a referencia no objeto de assinatura
            signedXml.AddReference(reference);

            //Cria um objeto KeyInfo
            KeyInfo keyInfo = new KeyInfo();

            //Adiciona o certificado no objeto KeyInfo
            keyInfo.AddClause(new KeyInfoX509Data(Cert));

            //Adiciona o objeto KeyInfo no objeto de assinatura do arquivo
            signedXml.KeyInfo = keyInfo;

            //"Computa" a assinatura
            signedXml.ComputeSignature();

            //Cria um elemento do arquivo xml com a assinatura do arquivo
            XmlElement xmlDigitalSignature = signedXml.GetXml();
            //Adiciona a tag de assinatura no Documento Xml
            doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, true));
            cRetorno = doc.InnerXml;

            return cRetorno;
        }

        public static string AssinaXmlDigitalmente(string xml, X509Certificate2 Certificado)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement XmlNfse = doc.DocumentElement;
            if (Certificado == null)
                throw new Exception("Não existe um certificado válido!");
            SignedXml signedXml = new SignedXml(doc);
            signedXml.SigningKey = Certificado.PrivateKey;
            Reference reference = new Reference();
            reference.Uri = "";
            reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
            reference.AddTransform(new XmlDsigC14NTransform());
            signedXml.AddReference(reference);
            KeyInfo keyInfo = new KeyInfo();
            keyInfo.AddClause(new KeyInfoX509Data(Certificado));
            signedXml.KeyInfo = keyInfo;
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
            signedXml.ComputeSignature();
            XmlElement xmlSignature = doc.CreateElement("Signature", "http://www.w3.org/2000/09/xmldsig#");
            XmlElement xmlSignatureInfo = signedXml.SignedInfo.GetXml();
            XmlElement xmlKeyInfo = signedXml.KeyInfo.GetXml();
            XmlElement xmlSignatureValue = doc.CreateElement("SignatureValue", xmlSignature.NamespaceURI);
            string SignBase64 = Convert.ToBase64String(signedXml.Signature.SignatureValue);
            XmlText text = doc.CreateTextNode(SignBase64);
            xmlSignatureValue.AppendChild(text);
            xmlSignature.AppendChild(doc.ImportNode(xmlSignatureInfo, true));
            xmlSignature.AppendChild(xmlSignatureValue);
            xmlSignature.AppendChild(doc.ImportNode(xmlKeyInfo, true));
            XmlNfse.AppendChild(xmlSignature);

            doc.AppendChild(XmlNfse);
            return doc.InnerXml;
        }

        private string CriarCabecalho()
        {
            XmlDocument XmlCabecalho = new XmlDocument();
            XmlElement xmlRoot = XmlCabecalho.CreateElement("cabecalho", "http://www.ginfes.com.br/cabecalho_v03.xsd");
            XmlAttribute xmlAtt = XmlCabecalho.CreateAttribute("versao");
            xmlAtt.Value = "3";
            xmlRoot.Attributes.Append(xmlAtt);
            XmlElement XmlVersaoDados = XmlCabecalho.CreateElement("versaoDados");
            XmlVersaoDados.InnerText = "3";

            xmlRoot.AppendChild(XmlVersaoDados);
            XmlCabecalho.AppendChild(xmlRoot);
            return XmlCabecalho.InnerXml;
        }
        private string CriarXmlConsultaLote(string LoteRps, string Cnpj, string InscrMunicipal)
        {
            XmlDocument XmlConsulta = new XmlDocument();
            XmlElement xmlRoot = XmlConsulta.CreateElement("ConsultarLoteRpsEnvio", "http://www.ginfes.com.br/servico_consultar_lote_rps_envio_v03.xsd");
            XmlElement xmlPrestador = XmlConsulta.CreateElement("Prestador", "http://www.ginfes.com.br/servico_consultar_lote_rps_envio_v03.xsd");
            XmlElement xmlCnpj = XmlConsulta.CreateElement("Cnpj", "http://www.ginfes.com.br/tipos_v03.xsd");
            xmlCnpj.InnerText = Cnpj;
            XmlElement xmlInscMunic = XmlConsulta.CreateElement("InscricaoMunicipal", "http://www.ginfes.com.br/tipos_v03.xsd");
            xmlInscMunic.InnerText = InscrMunicipal;
            XmlElement xmlProtocolo = XmlConsulta.CreateElement("Protocolo", "http://www.ginfes.com.br/servico_consultar_lote_rps_envio_v03.xsd");
            xmlProtocolo.InnerText = LoteRps.ToString();
            XmlDeclaration xmlDec = XmlConsulta.CreateXmlDeclaration("1.0", "utf-8", null);

            XmlConsulta.AppendChild(xmlDec);
            xmlPrestador.AppendChild(xmlCnpj);
            xmlPrestador.AppendChild(xmlInscMunic);
            xmlRoot.AppendChild(xmlPrestador);
            xmlRoot.AppendChild(xmlProtocolo);
            XmlConsulta.AppendChild(xmlRoot);
            return XmlConsulta.InnerXml;
        }
        private string CriarXmlCancelarLoteV2(string NumeroNfse, string CNPJ, string InscrMunicipal)
        {
            XmlDocument XmlCancelamento = new XmlDocument();
            XmlElement xmlRoot = XmlCancelamento.CreateElement("CancelarNfseEnvio", "http://www.ginfes.com.br/servico_cancelar_nfse_envio");
            XmlElement xmlPrestador = XmlCancelamento.CreateElement("Prestador", "http://www.ginfes.com.br/servico_cancelar_nfse_envio");
            XmlElement xmlCnpj = XmlCancelamento.CreateElement("Cnpj", "http://www.ginfes.com.br/tipos");
            XmlElement xmlInscMunic = XmlCancelamento.CreateElement("InscricaoMunicipal", "http://www.ginfes.com.br/tipos");
            XmlElement xmlNumero = XmlCancelamento.CreateElement("NumeroNfse", "http://www.ginfes.com.br/servico_cancelar_nfse_envio");
            XmlDeclaration xmlDec = XmlCancelamento.CreateXmlDeclaration("1.0", "utf-8", null);
            xmlCnpj.InnerText = CNPJ;
            xmlInscMunic.InnerText = InscrMunicipal;
            xmlNumero.InnerText = NumeroNfse.ToString();

            xmlPrestador.AppendChild(xmlCnpj);
            xmlPrestador.AppendChild(xmlInscMunic);
            xmlRoot.AppendChild(xmlPrestador);
            xmlRoot.AppendChild(xmlNumero);
            XmlCancelamento.AppendChild(xmlDec);
            XmlCancelamento.AppendChild(xmlRoot);

            return XmlCancelamento.InnerXml;
        }
        [Obsolete]
        private string CriarXmlCancelarLoteV3(string NumeroNfse, string CNPJ, string InscrMunicipal, string CodMunicipio, string cMotivoCancelamento)
        {
            XmlDocument XmlCancelamento = new XmlDocument();
            XmlElement xmlRoot = XmlCancelamento.CreateElement("CancelarNfseEnvio", "http://www.ginfes.com.br/servico_cancelar_nfse_envio_v03.xsd");
            XmlElement xmlPedido = XmlCancelamento.CreateElement("Pedido");
            XmlElement xmlInfPedido = XmlCancelamento.CreateElement("InfPedidoCancelamento", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlIdentNf = XmlCancelamento.CreateElement("IdentificacaoNfse", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlNumero = XmlCancelamento.CreateElement("Numero", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlCnpj = XmlCancelamento.CreateElement("Cnpj", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlInscMunic = XmlCancelamento.CreateElement("InscricaoMunicipal", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlCodiCanc = XmlCancelamento.CreateElement("CodigoCancelamento", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlElement xmlCodiMun = XmlCancelamento.CreateElement("CodigoMunicipio", "http://www.ginfes.com.br/tipos_v03.xsd");
            XmlDeclaration xmlDec = XmlCancelamento.CreateXmlDeclaration("1.0", "utf-8", null);
            xmlNumero.InnerText = NumeroNfse.ToString();
            xmlCnpj.InnerText = CNPJ;
            xmlInscMunic.InnerText = InscrMunicipal;
            xmlCodiMun.InnerText = CodMunicipio;
            xmlCodiCanc.InnerText = cMotivoCancelamento;

            xmlIdentNf.AppendChild(xmlNumero);
            xmlIdentNf.AppendChild(xmlCnpj);
            xmlIdentNf.AppendChild(xmlInscMunic);
            xmlIdentNf.AppendChild(xmlCodiMun);
            xmlInfPedido.AppendChild(xmlIdentNf);
            xmlInfPedido.AppendChild(xmlCodiCanc);
            xmlPedido.AppendChild(xmlInfPedido);
            xmlRoot.AppendChild(xmlPedido);
            XmlCancelamento.AppendChild(xmlDec);
            XmlCancelamento.AppendChild(xmlRoot);

            return XmlCancelamento.InnerXml;
        }

        private static void ValidaXml(string xml)
        {
            try
            {
                string cpath = @"C:\Users\SergioEli\source\repos\NFSe.Sil\ClassLibrary1\schemas_v301_atual\";
                StringReader sr = new StringReader(xml);
                XmlReaderSettings xmlRS = new XmlReaderSettings();
                XmlUrlResolver xmlRes = new XmlUrlResolver();
                xmlRes.Credentials = CredentialCache.DefaultNetworkCredentials;
                xmlRS.Schemas.Add("http://www.ginfes.com.br/servico_enviar_lote_rps_envio_v03.xsd", string.Format("{0}{1}", cpath, "servico_enviar_lote_rps_envio_v03.xsd"));
                xmlRS.Schemas.Add("http://www.ginfes.com.br/tipos_v03.xsd", string.Format("{0}{1}", cpath, "tipos_v03.xsd"));
                xmlRS.Schemas.Add("http://www.ginfes.com.br/cabecalho_v03.xsd", string.Format("{0}{1}", cpath, "cabecalho_v03.xsd"));
                xmlRS.Schemas.Add("http://www.ginfes.com.br/servico_consultar_lote_rps_envio_v03.xsd", string.Format("{0}{1}", cpath, "servico_consultar_lote_rps_envio_v03.xsd"));
                xmlRS.Schemas.Add("http://www.ginfes.com.br/servico_consultar_situacao_lote_rps_resposta_v03.xsd", string.Format("{0}{1}", cpath, "servico_consultar_situacao_lote_rps_resposta_v03.xsd"));
                xmlRS.Schemas.Add("http://www.ginfes.com.br/servico_cancelar_nfse_envio_v03.xsd", string.Format("{0}{1}", cpath, "servico_cancelar_nfse_envio_v03.xsd"));
                xmlRS.XmlResolver = xmlRes;
                xmlRS.ValidationType = ValidationType.Schema;
                XmlReader xmlRR = XmlReader.Create(sr, xmlRS);
                XmlDocument NovoXml = new XmlDocument();
                NovoXml.Load(xmlRR);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
    public enum StatusIntegracao
    {
        [Description("0-Certificado não encontrado")]
        NaoEncontrado = 0,
        [Description("1-Sucesso")]
        Sucesso = 1,
        [Description("2-Erro:")]
        Erro = 2
    }
    public static class Commons
    {
        public static String ObterDescricao(this Enum valor)
        {
            FieldInfo fieldInfo = valor.GetType().GetField(valor.ToString());

            DescriptionAttribute[] atributos = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);

            return atributos.Length > 0 ? atributos[0].Description ?? "Nulo" : valor.ToString();
        }
    }
}
