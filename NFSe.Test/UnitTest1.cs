using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using GInfes.Sil;
using System.IO;
using System.Xml;

namespace NFSe.Test
{
    [TestClass]
    public class UnitTest1
    {
        private string certName
        {
            get
            {
                return "CN=O3 SAUDE E BEM ESTAR LTDA:30911672000110, OU=Certificado PJ A1, OU=Videoconferencia, OU=19652495000161, OU=AC INFOCO DIGITAL v5, L=Sao Bernardo do Campo, S=SP, O=ICP-Brasil, C=BR";
            }
        }

        private string Ccnpj
        {
            get
            {
                return "30911672000110";
            }
        }

        private string cIE
        {
            get
            {
                return "264656";
            }
        }

        [TestMethod]
        public void CertificadoNaoEncontrado()
        {
            string xml = string.Empty;
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.EnvioLoteRPS(xml, string.Empty, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.NaoEncontrado.ObterDescricao()));
        }

        [TestMethod]
        public void ErroAssinaturaXml()
        {
            string xml = string.Empty;
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.EnvioLoteRPS(xml, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Erro.ObterDescricao()));
        }

        [TestMethod]
        public void SucessoEnvio()
        {
            string xml = File.ReadAllText(@"C:\Users\SergioEli\source\repos\NFSe.Sil\RPS.xml");
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.EnvioLoteRPS(xml, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Sucesso.ObterDescricao()));
        }

        [TestMethod]
        public void ErroConsultaLote()
        {
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.ConsultaLoteRPS("12861282", Ccnpj, cIE, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Erro.ObterDescricao()));
        }

        [TestMethod]
        public void SucessoConsultaLote()
        {
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.ConsultaLoteRPS("12862021", Ccnpj, cIE, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Sucesso.ObterDescricao()));
        }


        [TestMethod]
        public void ErroCancelaLote()
        {
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.CancelaNFSE("12861290", Ccnpj, cIE, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Erro.ObterDescricao()));

        }

        [TestMethod]
        public void SucessoCancelaLote()
        {
            GInfes.Sil.GInfes gInfes = new GInfes.Sil.GInfes();
            string cMensagem = gInfes.CancelaNFSE("12862015", Ccnpj, cIE, certName, "1");
            Assert.IsTrue(cMensagem.Contains(StatusIntegracao.Sucesso.ObterDescricao()));
        }

    }
}
