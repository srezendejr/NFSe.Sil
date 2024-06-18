using System.Security.Cryptography.X509Certificates;

namespace NFSe.Sil
{
    public class Ginfes
    {
        public string EnvioLoteRPS(string xml, string certName)
        {
            string cRetorno = string.Empty;
            X509Certificate2? Certificado = LocalizaCertificadoValido(certName);
            if (Certificado == null)
                cRetorno = "Certificado não encontrado";
            return cRetorno;
        }

        public string ConsultaLoteRPS(string xml, string certName)
        {
            string cRetorno = string.Empty;
            X509Certificate2? Certificado = LocalizaCertificadoValido(certName);
            if (Certificado == null)
                cRetorno = "Certificado não encontrado";
            return cRetorno;
        }

        public string CancelaLoteRPS(string xml, string certName)
        {
            string cRetorno = string.Empty;
            X509Certificate2? Certificado = LocalizaCertificadoValido(certName);
            if (Certificado == null)
                cRetorno = "Certificado não encontrado";
            return cRetorno;
        }

        private X509Certificate2? LocalizaCertificadoValido(string certName)
        {
            // Get the certificate store for the current user.
            X509Store store = new(StoreLocation.CurrentUser);
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
    }
}