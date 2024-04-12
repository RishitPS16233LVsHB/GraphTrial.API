using Azure.Identity;
using Microsoft.Graph;

namespace GraphTrial.API.ServiceClient
{
    public static class GraphClient
    {
        private static string clientId = "<your client id here>";
        private static string clientSecret = "<your client secret here>";
        private static string tenantId = "<your tenant ID here>";


        /// <summary>
        /// creates a client which communicates with graph api to perform microsoft office work directly through the web app
        /// </summary>
        /// <returns></returns>
        public static GraphServiceClient CreateGraphServiceClient()
        {
            try
            {
                var tokenCredentialsOptions = new TokenCredentialOptions()
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredentials = new ClientSecretCredential(tenantId, clientId, clientSecret, tokenCredentialsOptions);
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                return new GraphServiceClient(clientSecretCredentials, scopes);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
