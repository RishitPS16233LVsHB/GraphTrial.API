using Azure.Identity;
using Microsoft.Graph;

namespace GraphTrial.API.ServiceClient
{
    public static class GraphClient
    {
        private static string ClientID = "<client id>";
        private static string ClientSecret = "<client secret>";
        private static string TenantID = "<tenant id>";

        // send message access token for sending messages in channels and chats
        public static string AccessToken = @"<access token>";

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

                var clientSecretCredentials = new ClientSecretCredential(TenantID, ClientID, ClientSecret, tokenCredentialsOptions);
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
