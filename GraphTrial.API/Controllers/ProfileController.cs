using GraphTrial.API.Common;
using GraphTrial.API.ServiceClient;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Threading.Channels;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Azure.Identity;
using System.Net.Http;
using GraphTrial.API.Entities;


namespace GraphTrial.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProfileController : ControllerBase
    {
        public ResponseResult result = new ResponseResult();
        public GraphServiceClient graphClient { get; set; }
        public ProfileController()
        {
            graphClient = GraphClient.CreateGraphServiceClient();
        }


        /// <summary>
        /// For ssm account use only
        /// </summary>
        /// <param name="tokenRequest"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> GetAccessToken([FromBody] Credential tokenRequest)
        {
            string tenant = GraphClient.TenantID;
            string clientId = GraphClient.ClientID;
            string url = $"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token";

            var formData = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", clientId),
                    new KeyValuePair<string, string>("scope", "user.read openid profile offline_access"),
                    new KeyValuePair<string, string>("username", tokenRequest.username),
                    new KeyValuePair<string, string>("password", tokenRequest.password),
                    new KeyValuePair<string, string>("grant_type", "password"),
                    new KeyValuePair<string, string>("client_secret", GraphClient.ClientSecret),
                }
            );

            using (var client = new HttpClient())
            {
                var response = await client.PostAsync(url, formData);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    return Ok(responseBody);
                }
                else
                {
                    return StatusCode((int)response.StatusCode, responseBody);
                }
            }
        }

    }
}
