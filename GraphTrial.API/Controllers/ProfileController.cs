using GraphTrial.API.Common;
using GraphTrial.API.ServiceClient;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Threading.Channels;

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


        [HttpGet]
        [Route("/Me")]
        public async Task<ResponseResult> Me()
        {
           try
           {
                //    result.Data = await graphClient.Me.GetAsync();
                //    result.Result = ResponseFlag.Success;
                //}
                //catch (Exception ex)
                //{
                //    result.Result = ResponseFlag.Error;
                //    result.Message = ex.Message;
                //}
                //return result;
                var res = "";
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {ServiceClient.GraphClient.AccessToken}");
                    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");

                    var response = await client.GetAsync("me");
                    if (response.IsSuccessStatusCode)
                    {
                        res = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        res = $"Failed to fetch user details. Status code: {response.StatusCode}";
                    }
                }

                result.Data = res;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

    }
}
