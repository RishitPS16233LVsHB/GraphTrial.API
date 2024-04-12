using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using GraphTrial.API.ServiceClient;
using GraphTrial.API.Common;
using Azure.Identity;
using Microsoft.Graph.External;
using Microsoft.Graph.Models;
using Microsoft.Graph;


namespace GraphTrial.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ChannelController : ControllerBase
    {
        public ResponseResult result = new ResponseResult();
        public GraphServiceClient graphClient { get; set; }
        public ChannelController()
        {
            graphClient = GraphClient.CreateGraphServiceClient();
        }

        [HttpGet]
        [Route("GetAllChannels/{teamId}")]
        public async Task<ResponseResult> GetAllChannels(string teamId)
        {
            try 
            {
                var res = await graphClient.Teams["{team-id}"].AllChannels.GetAsync();
                result.Data = res;
                result.Result = ResponseFlag.Success;
            }   
            catch(Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        [HttpGet]
        [Route("GetAllSharedChannels/{teamId}")]
        public async Task<ResponseResult> GetSharedChannels(string teamId)
        {
            try
            {
                var res = await graphClient.Teams["{team-id}"].AllChannels.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Filter = "membershipType eq 'shared'";
                });
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

        [HttpGet]
        [Route("GetChannels/{teamId}")]
        public async Task<ResponseResult> GetChannels(string teamId)
        {
            try
            {
                var res = await graphClient.Teams["{team-id}"].Channels.GetAsync();
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

        [HttpGet]
        [Route("GetPrivateChannels/{teamId}")]
        public async Task<ResponseResult> GetPrivateChannels(string teamId)
        {
            try
            {
                var res = await graphClient.Teams["{team-id}"].Channels.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Filter = "membershipType eq 'private'";
                });
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
