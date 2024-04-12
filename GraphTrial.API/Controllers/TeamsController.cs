using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph.Models;
using Microsoft.Graph.External;
using Microsoft.Graph;
using Azure.Identity;
using GraphTrial.API.ServiceClient;
using GraphTrial.API.Common;

namespace GraphTrial.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TeamsController : ControllerBase
    {
        public ResponseResult result = new ResponseResult();
        public GraphServiceClient graphClient { get; set; }
        public TeamsController()
        {
            graphClient = GraphClient.CreateGraphServiceClient();
        }

        /// <summary>
        /// Creates a team with a name and description 
        /// </summary>
        /// <returns></returns>
        // code written directly from microsoft's website for graph api samples
        [HttpPost]
        [Route("CreateTeam")]
        public async Task<ResponseResult> CreateTeam()
        {
            try
            {
                
                var requestBody = new Team
                {
                    DisplayName = "My Sample Team",
                    Description = "My Sample Team’s Description",
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "template@odata.bind" , "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                        },
                    },
                };
                Team team = await graphClient.Teams.PostAsync(requestBody);
                result.Data = team;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex) {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Adds a new conversation team member to the team using team member id(office Id or also you can provide user principal name)
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="teamMemberId"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("AddTeamMember/{teamId}/{teamMemberId}")]
        public async Task<ResponseResult> AddTeamMember(string teamId, string teamMemberId)
        {
            try
            {
                var requestBody = new AadUserConversationMember
                {
                    OdataType = "#microsoft.graph.aadUserConversationMember",
                    Roles = new List<string>
                    {
                        "owner",
                    },
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users('{teamMemberId}')"
                        },
                    },
                };

                ConversationMember conversationMember = await graphClient.Teams[teamId].Members.PostAsync(requestBody);
                result.Data = conversationMember;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Get team details
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetTeam/{teamId}")]
        public async Task<ResponseResult> GetTeam(string teamId)
        {
            try
            {
                var team = await graphClient.Teams[teamId].GetAsync();
                result.Data = team;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Gets list of conversation team members in the team
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetTeamMembers/{teamId}")]
        public async Task<ResponseResult> GetTeamMembers(string teamId)
        {
            try
            {
                ConversationMemberCollectionResponse teamMembers = await graphClient.Teams[teamId].Members.GetAsync();
                result.Data = teamMembers;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetTeamMember/{teamId}/{conversationMemberId}")]
        public async Task<ResponseResult> GetTeamMembers(string teamId, string conversationMemberId)
        {
            try
            {
                ConversationMember teamMembers = await graphClient.Teams[teamId].Members[conversationMemberId].GetAsync();
                result.Data = teamMembers;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Deletes a team
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpDelete]
        [Route("DeleteTeam/{teamId}")]
        public async Task<ResponseResult> DeleteTeam(string teamId)
        {
            try
            {
                await graphClient.Teams[teamId].DeleteAsync(); // returns void
                result.Data = "Deleted successfully";
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Removes a conversation team member from the team
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="conversationTeamMemberId"></param>
        /// <returns></returns>
        [HttpDelete]
        [Route("RemoveTeamMember/{teamId}/{conversationTeamMemberId}")]
        public async Task<ResponseResult> RemoveTeamMember(string teamId, string conversationTeamMemberId)
        {
            try
            {
                                await graphClient.Teams[teamId].Members[conversationTeamMemberId].DeleteAsync(); // returns void
                result.Data = "Removed team member successfully";
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Updates team details
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpPatch]
        [Route("UpdateTeam/{teamId}")]
        public async Task<ResponseResult> UpdateTeam(string teamId)
        {
            try
            {
                var requestBody = new Team
                {
                    MemberSettings = new TeamMemberSettings
                    {
                        AllowCreateUpdateChannels = true,
                    },
                    MessagingSettings = new TeamMessagingSettings
                    {
                        AllowUserEditMessages = true,
                        AllowUserDeleteMessages = true,
                    },
                    FunSettings = new TeamFunSettings
                    {
                        AllowGiphy = true,
                        GiphyContentRating = GiphyRatingType.Strict,
                    },
                };

                Team team = await graphClient.Teams[teamId].PatchAsync(requestBody);
                result.Data = team;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

        /// <summary>
        /// Updates conversation team member details in the team
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="conversationTeamMemberId"></param>
        /// <returns></returns>
        [HttpPatch]
        [Route("UpdateTeamMember/{teamId}/{conversationTeamMemberId}")]
        public async Task<ResponseResult> UpdateTeamMember(string teamId,string conversationTeamMemberId)
        {
            try
            {
                
                var requestBody = new AadUserConversationMember
                {
                    OdataType = "#microsoft.graph.aadUserConversationMember",
                    Roles = new List<string>
                    {
                        "owner",
                    },
                };

                ConversationMember team = await graphClient.Teams[teamId].Members[conversationTeamMemberId].PatchAsync(requestBody);
                result.Data = team;
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                result.Result = ResponseFlag.Error;
            }

            return result;
        }

    }
}
