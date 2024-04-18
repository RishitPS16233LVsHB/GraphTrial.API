using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph.Models;
using Microsoft.Graph.External;
using Microsoft.Graph;
using Azure.Identity;
using GraphTrial.API.ServiceClient;
using GraphTrial.API.Common;
using System.Diagnostics;
using GraphTrial.API.Entities;

namespace GraphTrial.API.Controllers
{
    [Route("api/[controller]/")]
    [ApiController]
    public class TeamsController : ControllerBase
    {
        public ResponseResult result = new ResponseResult();
        public GraphServiceClient graphClient { get; set; }
        public TeamsController()
        {
            graphClient = GraphClient.CreateGraphServiceClient();
        }

        [HttpGet]
        [Route("GetTeam/{userPrincipal}")]
        public async Task<ResponseResult> GetTeamsUnderUser(string userPrincipal)
        {
            try {
                var res = await graphClient.Users[userPrincipal].JoinedTeams.GetAsync();
                result.Data = res;
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
        /// Creates a team with a name and description 
        /// </summary>
        /// <returns></returns>
        // code written directly from microsoft's website for graph api samples
        [HttpPost]
        [Route("CreateTeam")]
        public async Task<ResponseResult> CreateTeam([FromBody] CreateTeam createTeam)
        {
            try
            {                
                var requestBody = new Team
                {
                    DisplayName = createTeam.TeamName,
                    Description = createTeam.TeamDescription,
                    Visibility = createTeam.IsPrivate ? TeamVisibilityType.Private : TeamVisibilityType.Public,
                    Members = new List<ConversationMember>()
                    {
                        new AadUserConversationMember
                        {
                            Roles = new List<string>()
                            {
                                "owner"
                            },
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{createTeam.OwnerUserPrincipal}')"}
                            }
                        }
                    },
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "template@odata.bind" , "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                        },
                    },
                };
                var team = await graphClient.Teams.PostAsync(requestBody);
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
        [HttpGet]
        [Route("{teamId}/AddTeamMember/{userPrincipal}")]
        public async Task<ResponseResult> AddTeamMember(string teamId, string userPrincipal)
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
                            "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users('{userPrincipal}')"
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
        [Route("{teamId}/GetTeam")]
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
        [Route("{teamId}/GetTeamMembers")]
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
        /// Deletes a team
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{teamId}/DeleteTeam")]
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
        [HttpGet]
        [Route("{teamId}/RemoveTeamMember/{conversationUserPrincipal}")]
        public async Task<ResponseResult> RemoveTeamMember(string teamId, string conversationUserPrincipal)
        {
            try
            {
                await graphClient.Teams[teamId].Members[conversationUserPrincipal].DeleteAsync(); // returns void
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
    }
}
