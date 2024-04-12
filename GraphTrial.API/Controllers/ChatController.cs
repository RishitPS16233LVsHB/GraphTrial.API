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
    public class ChatController : ControllerBase
    {
        public ResponseResult result;
        public GraphServiceClient graphClient;
        public ChatController() 
        { 
            result = new ResponseResult();
            graphClient = GraphClient.CreateGraphServiceClient();
        }


        /// <summary>
        /// Get List of Teams Chats for the given user
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{userId}/GetChats")]
        public async Task<ResponseResult> GetChats(string userId)
        {
            try 
            { 
                result.Data = await graphClient.Users[userId].Chats.GetAsync();
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex) 
            { 
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;    
            }
            return result;
        }

        /// <summary>
        /// Delete Teams chats
        /// </summary>
        /// <param name="chatId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{chatId}/RemoveChat")]
        public async Task<ResponseResult> RemoveChat(string chatId)
        {
            try
            {
                await graphClient.Chats[chatId].DeleteAsync();
                result.Data = "Chat removed successfully!";
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Creates chat
        /// </summary>
        /// <param name="userId"></param>
        /// <param name="conversationMemberId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{userId}/CreateChat/{conversationMemberId}")]
        public async Task<ResponseResult> CreateChat(string userId,string conversationMemberId)
        {
            try
            {
                var requestBody = new Chat
                {
                    ChatType = ChatType.OneOnOne,
                    Members = new List<ConversationMember>
                    {
                        new AadUserConversationMember
                        {
                            OdataType = "#microsoft.graph.aadUserConversationMember",
                            Roles = new List<string>
                            {
                                "owner",
                            },
                            AdditionalData = new Dictionary<string, object>
                            {
                                {
                                    "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users('{userId}')"
                                },
                            },
                        },
                        new AadUserConversationMember
                        {
                            OdataType = "#microsoft.graph.aadUserConversationMember",
                            Roles = new List<string>
                            {
                                "owner",
                            },
                            AdditionalData = new Dictionary<string, object>
                            {
                                {
                                    "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users('{conversationMemberId}')"
                                },
                            },
                        },
                    },
                };
                result.Data = await graphClient.Chats.PostAsync(requestBody);
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Adds a member to the teams chat
        /// </summary>
        /// <param name="chatId"></param>
        /// <param name="conversationMemberId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{chatId}/AddMember/{conversationMemberId}")]
        public async Task<ResponseResult> AddMember(string chatId, string conversationMemberId)
        {
            try
            {
                var requestBody = new AadUserConversationMember
                {
                    OdataType = "#microsoft.graph.aadUserConversationMember",
                    VisibleHistoryStartDateTime = DateTimeOffset.Now,
                    Roles = new List<string>
                    {
                        "owner",
                    },
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users/{conversationMemberId}"
                        },
                    },
                };
                result.Data = await graphClient.Chats[chatId].Members.PostAsync(requestBody);
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Removes a team member from the chat
        /// </summary>
        /// <param name="chatId"></param>
        /// <param name="conversationMemberId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{chatId}/RemoveMember/{conversationMemberId}")]
        public async Task<ResponseResult> RemoveChat(string chatId,string conversationMemberId)
        {
            try
            {
                await graphClient.Chats[chatId].Members[conversationMemberId].DeleteAsync();
                result.Data = "Chat member removed successfully!";
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Lists all members associated with the chat
        /// </summary>
        /// <param name="chatId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("{chatId}/ListMembers")]
        public async Task<ResponseResult> ListMembers(string chatId)
        {
            try
            {
                result.Data = await graphClient.Chats[chatId].Members.GetAsync();
                result.Result = ResponseFlag.Success;
            }
            catch (Exception ex)
            {
                result.Result = ResponseFlag.Error;
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Sends message in the chat
        /// </summary>
        /// <param name="chatId"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("{chatId}/SendMessage")]
        public async Task<ResponseResult> SendMessage(string chatId, [FromBody] string message)
        {
            try
            {
                var requestBody = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        Content = message,
                    },
                };

                var res = await graphClient.Chats[chatId].Messages.PostAsync(requestBody);
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


        /// <summary>
        /// Gets message of the chat
        /// </summary>
        /// <param name="chatId"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("{chatId}/GetMessages")]
        public async Task<ResponseResult> GetMessages(string chatId)
        {
            try
            {                
                var res = await graphClient.Chats[chatId].Messages.GetAsync();
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
