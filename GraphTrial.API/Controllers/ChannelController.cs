﻿using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using GraphTrial.API.ServiceClient;
using GraphTrial.API.Common;
using Azure.Identity;
using Microsoft.Graph.External;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Graph.DeviceManagement.NotificationMessageTemplates.Item.SendTestMessage;


namespace GraphTrial.API.Controllers
{
    [Route("api/[controller]/{teamId}")]
    [ApiController]
    public class ChannelController : ControllerBase
    {
        public ResponseResult result = new ResponseResult();
        public GraphServiceClient graphClient { get; set; }
        public ChannelController()
        {
            graphClient = GraphClient.CreateGraphServiceClient();
        }

        /// <summary>
        /// Gets the list of all channels present for the team
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetChannels")]
        public async Task<ResponseResult> GetChannels(string teamId)
        {
            try
            {
                var res = await graphClient.Teams[teamId].Channels.GetAsync();
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
        /// Gets the general channel for the team
        /// </summary>
        /// <param name="teamId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetGeneralChannel")]
        public async Task<ResponseResult> GetGeneralChannel(string teamId)
        {
            try
            {
                var res = await graphClient.Teams[teamId].PrimaryChannel.GetAsync();
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
        /// Gets channel members of a given channel
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="channelId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetMembers/{channelId}")]
        public async Task<ResponseResult> GetChannelMembers(string teamId,string channelId)
        {
            try
            {
                var res = await graphClient.Teams[teamId].Channels[channelId].Members.GetAsync();
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
        /// Adds a member to the channel
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="channelId"></param>
        /// <param name="conversationMemberId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("AddChannelMember/{channelId}/{userPrincipal}")]
        public async Task<ResponseResult> AddChannelMember(string teamId, string channelId, string userPrincipal)
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

                var res = await graphClient.Teams[teamId].Channels[channelId].Members.PostAsync(requestBody);
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
        /// Removes a member from the channel
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="channelId"></param>
        /// <param name="conversationMemberId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("RemoveChannelMember/{channelId}/{conversationMemberId}")]
        public async Task<ResponseResult> RemoveChannelMember(string teamId, string channelId, string conversationMemberId)
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
                            "user@odata.bind" , $"https://graph.microsoft.com/v1.0/users('{conversationMemberId}')"
                        },
                    },
                };

                await graphClient.Teams[teamId].Channels[channelId].Members[conversationMemberId].DeleteAsync();
                result.Data = "Member removed successgully";
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
        /// Sends a post in the channel
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="channelId"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SendMessage/{channelId}/From/{userId}/{userName}")]
        public async Task<ResponseResult> SendMessage(string teamId,string channelId,string userId,string userName, [FromBody] string message) {
            try
            {
                //var requestBody = new ChatMessage
                //{
                //    Body = new ItemBody
                //    {
                //        Content = message,
                //    },
                //};

                //var res = await graphClient.Teams[teamId].Channels[channelId].Messages.PostAsync(requestBody);

                var requestBody = new ChatMessage
                {
                    CreatedDateTime = DateTimeOffset.Parse("2019-02-04T19:58:15.511Z"),
                    From = new ChatMessageFromIdentitySet
                    {
                        User = new Identity
                        {
                            Id = userId,
                            DisplayName = userName,
                            AdditionalData = new Dictionary<string, object>
                            {
                                {
                                    "userIdentityType" , "aadUser"
                                },
                            },
                        },
                    },
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = message,
                    },
                };

                // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
                var res = await graphClient.Teams[teamId].Channels[channelId].Messages.PostAsync(requestBody);


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
        /// Gets list of the posts sent in the channel
        /// </summary>
        /// <param name="teamId"></param>
        /// <param name="channelId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetMessages/{channelId}")]
        public async Task<ResponseResult> GetMessages(string teamId, string channelId)
        {
            try
            {
                var res = await graphClient.Teams[teamId].Channels[channelId].Messages.GetAsync();
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
