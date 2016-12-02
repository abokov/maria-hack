using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;

/// <summary>

using Microsoft.Bot.Builder.FormFlow;
using System.Collections.Generic;
using Microsoft.Bot.Builder.Dialogs;
using System.Diagnostics;
#pragma warning disable 649
/// </summary>


namespace MaryaBotApplication
{   /*     
    public enum ReportParam1
    {
        dat, val, gr, type_per
    };    
    */
    [Serializable]
    public class Params
    {
      //  public ReportParam1? Param1;
      
        [Prompt("Введите дату отчета:")]
        public string param_dat { get; set; }//Параметр дата
        [Prompt("Введите тип отчета:")]
        public string param_type_per { get; set; }//Тип отчета

        public static IForm<Params> BuildForm()
        {
            OnCompletionAsyncDelegate<Params> processOrder = async (context, state) =>
            {
                await context.PostAsync("We are currently processing your report. We will message you the status.");

                string st_url = "https://maria-function-gen.azurewebsites.net/api/HttpTriggerGen?code=fa3K7cUzvuAO1i95k0ICfXhhaKjD8MOzewMk77OefnQn1Uvp4SZCSw==&param_type_per=" + state.param_type_per;
                HttpWebRequest request_get = (HttpWebRequest)HttpWebRequest.Create(st_url);
                request_get.Method = "GET";

                HttpWebResponse response_get = (HttpWebResponse)request_get.GetResponse();

                await context.PostAsync(response_get.StatusCode.ToString());

            };

            return new FormBuilder<Params>()
                    .Message("Welcome to the simple marya bot!").OnCompletion(processOrder)
                    .Build();
        }
    };



    [BotAuthentication]
    public class MessagesController : ApiController
    {
        internal static IDialog<Params> MakeRootDialog()
        {
            return Chain.From(() => FormDialog.FromForm(Params.BuildForm));
        }
        [ResponseType(typeof(void))]

        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
     public virtual async Task<HttpResponseMessage> Post([FromBody] Activity activity)
     {
            if (activity != null)
            {
                // one of these will have an interface and process it
                switch (activity.GetActivityType())
                {
                    case ActivityTypes.Message:
                        await Conversation.SendAsync(activity, MakeRootDialog);
                        break;
                    case ActivityTypes.ConversationUpdate:
                    case ActivityTypes.ContactRelationUpdate:
                    case ActivityTypes.Typing:
                    case ActivityTypes.DeleteUserData:
                    default:
                        Trace.TraceError($"Unknown activity type ignored: {activity.GetActivityType()}");
                        break;
                }

            }

            

            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        private Activity HandleSystemMessage(Activity message)
                            {
                                if (message.Type == ActivityTypes.DeleteUserData)
                                {
                                    // Implement user deletion here
                                    // If we handle user deletion, return a real message
                                }
                                else if (message.Type == ActivityTypes.ConversationUpdate)
                                {
                                    // Handle conversation state changes, like members being added and removed
                                    // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                                    // Not available in all channels
                                }
                                else if (message.Type == ActivityTypes.ContactRelationUpdate)
                                {
                                    // Handle add/remove from contact lists
                                    // Activity.From + Activity.Action represent what happened
                                }
                                else if (message.Type == ActivityTypes.Typing)
                                {
                                    // Handle knowing tha the user is typing
                                }
                                else if (message.Type == ActivityTypes.Ping)
                                {
                                }

                                return null;
                            }
    }
    
}