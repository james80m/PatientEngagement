using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Dialogs;
using System.Web.Http.Description;
using System.Net.Http;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using PatientEngagementBot.Dynamics;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using Microsoft.Bot.Builder.Dialogs.Internals;
using System.Threading;
using Newtonsoft.Json;
using static PatientEngagementBot.Dynamics.Dynamics;

namespace Microsoft.Bot.Sample.PatientEngagementBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// receive a message from a user and send replies
        /// </summary>
        /// <param name="activity"></param>
        [ResponseType(typeof(void))]
        public virtual async Task<HttpResponseMessage> Post([FromBody] Activity activity)
        {
            
            // check if activity is of type message
            if (activity != null && activity.GetActivityType() == ActivityTypes.Message)
            {
                await Conversation.SendAsync(activity, () => new EchoDialog());
            }

            if(activity.GetActivityType() == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
                IConversationUpdateActivity iConversationUpdated = activity as IConversationUpdateActivity;
                ConnectorClient connector = new ConnectorClient(new System.Uri(activity.ServiceUrl));
                if (iConversationUpdated != null)
                {
                    if ((string)activity.Text == "Address" || (activity.From.Name == "User" && iConversationUpdated.MembersAdded.Where(m => m.Id == iConversationUpdated.Recipient.Id).Count() > 0))
                    {
                        //Initialize the CRM connection
                        IOrganizationService service = Dynamics.GetService();

                        //Extract the direct line address (if applicable)
                        CXAddress address = JsonConvert.DeserializeObject<CXAddress>(activity.Value != null ? (string)activity.Value : "{}");
                        if (address.Address == null)
                        {
                            address.Address = "3145783471";
                            address.Message = "Initialize Conversation";
                            address.Bot = "Default Bot";
                        }

                        

                        //get bot
                        EntityCollection botResults = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' no-lock='true'>
                                        <entity name='aspect_cxbot'>
                                        <attribute name='aspect_cxbotid' />
                                        <attribute name='aspect_name' />
                                        <attribute name='createdon' />
                                        <order attribute='aspect_name' descending='false' />
                                        <filter type='and'>
                                            <filter type='or'>
                                            <condition attribute='aspect_name' operator='eq' value='{0}' />
                                            <condition attribute='aspect_default' operator='eq' value='1' />
                                            </filter>
                                        </filter>
                                        </entity>
                                    </fetch>" , address.Bot == null ? string.Empty : address.Bot)));

                        if(botResults.Entities.Count() == 0)
                        {
                            var responseMessage = activity.CreateReply();
                            responseMessage.Text = "There is currently no default CX bot configured in your environment or the bot speicified does not exist.";
                            responseMessage.Speak = "There is currently no default CX bot configured in your environment or the bot speicified does not exist.";
                            ResourceResponse msgResponse = connector.Conversations.ReplyToActivity(responseMessage);
                            return new HttpResponseMessage(System.Net.HttpStatusCode.Accepted);
                        }

                        Xrm.Sdk.Entity cxBot = botResults.Entities[0];
                        var namedMatchedBot = (from b in botResults.Entities where (string)b["aspect_name"] == address.Address select b);
                        if (namedMatchedBot.Count() > 0) cxBot = namedMatchedBot.First();
                            


                        //get first step
                        Xrm.Sdk.Entity cxCurrentStep = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                        <entity name='aspect_cxstep'>
                                        <all-attributes />
                                        <link-entity name='aspect_cxbot' to='aspect_cxbotid' from='aspect_cxbotid' alias='aspect_cxbot'>
                                            <all-attributes />
                                        </link-entity>
                                        <order descending='false' attribute='aspect_name' />
                                        <filter type='and'>
                                            <condition value='1' attribute='aspect_root' operator='eq' />
                                            <condition value='{0}' attribute='aspect_cxbotid' operator='eq' />
                                        </filter>
                                        </entity>
                                    </fetch>", cxBot.Id))).Entities[0];


                        //create initial conversation
                        Xrm.Sdk.Entity cxConversation = new Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", activity.Conversation.Id);
                        cxConversation["aspect_name"] = activity.Conversation.Id;
                        cxConversation["aspect_from"] = address.Address;
                        cxConversation["aspect_lastanswer"] = address.Message;
                        cxConversation["aspect_cxbotid"] = new EntityReference(cxBot.LogicalName, cxBot.Id);
                        cxConversation["aspect_currentcxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
                        UpsertResponse cxConversationResponse = (UpsertResponse)service.Execute(new UpsertRequest()
                        {
                            Target = cxConversation
                        });

                        if(!string.IsNullOrEmpty(address.Message))
                        {
                            //Create a conversation message with the initial
                            Xrm.Sdk.Entity conversationClient = new Xrm.Sdk.Entity("aspect_cxconversationmessage");
                            conversationClient["aspect_cxconversationid"] = new EntityReference("aspect_cxconversation", "aspect_conversationid", activity.Conversation.Id);
                            conversationClient["aspect_cxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
                            conversationClient["aspect_direction"] = false;
                            conversationClient["aspect_name"] = activity.Conversation.Id;
                            service.Create(conversationClient);
                        }


                        //ROUTE TO NEXT QUESTION
                        while (cxCurrentStep != null)
                        {
                            //FIRE ENTRY SEARCHES
                            Dynamics.ExecuteSearch(service, (EntityReference)cxConversationResponse.Results["Target"], new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id), true);

                            //FIRE ENTRY WORKFLOW
                            Dynamics.FireWorkflow(service, cxCurrentStep, false);


                            switch (Dynamics.CXGetType((OptionSetValue)cxCurrentStep["aspect_type"]))
                            {
                                case "MESSAGE":
                                    var cxNextStep = await Dynamics.PromptInitialQuestion(activity
                                        , connector
                                        , service
                                        , new Xrm.Sdk.Entity(((EntityReference)cxConversationResponse.Results["Target"]).LogicalName) {
                                            Id = ((EntityReference)cxConversationResponse.Results["Target"]).Id
                                        }
                                        , cxCurrentStep);
                                    if (cxNextStep != null)
                                    {
                                        cxCurrentStep = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                                            <entity name='aspect_cxstep'>
                                                            <all-attributes />
                                                            <link-entity name='aspect_cxbot' to='aspect_cxbotid' from='aspect_cxbotid' alias='aspect_cxbot'>
                                                                <all-attributes />
                                                            </link-entity>
                                                            <filter type='and'>
                                                                <condition value='{0}' attribute='aspect_cxstepid' operator='eq' />
                                                            </filter>
                                                            </entity>
                                                        </fetch>", cxNextStep.Id))).Entities[0];
                                    }
                                    break;
                                case "QUESTION": 
                                case "MENU":
                                case "RECORD":
                                case "TRANSFER":
                                    if (cxCurrentStep != null)
                                    {
                                        await Dynamics.PromptInitialQuestion(activity
                                            , connector
                                            , service
                                            , new Xrm.Sdk.Entity(((EntityReference)cxConversationResponse.Results["Target"]).LogicalName)
                                                {
                                                    Id = ((EntityReference)cxConversationResponse.Results["Target"]).Id
                                                }
                                            , cxCurrentStep);
                                        cxCurrentStep = null;
                                    }
                                    break;
                            }
                        }
                    }
                }

                await HandleSystemMessage(activity);
            }
            else
            {
                await HandleSystemMessage(activity);
            }

            return new HttpResponseMessage(System.Net.HttpStatusCode.Accepted);
        }

        private async Task<Activity> HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
               
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {

            }
            else if (message.Type == ActivityTypes.Invoke)
            {
                
            }

            return null;
        }


        
    }
}