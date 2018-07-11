using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Dialogs;
using System.Net.Http;
using Microsoft.Xrm.Sdk;
using PatientEngagementBot.Dynamics;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Autofac;
using System.Threading;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Newtonsoft.Json;
using Microsoft.Crm.Sdk.Messages;
using System.Web;

namespace Microsoft.Bot.Sample.PatientEngagementBot
{
    [Serializable]
    public class EchoDialog : IDialog<object>
    {
        protected int count = 1;

        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument)
        {
            var message = await argument;


            IOrganizationService service = Dynamics.GetService();


            //Get the current and next step
            Xrm.Sdk.Entity cxCurrentStep = Dynamics.GetCXCurrentStepEntityByConversationId(service, message.Conversation.Id);
            EntityReference cxNextStep = cxCurrentStep.Contains("aspect_nextcxstepid") ? (EntityReference)cxCurrentStep["aspect_nextcxstepid"] : null; 


            //Create a conversation message with the users response
            Xrm.Sdk.Entity conversationClient = new Xrm.Sdk.Entity("aspect_cxconversationmessage");
            conversationClient["aspect_cxconversationid"] = new EntityReference("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
            conversationClient["aspect_cxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
            conversationClient["aspect_direction"] = false;
            conversationClient["aspect_name"] = message.Id;
            conversationClient["aspect_message"] = message.Text;
            service.Create(conversationClient);


            //TODO: Check global utterances
            Xrm.Sdk.Entity globalUtterance = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                                      <entity name='aspect_cxglobalutterance'>
                                                        <attribute name='aspect_cxglobalutteranceid' />
                                                        <attribute name='aspect_cxstepid' />
                                                        <attribute name='aspect_answers' />
                                                        <order descending='false' attribute='aspect_name' />
                                                        <filter type='and'>
                                                          <condition value='%{0}%' attribute='aspect_answers' operator='like' />
                                                        </filter>
                                                      </entity>
                                                    </fetch>", HttpUtility.UrlEncode(message.Text)))).Entities.FirstOrDefault();
            

            if(globalUtterance != null)
            {
                cxNextStep = (EntityReference)globalUtterance["aspect_cxstepid"];
            }
            else
            {
                //validate user input
                string utteranceMatch = UtteranceMatch(message.Text, (string)(cxCurrentStep.Contains("aspect_answers") ? (string)cxCurrentStep["aspect_answers"] : null));
                if (Dynamics.CXGetType((OptionSetValue)cxCurrentStep["aspect_type"]) == "MENU")
                {
                    Xrm.Sdk.EntityCollection cxStepAnswers = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                                    <entity name='aspect_cxstepanswer'>
                                                    <all-attributes />
                                                    <order descending='false' attribute='aspect_name' />
                                                    <filter type='and'>
                                                        <condition value='{0}' attribute='aspect_cxstepid' operator='eq' />
                                                    </filter>
                                                    </entity>
                                                </fetch>", cxCurrentStep.Id)));

                    foreach (Xrm.Sdk.Entity e in cxStepAnswers.Entities)
                    {
                        utteranceMatch = UtteranceMatch(message.Text, e.Contains("aspect_answers") ? (string)e["aspect_answers"] : null);
                        if (!string.IsNullOrEmpty(utteranceMatch))
                        {
                            cxNextStep = e.Contains("aspect_nextcxstepid") ? (EntityReference)e["aspect_nextcxstepid"] : null;
                            break;
                        }
                    }
                }

                if (utteranceMatch == null)
                {
                    string noMatchMessage = string.Format("I'm sorry, I didn't understand '{0}'.", message.Text); //TODO make this message come from CX Bot
                    Xrm.Sdk.Entity noMatchResponse = new Xrm.Sdk.Entity("aspect_cxconversationmessage");
                    conversationClient["aspect_cxconversationid"] = new EntityReference("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
                    conversationClient["aspect_cxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
                    conversationClient["aspect_direction"] = true;
                    conversationClient["aspect_name"] = message.Id;
                    conversationClient["aspect_message"] = noMatchMessage;
                    service.Create(conversationClient);

                    IMessageActivity responseNoMatchMessage = context.MakeMessage();
                    responseNoMatchMessage.Text = string.Format(noMatchMessage);
                    responseNoMatchMessage.Value = message.Value;
                    await context.PostAsync(responseNoMatchMessage);

                    context.Wait(MessageReceivedAsync);
                    return;
                }
            }

            
            //Update the last answer in the conversation
            Microsoft.Xrm.Sdk.Entity upsConversation2 = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
            upsConversation2["aspect_lastanswer"] = message.Text;
            UpsertResponse cxConversationResponse = (UpsertResponse)service.Execute(new UpsertRequest()
            {
                Target = upsConversation2
            });


            int maxLoopCount = 10;
            int currentLoopCount = 0;

            
            //EXECUTE EXIT SEARCH
            Dynamics.ExecuteSearch(service, (EntityReference)cxConversationResponse.Results["Target"], new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id), false);

            //FIRE EXIT WORKFLOW
            Dynamics.FireWorkflow(service, cxCurrentStep, false);

            //REFRESH THE NEXT CX STEP IN CASE A SEARCH/WORKFLOW ALTERED IT
            Xrm.Sdk.Entity rerouteCXStep = Dynamics.GetCXNextStepEntityByConversationId(service, message.Conversation.Id, cxNextStep);

            if(rerouteCXStep != null)
            {
                cxNextStep = new EntityReference(rerouteCXStep.LogicalName, rerouteCXStep.Id);
            }

            while (cxNextStep != null && currentLoopCount <= maxLoopCount)
            {
                if (currentLoopCount > 0)
                {
                    //EXECUTE EXIT SEARCH
                    Dynamics.ExecuteSearch(service, (EntityReference)cxConversationResponse.Results["Target"], new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id), false);

                    //FIRE EXIT WORKFLOW
                    Dynamics.FireWorkflow(service, cxCurrentStep, false);

                    //REFRESH THE NEXT CX STEP IN CASE A SEARCH/WORKFLOW ALTERED IT
                    rerouteCXStep = Dynamics.GetCXNextStepEntityByConversationId(service, message.Conversation.Id, cxNextStep);
                }

                //EXIT IF THERE IS NO NEXT STEP
                if (cxNextStep == null) return;

                //FIRE ENTRY WORKFLOW
                Dynamics.FireWorkflow(service, rerouteCXStep, true);

                //EXECUTE ENTRY SEARCH
                Dynamics.ExecuteSearch(service, (EntityReference)cxConversationResponse.Results["Target"], new EntityReference(rerouteCXStep.LogicalName, rerouteCXStep.Id), true);


                //process question
                switch (Dynamics.CXGetType((OptionSetValue)rerouteCXStep["aspect_type"]))
                {
                    case "MESSAGE":
                        await Dynamics.PromptNextQuestion(context, service, cxCurrentStep, rerouteCXStep, message);
                        cxCurrentStep = Dynamics.GetCXCurrentStepEntityByConversationId(service, message.Conversation.Id);
                        cxNextStep = cxCurrentStep.Contains("aspect_nextcxstepid") ? (EntityReference)cxCurrentStep["aspect_nextcxstepid"] : null;
                        if(cxNextStep == null)
                        {
                            //if the path ended with a search but no default next step execute the exit workflows
                            //EXECUTE EXIT SEARCH
                            Dynamics.ExecuteSearch(service, (EntityReference)cxConversationResponse.Results["Target"], new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id), false);
                            //REFRESH THE NEXT CX STEP IN CASE A SEARCH/WORKFLOW ALTERED IT
                            rerouteCXStep = Dynamics.GetCXNextStepEntityByConversationId(service, message.Conversation.Id, cxNextStep);
                            if (rerouteCXStep != null && rerouteCXStep.Id != cxCurrentStep.Id)
                            {
                                cxNextStep = new EntityReference(rerouteCXStep.LogicalName, rerouteCXStep.Id);
                            }
                            
                        }
                        currentLoopCount++;
                        break;
                    case "QUESTION":
                    case "MENU":
                    case "RECORD":
                    case "TRANSFER":
                        await Dynamics.PromptNextQuestion(context, service, cxCurrentStep, rerouteCXStep, message);
                        cxNextStep = null;
                        break;
                }
            }

            context.Wait(MessageReceivedAsync);
        }


        private string UtteranceMatch(string message, string answerList)
        {
            List<string> answers = !string.IsNullOrEmpty(answerList) && !string.IsNullOrWhiteSpace(answerList) ? answerList.Split(',').ToList() : new List<string>(new string[] { message });

            string utteranceMatch = (from a in answers where a.ToLowerInvariant().Trim() == message.ToLowerInvariant().Trim() select a).FirstOrDefault();
            if (utteranceMatch == null)
            {
                foreach (string s in answers)
                {
                    if (s.Trim().Length > 0 && message.ToLowerInvariant().Trim().Contains(s.ToLowerInvariant().Trim()))
                    {
                        utteranceMatch = s;
                        break;
                    }
                }
            }

            return utteranceMatch;
        }

    }
}