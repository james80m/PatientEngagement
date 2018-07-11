using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace PatientEngagementBot.Dynamics
{
    public static class Dynamics
    {
        public static IOrganizationService GetService()
        {
            try
            {
                var connString = "Url=https://xxxxxx.crm.dynamics.com; Username=user@xxxxxx.onmicrosoft.com; Password=xxxxxx; authtype=Office365";
                CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connString);
                IOrganizationService service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                var rtn = service.Execute(new WhoAmIRequest());
                return service;
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to connect to Dynamics");
            }
        }


        /// <summary>
        /// Returns the current step along with conversation data
        /// </summary>
        /// <param name="service"></param>
        /// <param name="conversationId"></param>
        /// <returns></returns>
        public static Microsoft.Xrm.Sdk.Entity GetCXCurrentStepEntityByConversationId(IOrganizationService service, string conversationId)
        {
            return service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='true' output-format='xml-platform' no-lock='true'>
                                                    <entity name='aspect_cxstep'>
                                                    <all-attributes />
                                                    <order descending='false' attribute='aspect_name' />
                                                    <link-entity name='aspect_cxconversation' to='aspect_cxstepid' from='aspect_currentcxstepid' alias='aspect_cxconversation'>
                                                        <all-attributes />
                                                        <filter type='and'>
                                                            <condition value='{0}' attribute='aspect_conversationid' operator='eq' />
                                                        </filter>
                                                    </link-entity>
                                                    <link-entity name='aspect_cxbot' to='aspect_cxbotid' from='aspect_cxbotid' alias='aspect_cxbot'>
                                                        <all-attributes />
                                                    </link-entity>
                                                    </entity>
                                                </fetch>", conversationId))).Entities.FirstOrDefault();
        }


        /// <summary>
        /// Returns the next step along with conversation data
        /// </summary>
        /// <param name="service"></param>
        /// <param name="conversationId"></param>
        /// <returns></returns>
        public static Microsoft.Xrm.Sdk.Entity GetCXNextStepEntityByConversationId(IOrganizationService service, string conversationId, EntityReference defaultNextCXStep)
        {
            //if a reroute next step has been populated use that
            var results = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='true' output-format='xml-platform' no-lock='true'>
                                                    <entity name='aspect_cxstep'>
                                                    <all-attributes />
                                                    <order descending='false' attribute='aspect_name' />
                                                    <link-entity name='aspect_cxconversation' to='aspect_cxstepid' from='aspect_nextcxstepid' alias='aspect_cxconversation'>
                                                        <all-attributes />
                                                        <filter type='and'>
                                                            <condition value='{0}' attribute='aspect_conversationid' operator='eq' />
                                                        </filter>
                                                    </link-entity>
                                                    <link-entity name='aspect_cxbot' to='aspect_cxbotid' from='aspect_cxbotid' alias='aspect_cxbot'>
                                                        <all-attributes />
                                                    </link-entity>
                                                    </entity>
                                                </fetch>", conversationId))).Entities.FirstOrDefault();

            if(results == null && defaultNextCXStep != null)
            {
                return service.Retrieve(defaultNextCXStep.LogicalName, defaultNextCXStep.Id, new ColumnSet(true));
            }

            return results;
        }

        /// <summary>
        /// Returns a speicified step with bot details
        /// </summary>
        /// <param name="service"></param>
        /// <param name="stepId"></param>
        /// <returns></returns>
        public static Microsoft.Xrm.Sdk.Entity GetCXStepEntity(IOrganizationService service, Guid stepId)
        {
            return service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='true' output-format='xml-platform' no-lock='true'>
                                                    <entity name='aspect_cxstep'>
                                                    <all-attributes />
                                                    <filter type='and'>
                                                        <condition value='{0}' attribute='aspect_cxstepid' operator='eq' />
                                                    </filter>
                                                    <link-entity name='aspect_cxbot' to='aspect_cxbotid' from='aspect_cxbotid' alias='aspect_cxbot'>
                                                        <all-attributes />
                                                    </link-entity>
                                                    </entity>
                                                </fetch>", stepId))).Entities[0];
        }

        public static async Task<Microsoft.Xrm.Sdk.Entity> PromptNextQuestion(IDialogContext context, IOrganizationService service, Microsoft.Xrm.Sdk.Entity cxCurrentStep, Microsoft.Xrm.Sdk.Entity cxNextStep, IMessageActivity message)
        {
            Microsoft.Xrm.Sdk.Entity cxStep = service.Retrieve(cxNextStep.LogicalName, cxNextStep.Id, new ColumnSet(true));

            //update conversation to initiate any workflow triggers
            Microsoft.Xrm.Sdk.Entity upsConversation = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
            upsConversation["aspect_nextcxstepid"] = new EntityReference(cxNextStep.LogicalName, cxNextStep.Id);
            UpsertResponse upsResponse = (UpsertResponse)service.Execute(new UpsertRequest()
            {
                Target = upsConversation
            });

            string responseText = ReplaceOutputText(service, (EntityReference)upsResponse["Target"], (string)cxStep["aspect_message"]);

            //Create a conversation message with outbound message
            Microsoft.Xrm.Sdk.Entity conversationMessageOutbound = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversationmessage");
            conversationMessageOutbound["aspect_cxconversationid"] = new EntityReference("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
            conversationMessageOutbound["aspect_cxstepid"] = new EntityReference(cxNextStep.LogicalName, cxNextStep.Id);
            conversationMessageOutbound["aspect_direction"] = true;
            conversationMessageOutbound["aspect_name"] = message.Id;
            conversationMessageOutbound["aspect_message"] = responseText;
            service.Create(conversationMessageOutbound);

            //Update conversation to reflect the current step
            Microsoft.Xrm.Sdk.Entity upsConversation2 = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", message.Conversation.Id);
            upsConversation2["aspect_currentcxstepid"] = new EntityReference(cxNextStep.LogicalName, cxNextStep.Id);
            upsConversation2["aspect_lastanswer"] = message.Text;
            upsConversation2["aspect_nextcxstepid"] = null;
            service.Execute(new UpsertRequest()
            {
                Target = upsConversation2
            });

            //send outbound message
            IMessageActivity responseMessage = context.MakeMessage();
            responseMessage.Text = responseText;
            responseMessage.Speak = responseText;
            responseMessage.InputHint = Dynamics.CXGetAnswers(service, cxStep);
            responseMessage.Value = JsonConvert.SerializeObject(new Dynamics.CXInformation()
            {
                //CXBotId = ((Guid)(((AliasedValue)["aspect_cxbot.aspect_cxbotid"]).Value)).ToString(),
                CXBotName = (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_name"]).Value),
                AudioDirectory = cxCurrentStep.Contains("aspect_cxbot.aspect_audiodirectory") ? (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_audiodirectory"]).Value) : string.Empty,
                RecordingDirectory = cxCurrentStep.Contains("aspect_cxbot.aspect_recordingdirectory") ? (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_recordingdirectory"]).Value) : string.Empty,
                CXStepId = cxStep.Id.ToString(),
                CXStepAudio = cxStep.Contains("aspect_audio") ? (string)cxStep["aspect_audio"] : string.Empty,
                CXText = responseText,
                CXAnswers = CXGetAnswers(service, cxStep),
                CXType = Dynamics.CXGetType((OptionSetValue)cxStep["aspect_type"])
            });
            await context.PostAsync(responseMessage);
            return cxNextStep;
        }

        public static async Task<EntityReference> PromptInitialQuestion(Activity context, ConnectorClient connector,  IOrganizationService service, Microsoft.Xrm.Sdk.Entity cxConversation, Microsoft.Xrm.Sdk.Entity cxCurrentStep)
        {
            Microsoft.Xrm.Sdk.Entity cxStep = cxCurrentStep;

            //update conversation to reflect the current step and initiate any workflow triggers
            Microsoft.Xrm.Sdk.Entity upsConversation1 = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", context.Conversation.Id);
            upsConversation1["aspect_currentcxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
            upsConversation1["aspect_nextcxstepid"] = null;
            service.Execute(new UpsertRequest()
            {
                Target = upsConversation1
            });

            string responseText = ReplaceOutputText(service, new EntityReference(cxConversation.LogicalName, cxConversation.Id), (string)cxStep["aspect_message"]); 


            //Create a conversation message with the text sent to the user
            Microsoft.Xrm.Sdk.Entity conversationClient = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversationmessage");
            conversationClient["aspect_cxconversationid"] = new EntityReference("aspect_cxconversation", cxConversation.Id);
            conversationClient["aspect_cxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
            conversationClient["aspect_direction"] = true;
            conversationClient["aspect_name"] = cxConversation.Id.ToString();
            conversationClient["aspect_message"] = responseText;
            service.Create(conversationClient);


            //send reply
            var responseMessage = context.CreateReply();
            responseMessage.Text = responseText;
            responseMessage.Speak = responseText;
            responseMessage.InputHint = Dynamics.CXGetAnswers(service, cxStep);
            responseMessage.Value = JsonConvert.SerializeObject(new Dynamics.CXInformation()
            {
                //CXBotId = ((Guid)(((AliasedValue)["aspect_cxbot.aspect_cxbotid"]).Value)).ToString(),
                CXBotName = (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_name"]).Value),
                AudioDirectory = cxCurrentStep.Contains("aspect_cxbot.aspect_audiodirectory") ? (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_audiodirectory"]).Value) : string.Empty,
                RecordingDirectory = cxCurrentStep.Contains("aspect_cxbot.aspect_recordingdirectory") ? (string)(((AliasedValue)cxCurrentStep["aspect_cxbot.aspect_recordingdirectory"]).Value) : string.Empty,
                CXStepId = cxStep.Id.ToString(),
                CXStepAudio = cxStep.Contains("aspect_audio") ? (string)cxStep["aspect_audio"] : string.Empty,
                CXText = responseText,
                CXAnswers = CXGetAnswers(service, cxStep),
                CXType = Dynamics.CXGetType((OptionSetValue)cxStep["aspect_type"])
            });
            ResourceResponse msgResponse = connector.Conversations.ReplyToActivity(responseMessage);

            //update conversation to reflect the current step and initiate any workflow triggers
            Microsoft.Xrm.Sdk.Entity upsConversation2 = new Microsoft.Xrm.Sdk.Entity("aspect_cxconversation", "aspect_conversationid", context.Conversation.Id);
            upsConversation2["aspect_currentcxstepid"] = new EntityReference(cxCurrentStep.LogicalName, cxCurrentStep.Id);
            upsConversation2["aspect_nextcxstepid"] = cxStep.Contains("aspect_nextcxstepid") ? (EntityReference)cxStep["aspect_nextcxstepid"] : null;
            service.Execute(new UpsertRequest()
            {
                Target = upsConversation2
            });

            return cxStep.Contains("aspect_nextcxstepid") ? (EntityReference)cxStep["aspect_nextcxstepid"] : null;
        }

        public static void ExecuteSearch(IOrganizationService service, Microsoft.Xrm.Sdk.EntityReference cxConversation, Microsoft.Xrm.Sdk.EntityReference cxStep, bool entrySearch)
        {
            Microsoft.Xrm.Sdk.Entity upsConversation = new Microsoft.Xrm.Sdk.Entity(cxConversation.LogicalName, cxConversation.Id);
            EntityCollection cxSearches = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                                              <entity name='aspect_cxsearch'>
                                                                <all-attributes />
                                                                <filter type='and'>
                                                                  <condition value='{0}' attribute='aspect_cxstepid' operator='eq' />
                                                                  <condition value='{1}' attribute='aspect_searchpoint' operator='eq' />
                                                                </filter>
                                                              </entity>
                                                            </fetch>"
                                                        , cxStep.Id
                                                        , entrySearch ? "0" : "1")));
            foreach (Microsoft.Xrm.Sdk.Entity search in cxSearches.Entities)
            {
                EntityCollection cxStepSearchEntities = service.RetrieveMultiple(new FetchExpression(ReplaceOutputText(service, cxConversation, (string)search["aspect_fetchxml"])));
                foreach (Microsoft.Xrm.Sdk.Entity cxStepSearchEntity in cxStepSearchEntities.Entities)
                {
                    switch (cxStepSearchEntity.LogicalName)
                    {
                        case "account":
                            upsConversation["aspect_accountid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "contact":
                            upsConversation["aspect_contactid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "incident":
                            upsConversation["aspect_caseid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "lead":
                            upsConversation["aspect_leadid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "opportunity":
                            upsConversation["aspect_opportunityid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "salesorder":
                            upsConversation["aspect_salesorderid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "invoice":
                            upsConversation["aspect_invoiceid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "appointment":
                            upsConversation["aspect_appointmentid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "task":
                            upsConversation["aspect_taskid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;


                        case "msemr_appointmentemr":
                            upsConversation["aspect_appointmentemrid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_careplan":
                            upsConversation["aspect_careplanid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_careplanactivity":
                            upsConversation["aspect_careplanactivityid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_careteam":
                            upsConversation["aspect_careteamid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_communication":
                            upsConversation["aspect_communicationid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_device":
                            upsConversation["aspect_deviceid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_encounter":
                            upsConversation["aspect_encounterid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_episodeofcarehistory":
                            upsConversation["aspect_episodeofcarehistoryid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_medicalidentifier":
                            upsConversation["aspect_medicalidentifierid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_medication":
                            upsConversation["aspect_medicationid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_observation":
                            upsConversation["aspect_observationid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_procedure":
                            upsConversation["aspect_procedureid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;
                        case "msemr_referralrequest":
                            upsConversation["aspect_referralrequestid"] = new EntityReference(cxStepSearchEntity.LogicalName, cxStepSearchEntity.Id);
                            break;

                    }
                }


                if (cxStepSearchEntities.Entities.Count == 0 && search.Contains("aspect_nomatchcxstepid"))
                {
                    upsConversation["aspect_nextcxstepid"] = (EntityReference)search["aspect_nomatchcxstepid"];
                }
                else if (cxStepSearchEntities.Entities.Count == 1 && search.Contains("aspect_singlematchcxstepid"))
                {
                    upsConversation["aspect_nextcxstepid"] = (EntityReference)search["aspect_singlematchcxstepid"];
                }
                else if (cxStepSearchEntities.Entities.Count > 1 && search.Contains("aspect_multiplematchcxstepid"))
                {
                    upsConversation["aspect_nextcxstepid"] = (EntityReference)search["aspect_multiplematchcxstepid"];
                }



                if (upsConversation.Attributes.Count > 0)
                {
                    service.Update(upsConversation);
                }

                if (cxStepSearchEntities.Entities.Count == 0 && search.Contains("aspect_nomatchworkflowid"))
                {
                    try
                    {
                        ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                        {
                            WorkflowId = ((EntityReference)search["aspect_nomatchworkflowid"]).Id,
                            EntityId = cxConversation.Id
                        };
                        // Execute the workflow.
                        ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)service.Execute(request);
                    }
                    catch (Exception ex) { }
                }
                else if (cxStepSearchEntities.Entities.Count == 1 && search.Contains("aspect_singlematchworkflowid"))
                {
                    try
                    {
                        ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                        {
                            WorkflowId = ((EntityReference)search["aspect_singlematchworkflowid"]).Id,
                            EntityId = cxConversation.Id
                        };
                        // Execute the workflow.
                        ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)service.Execute(request);
                    }
                    catch (Exception ex) { }
                }
                else if (cxStepSearchEntities.Entities.Count > 1 && search.Contains("aspect_multiplematchworkflowid"))
                {
                    try
                    {
                        ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                        {
                            WorkflowId = ((EntityReference)search["aspect_multiplematchworkflowid"]).Id,
                            EntityId = cxConversation.Id
                        };
                        // Execute the workflow.
                        ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)service.Execute(request);
                    }
                    catch (Exception ex) { }
                }
            }
        }

        public static void FireWorkflow(IOrganizationService service, Microsoft.Xrm.Sdk.Entity cxStep, bool entryWorkflow)
        {
            try
            {
                string workflowProperty = entryWorkflow ? "aspect_entryworkflowid" : "aspect_answerworkflowid";
                if (cxStep.Contains(workflowProperty))
                {
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                    {
                        WorkflowId = ((EntityReference)cxStep[workflowProperty]).Id,
                        EntityId = (Guid)(((AliasedValue)cxStep["aspect_cxconversation.aspect_cxconversationid"]).Value)
                    };
                    // Execute the workflow.
                    ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)service.Execute(request);
                }
            }
            catch (Exception ex)
            {
                //user defined workflow errored
            }
        }

        public static string CXGetAnswers(IOrganizationService service, Microsoft.Xrm.Sdk.Entity cxStep)
        {
            switch (((OptionSetValue)cxStep["aspect_type"]).Value)
            {
                case 126460000: //MESSAGE
                    return string.Empty;
                    break;
                case 126460001: //QUESTION
                    return cxStep.Contains("aspect_answers") ? (string)cxStep["aspect_answers"] : string.Empty;
                    break;
                case 126460002: //MENU
                    EntityCollection cxStepAnswers = service.RetrieveMultiple(new FetchExpression(string.Format(@"<fetch mapping='logical' version='1.0' distinct='false' output-format='xml-platform' no-lock='true'>
                                                      <entity name='aspect_cxstepanswer'>
                                                        <all-attributes />
                                                        <order descending='false' attribute='aspect_name' />
                                                        <filter type='and'>
                                                          <condition value='{0}' attribute='aspect_cxstepid' operator='eq' />
                                                        </filter>
                                                      </entity>
                                                    </fetch>", cxStep.Id)));

                    List<string> allAnswers = new List<string>();
                    foreach (Microsoft.Xrm.Sdk.Entity e in cxStepAnswers.Entities)
                    {
                        allAnswers.AddRange(((string)e["aspect_answers"]).Split(','));
                    }
                    return string.Join(",", (from a in allAnswers where !string.IsNullOrEmpty(a) && !string.IsNullOrWhiteSpace(a) select a.Trim()));
                    break;
                case 126460003: //TRANSFER
                    return string.Empty;
                    break;
                default:
                    return string.Empty;
            }
        }

        public static string ReplaceOutputText(IOrganizationService service, Microsoft.Xrm.Sdk.EntityReference cxConversationId, string replacementString)
        {
            if (!replacementString.Contains("{") || !replacementString.Contains("}")) return ReplaceInlineFunctions(replacementString);

            Microsoft.Xrm.Sdk.Entity cxConversation = service.Retrieve(cxConversationId.LogicalName, cxConversationId.Id, new ColumnSet(true));
            foreach(var a in cxConversation.Attributes)
            {
                if (!replacementString.Contains("{") || !replacementString.Contains("}")) return ReplaceInlineFunctions(replacementString);
                if (a.Key == "owningbusinessunit" || a.Key == "modifiedonbehalfby" || a.Key == "createdby" || a.Key == "owningteam" || a.Key == "owninguser" || a.Key == "createdonbehalfby") continue;

                switch (a.Value.GetType().ToString())
                {
                    case "Microsoft.Xrm.Sdk.EntityReference":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((EntityReference)cxConversation[a.Key]).Id.ToString());

                        Microsoft.Xrm.Sdk.Entity referencedEntity = service.Retrieve(((EntityReference)cxConversation[a.Key]).LogicalName, ((EntityReference)cxConversation[a.Key]).Id, new ColumnSet(true));
                        foreach (var b in (from t in referencedEntity.Attributes where replacementString.Contains("{" + t.Key + "}") select t))
                        {
                            if (b.Key == "owningbusinessunit" || b.Key == "modifiedonbehalfby" || b.Key == "createdby" || b.Key == "owningteam" || b.Key == "owninguser" || b.Key == "createdonbehalfby") continue;

                            switch (b.Value.GetType().ToString())
                            {
                                case "Microsoft.Xrm.Sdk.EntityReference":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((EntityReference)referencedEntity[b.Key]).Name.ToString());
                                    break;
                                case "Microsoft.Xrm.Sdk.OptionSetValue":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((OptionSetValue)referencedEntity[b.Key]).ToString());
                                    break;
                                case "System.Boolean":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((bool)referencedEntity[b.Key]).ToString());
                                    break;
                                case "System.DateTime":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((DateTime)referencedEntity[b.Key]).ToString("M/dd/yyyy h:mm tt"));
                                    break;
                                case "System.Guid":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((Guid)referencedEntity[b.Key]).ToString());
                                    break;
                                case "System.Int32":
                                case "System.Int64":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", ((int)referencedEntity[b.Key]).ToString());
                                    break;
                                case "System.String":
                                    replacementString = replacementString.Replace("{" + b.Key + "}", (string)referencedEntity[b.Key]);
                                    break;
                            }
                        }

                        break;
                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((OptionSetValue)cxConversation[a.Key]).ToString());
                        break;
                    case "System.Boolean":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((bool)cxConversation[a.Key]).ToString());
                        break;
                    case "System.DateTime":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((DateTime)cxConversation[a.Key]).ToString("M/dd/yyyy h:mm tt"));
                        break;
                    case "System.Guid":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((Guid)cxConversation[a.Key]).ToString());
                        break;
                    case "System.Int32":
                    case "System.Int64":
                        replacementString = replacementString.Replace("{" + a.Key + "}", ((int)cxConversation[a.Key]).ToString());
                        break;
                    case "System.String":
                        replacementString = replacementString.Replace("{" + a.Key + "}", (string)cxConversation[a.Key]);
                        break;
                }
            }

            return replacementString;
        }
        
        public static string ReplaceInlineFunctions(string replacementString)
        {
            try
            {
                //TTS phone
                replacementString = Regex.Replace(replacementString, "TTSPhone\\(\"(.*?)\"\\)", delegate (Match match)
                {
                    string v = Regex.Replace(match.ToString(), "[^0-9.]", string.Empty);
                    var arr = v.ToCharArray();
                    return string.Join(" ", arr);
                });


                //TTS date and time
                replacementString = Regex.Replace(replacementString, "TTSDateTime\\(\"(.*?)\"\\)", delegate (Match match)
                {
                    string baseMatch = Regex.Replace(match.ToString(), "\\(\"(.*?)\"\\)", delegate (Match match2)
                    {
                        DateTime parsedDate = DateTime.Parse(match2.ToString().Trim('(').Trim(')').Trim('"'));

                        return string.Format("{0} {1} at {2}"
                            , parsedDate.ToString("dddd MMMM")
                            , (new string[]{ "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth",
                             "ninth", "tenth", "eleventh", "twelfth", "thirteenth", "fourteenth", "fifteenth", "sixteenth",
                             "seventeenth", "eighteenth", "nineteenth", "twentieth", "twenty first", "twenty second", "twenty third", "twenty fourth",
                             "twenty fifth", "twenty sixth", "twenty seventh", "twenty eighth", "twenty ninth", "thirtieth",
                             "thirty first"}[parsedDate.Day - 1])
                            , parsedDate.ToString("h:mm tt"));
                    });

                    return baseMatch.Replace("TTSDateTime", string.Empty);
                });

                //TTS date
                replacementString = Regex.Replace(replacementString, "TTSDate\\(\"(.*?)\"\\)", delegate (Match match)
                {
                    string baseMatch = Regex.Replace(match.ToString(), "\\(\"(.*?)\"\\)", delegate (Match match2)
                    {
                        DateTime parsedDate = DateTime.Parse(match2.ToString().Trim('(').Trim(')').Trim('"'));

                        return string.Format("{0} {1}"
                            , parsedDate.ToString("dddd MMMM")
                            , (new string[]{ "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth",
                             "ninth", "tenth", "eleventh", "twelfth", "thirteenth", "fourteenth", "fifteenth", "sixteenth",
                             "seventeenth", "eighteenth", "nineteenth", "twentieth", "twenty first", "twenty second", "twenty third", "twenty fourth",
                             "twenty fifth", "twenty sixth", "twenty seventh", "twenty eighth", "twenty ninth", "thirtieth",
                             "thirty first"}[parsedDate.Day - 1]));
                    });

                    return baseMatch.Replace("TTSDate", string.Empty);
                });

                //TTS time
                replacementString = Regex.Replace(replacementString, "TTSTime\\(\"(.*?)\"\\)", delegate (Match match)
                {
                    string baseMatch = Regex.Replace(match.ToString(), "\\(\"(.*?)\"\\)", delegate (Match match2)
                    {
                        DateTime parsedDate = DateTime.Parse(match2.ToString().Trim('(').Trim(')').Trim('"'));

                        return string.Format("{0}", parsedDate.ToString("h:mm tt"));
                    });

                    return baseMatch.Replace("TTSTime", string.Empty);
                });

                return replacementString;
            }
            catch(Exception ex) { return replacementString; }
        }

        public static string CXGetType(OptionSetValue optionsetValue)
        {
            switch(optionsetValue.Value)
            {
                case 126460000:
                    return "MESSAGE";
                    break;
                case 126460001:
                    return "QUESTION";
                    break;
                case 126460002:
                    return "MENU";
                    break;
                case 126460003:
                    return "RECORD";
                    break;
                case 126460004:
                    return "TRANSFER";
                    break;
                default:
                    return "MESSAGE";
            }
        }

        public class CXAddress
        {
            public string Address { get; set; }
            public string Bot { get; set; }
            public string Message { get; set; }
        }

        public class CXInformation
        {
            public string CXBotId { get; set; }
            public string CXBotName { get; set; }
            public string AudioDirectory { get; set; }
            public string RecordingDirectory { get; set; }
            public string CXStepAudio { get; set; }
            public string CXStepId { get; set; }
            public string CXText { get; set; }
            public string CXAnswers { get; set; }
            public string CXType { get; set; }
        }
    }
}