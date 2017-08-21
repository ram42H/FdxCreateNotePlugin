using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using System.Collections;
using System.ServiceModel;

namespace FdxCreateNotePlugin
{
    public class CopyNoteOnLeadQualify : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("LeadId") && context.InputParameters["LeadId"] is EntityReference)
            {
                #region setup plugin steps
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                EntityReference lead = (EntityReference)context.InputParameters["LeadId"];

                int step = 1;
                if (lead.LogicalName != "lead")
                    return;

                #endregion

                #region Business Code
                step = 2;
                try
                {
                    #region Fetch Annotation of a Lead
                    step = 3;
                    QueryExpression queryNote = new QueryExpression();
                    queryNote.EntityName = "annotation";
                    queryNote.ColumnSet = new ColumnSet(true);
                    queryNote.Distinct = false;
                    queryNote.Criteria = new FilterExpression();
                    queryNote.Criteria.FilterOperator = LogicalOperator.And;
                    step = 4;
                    queryNote.Criteria.AddCondition("objectid", ConditionOperator.Equal, lead.Id);
                    queryNote.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, lead.LogicalName);

                    step = 5;
                    EntityCollection notes = service.RetrieveMultiple(queryNote);

                    #endregion

                    #region Process all Annotations

                    if (notes.Entities.Count > 0)
                    {
                        step = 6;
                        foreach (Entity singleNote in notes.Entities)
                        {
                            #region Copy Annotation details into a new Object
                            step = 7;
                            Entity newNote = new Entity("annotation");
                            if (singleNote.Attributes.ContainsKey("documentbody"))
                            {
                                newNote.Attributes["documentbody"] = singleNote.Attributes["documentbody"];
                            }
                            if (singleNote.Attributes.ContainsKey("filename"))
                            {
                                newNote.Attributes["filename"] = singleNote.Attributes["filename"];
                            }
                            if (singleNote.Attributes.ContainsKey("filename"))
                            {
                                newNote.Attributes["filename"] = singleNote.Attributes["filename"];
                            }
                            if (singleNote.Attributes.ContainsKey("isdocument"))
                            {
                                newNote.Attributes["isdocument"] = singleNote.Attributes["isdocument"];
                            }
                            if (singleNote.Attributes.ContainsKey("langid"))
                            {
                                newNote.Attributes["langid"] = singleNote.Attributes["langid"];
                            }
                            if (singleNote.Attributes.ContainsKey("mimetype"))
                            {
                                newNote.Attributes["mimetype"] = singleNote.Attributes["mimetype"];
                            }
                            if (singleNote.Attributes.ContainsKey("notetext"))
                            {
                                newNote.Attributes["notetext"] = singleNote.Attributes["notetext"];
                            }
                            if (singleNote.Attributes.ContainsKey("stepid"))
                            {
                                newNote.Attributes["stepid"] = singleNote.Attributes["stepid"];
                            }
                            if (singleNote.Attributes.ContainsKey("subject"))
                            {
                                newNote.Attributes["subject"] = singleNote.Attributes["subject"];
                            }

                            #endregion

                            #region Query Lead for Parent Account and Contact(If Exists)
                            step = 8;
                            Entity leadSet = new Entity();
                            QueryExpression queryLead = new QueryExpression();
                            queryLead.EntityName = "lead";
                            queryLead.ColumnSet = new ColumnSet("parentaccountid", "parentcontactid");
                            queryLead.Criteria.AddCondition("leadid", ConditionOperator.Equal, lead.Id);

                            step = 9;
                            leadSet = service.RetrieveMultiple(queryLead).Entities[0];

                            #endregion

                            #region Create Annotation under Account

                            if (leadSet.Attributes.Contains("parentaccountid"))
                            {
                                step = 901;
                                newNote.Attributes["objectid"] = new EntityReference("account", ((EntityReference)leadSet.Attributes["parentaccountid"]).Id);
                                step = 902;
                                newNote.Attributes["objecttypecode"] = "account";
                                service.Create(newNote);
                                step = 903;
                            }
                            else if (context.OutputParameters.Contains("CreatedEntities"))
                            {
                                step = 904;
                                foreach (EntityReference crEntities in ((IEnumerable)context.OutputParameters["CreatedEntities"]))
                                {
                                    step = 905;
                                    if (crEntities.LogicalName == "account")
                                    {
                                        step = 906;
                                        Entity account = service.Retrieve(crEntities.LogicalName, crEntities.Id
                                            , new ColumnSet("accountid", "name", "originatingleadid"));
                                        step = 907;
                                        newNote.Attributes["objectid"] = new EntityReference(account.LogicalName, account.Id);
                                        step = 908;
                                        newNote.Attributes["objecttypecode"] = account.LogicalName;
                                        service.Create(newNote);
                                        step = 909;
                                    }
                                }
                            }

                            #endregion

                            #region Create Annotation Under Contact

                            step = 10;
                            if (leadSet.Attributes.Contains("parentcontactid"))
                            {
                                step = 101;
                                newNote.Attributes["objectid"] = new EntityReference("contact", ((EntityReference)leadSet.Attributes["parentcontactid"]).Id);
                                step = 102;
                                newNote.Attributes["objecttypecode"] = "contact";
                                service.Create(newNote);
                                step = 103;
                            }
                            else if (context.OutputParameters.Contains("CreatedEntities"))
                            {
                                step = 104;
                                foreach (EntityReference crEntities in ((IEnumerable)context.OutputParameters["CreatedEntities"]))
                                {
                                    step = 105;
                                    if (crEntities.LogicalName == "contact")
                                    {
                                        step = 106;
                                        Entity contact = service.Retrieve(crEntities.LogicalName, crEntities.Id
                                            , new ColumnSet("contactid", "originatingleadid"));
                                        step = 107;
                                        newNote.Attributes["objectid"] = new EntityReference(contact.LogicalName, contact.Id);
                                        step = 108;
                                        newNote.Attributes["objecttypecode"] = contact.LogicalName;
                                        service.Create(newNote);
                                        step = 109;
                                    }
                                }
                            }

                            #endregion

                            #region Create Annotation under Opportunity

                            step = 11;
                            if (context.OutputParameters.Contains("CreatedEntities"))
                            {
                                step = 111;
                                foreach (EntityReference crEntities in ((IEnumerable)context.OutputParameters["CreatedEntities"]))
                                {
                                    step = 112;
                                    if (crEntities.LogicalName == "opportunity")
                                    {
                                        step = 113;
                                        Entity opportunity = service.Retrieve(crEntities.LogicalName, crEntities.Id
                                            , new ColumnSet("opportunityid"));
                                        step = 114;
                                        newNote.Attributes["objectid"] = new EntityReference(opportunity.LogicalName, opportunity.Id);
                                        step = 115;
                                        newNote.Attributes["objecttypecode"] = opportunity.LogicalName;
                                        Guid annotationid = service.Create(newNote);
                                        step = 116;

                                        QueryExpression queryAnnotation = new QueryExpression();
                                        queryAnnotation.EntityName = "annotation";
                                        queryAnnotation.ColumnSet = new ColumnSet("createdon", "objectid", "objecttypecode");
                                        queryAnnotation.Criteria.AddCondition("annotationid", ConditionOperator.Equal, annotationid);

                                        step = 117;
                                        Entity newAnnotation = (Entity)service.RetrieveMultiple(queryAnnotation).Entities[0];

                                        if (newAnnotation.Attributes.Contains("createdon"))
                                        {
                                            step = 118;
                                            opportunity.Attributes["fdx_notescopiedontimestamp"] = ((DateTime)newAnnotation.Attributes["createdon"]).AddSeconds(2);
                                            step = 119;
                                            service.Update(opportunity);
                                            step = 120;
                                        }
                                    }
                                }
                            }
                            else
                            {

                            }

                            #endregion
                        }
                    }
                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("Plugin Error: " + ex.Message + ". Exception occurred at step = {0}.", step));
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("Plugin Error: " + ex.Message + ". Exception occurred at step = {0}.", step));
                }
                #endregion
            }
        }
    }
}
