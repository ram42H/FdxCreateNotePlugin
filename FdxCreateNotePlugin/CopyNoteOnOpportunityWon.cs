using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxCreateNotePlugin
{
    public class CopyNoteOnOpportunityWon : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity)
            {
                #region setup plugin steps
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                Entity opportunity = (Entity)context.InputParameters["OpportunityClose"];

                int step = 1;
                if (opportunity.LogicalName != "opportunityclose")
                    return;

                #endregion

                #region Business Code
                step = 2;
                try
                {
                    QueryExpression opportunityQuery = new QueryExpression();
                    opportunityQuery.EntityName = "opportunity";
                    opportunityQuery.ColumnSet = new ColumnSet("fdx_notescopiedontimestamp", "parentaccountid", "parentcontactid", "createdon");
                    opportunityQuery.Criteria.AddCondition("opportunityid",ConditionOperator.Equal, ((EntityReference)opportunity.Attributes["opportunityid"]).Id);

                    Entity opportunitySet = (Entity)service.RetrieveMultiple(opportunityQuery).Entities[0];

                    #region Fetch Annotations of an opportunity
                    step = 3;
                    QueryExpression queryNote = new QueryExpression();
                    queryNote.EntityName = "annotation";
                    queryNote.ColumnSet = new ColumnSet(true);
                    queryNote.Distinct = false;
                    queryNote.Criteria = new FilterExpression();
                    queryNote.Criteria.FilterOperator = LogicalOperator.And;
                    step = 4;
                    if (opportunitySet.Attributes.Contains("opportunityid"))
                    {
                        queryNote.Criteria.AddCondition("objectid", ConditionOperator.Equal, (Guid)opportunitySet.Attributes["opportunityid"]);
                        queryNote.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, opportunitySet.LogicalName);
                        if (opportunitySet.Attributes.Contains("fdx_notescopiedontimestamp"))
                        {
                            queryNote.Criteria.AddCondition("createdon", ConditionOperator.GreaterThan, (DateTime)opportunitySet.Attributes["fdx_notescopiedontimestamp"]);
                        }
                        else if(opportunitySet.Attributes.Contains("createdon"))
                        {
                            queryNote.Criteria.AddCondition("createdon", ConditionOperator.GreaterThan, (DateTime)opportunitySet.Attributes["createdon"]);
                        }
                    }

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

                            #region Create Annotation under Account

                            if (opportunitySet.Attributes.Contains("parentaccountid"))
                            {
                                step = 901;
                                newNote.Attributes["objectid"] = new EntityReference("account", ((EntityReference)opportunitySet.Attributes["parentaccountid"]).Id);
                                step = 902;
                                newNote.Attributes["objecttypecode"] = "account";
                                service.Create(newNote);
                                step = 903;
                            }
                            #endregion

                            #region Create Annotation Under Contact

                            step = 10;
                            if (opportunitySet.Attributes.Contains("parentcontactid"))
                            {
                                step = 101;
                                newNote.Attributes["objectid"] = new EntityReference("contact", ((EntityReference)opportunitySet.Attributes["parentcontactid"]).Id);
                                step = 102;
                                newNote.Attributes["objecttypecode"] = "contact";
                                service.Create(newNote);
                                step = 103;
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
