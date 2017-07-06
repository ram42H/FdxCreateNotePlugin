using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxCreateNotePlugin
{
    class CopyNoteOnCreateOnOpportunity : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth == 1)
            {
                #region setup plugin steps
                //throw new InvalidPluginExecutionException("Plugin entered plugin");
                string entityName = "opportunity";
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                Entity note = (Entity)context.InputParameters["Target"];

                int step = 1;
                if (note.LogicalName != "annotation")
                    return;
                else if (note.Attributes.Contains("objecttypecode"))
                {
                    if (entityName != note.Attributes["objecttypecode"].ToString())
                    {
                        return;
                    }
                }
                #endregion

                #region Business Code
                step = 2;
                try
                {
                    //#region Fetch Annotation of a Lead
                    //step = 3;
                    //QueryExpression queryNote = new QueryExpression();
                    //queryNote.EntityName = "annotation";
                    //queryNote.ColumnSet = new ColumnSet(true);
                    //queryNote.Distinct = false;
                    //queryNote.Criteria = new FilterExpression();
                    //queryNote.Criteria.FilterOperator = LogicalOperator.And;
                    //step = 4;
                    //queryNote.Criteria.AddCondition("objectid", ConditionOperator.Equal, lead.Id);
                    //queryNote.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, lead.LogicalName);

                    //step = 5;
                    //EntityCollection notes = service.RetrieveMultiple(queryNote);

                    //#endregion

                    #region Process Annotation

                    step = 6;
                    #region Copy Annotation Details to new object

                    Entity newNote = new Entity("annotation");
                    //newNote.Attributes["fdx_notesource"] = 1;
                    if (note.Attributes.ContainsKey("documentbody"))
                    {
                        newNote.Attributes["documentbody"] = note.Attributes["documentbody"];
                    }
                    if (note.Attributes.ContainsKey("filename"))
                    {
                        newNote.Attributes["filename"] = note.Attributes["filename"];
                    }
                    if (note.Attributes.ContainsKey("filename"))
                    {
                        newNote.Attributes["filename"] = note.Attributes["filename"];
                    }
                    if (note.Attributes.ContainsKey("isdocument"))
                    {
                        newNote.Attributes["isdocument"] = note.Attributes["isdocument"];
                    }
                    if (note.Attributes.ContainsKey("langid"))
                    {
                        newNote.Attributes["langid"] = note.Attributes["langid"];
                    }
                    if (note.Attributes.ContainsKey("mimetype"))
                    {
                        newNote.Attributes["mimetype"] = note.Attributes["mimetype"];
                    }
                    if (note.Attributes.ContainsKey("notetext"))
                    {
                        newNote.Attributes["notetext"] = note.Attributes["notetext"];
                    }
                    if (note.Attributes.ContainsKey("stepid"))
                    {
                        newNote.Attributes["stepid"] = note.Attributes["stepid"];
                    }
                    if (note.Attributes.ContainsKey("subject"))
                    {
                        newNote.Attributes["subject"] = note.Attributes["subject"];
                    }

                    #endregion

                    #region Query Opportunity for Parent Account and Contact(If Exists)
                    step = 8;
                    Entity opportunitySet = new Entity();
                    QueryExpression queryOpportunity = new QueryExpression();
                    queryOpportunity.EntityName = "opportunity";
                    queryOpportunity.ColumnSet = new ColumnSet("parentaccountid", "parentcontactid");
                    //queryLead.Criteria = new FilterExpression();
                    queryOpportunity.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, ((EntityReference)note.Attributes["objectid"]).Id);

                    step = 9;
                    opportunitySet = service.RetrieveMultiple(queryOpportunity).Entities[0];

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

                    #endregion
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("Exception occurred at step = {0}. With Message - " + ex.Message, step));
                }
                #endregion
            }
        }
    }
}
