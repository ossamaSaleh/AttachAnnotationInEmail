using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinkDev.CS.SendingEmailWithAttachment
{
    public class EmailActivity : CodeActivity
    {
        //define output variable

        [Input("SourceEmail")]

        [ReferenceTarget("email")]

        public InArgument<EntityReference> SourceEmail { get; set; }

        protected override void Execute(CodeActivityContext executionContext)

        {

            // Get workflow context

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            //Create service factory

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            
            // Create Organization service

            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Get the target entity from the context
            var Entity = GetConfiguration(service);
            foreach (var item in Entity)
            {
                if (item.Attributes["ldv_name"].ToString() == context.PrimaryEntityName)
                {
                    Entity CaseID = (Entity)service.Retrieve(item.Attributes["ldv_name"].ToString(), context.PrimaryEntityId, new ColumnSet(new string[] { item.Attributes["ldv_primarykey"].ToString() }));

                    AddAttachmentToEmailRecord(service, CaseID.Id, SourceEmail.Get<EntityReference>(executionContext));
                }
            }

            

        }

        private void AddAttachmentToEmailRecord(IOrganizationService service, Guid SourceAccountID, EntityReference SourceEmailID)

        {

            //create email object

            Entity _ResultEntity = service.Retrieve("email", SourceEmailID.Id, new ColumnSet(true));

            QueryExpression _QueryNotes = new QueryExpression("annotation");

            _QueryNotes.ColumnSet = new ColumnSet(new string[] { "subject", "mimetype", "filename", "documentbody" });

            _QueryNotes.Criteria = new FilterExpression();

            _QueryNotes.Criteria.FilterOperator = LogicalOperator.And;

            _QueryNotes.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, SourceAccountID));

            EntityCollection _MimeCollection = service.RetrieveMultiple(_QueryNotes);

            if (_MimeCollection.Entities.Count > 0)

            {  //we need to fetch first attachment

                Entity _NotesAttachment = _MimeCollection.Entities.First();

                //Create email attachment

                Entity _EmailAttachment = new Entity("activitymimeattachment");

                if (_NotesAttachment.Contains("subject"))

                    _EmailAttachment["subject"] = _NotesAttachment.GetAttributeValue<string>("subject");

                _EmailAttachment["objectid"] = new EntityReference("email", _ResultEntity.Id);

                _EmailAttachment["objecttypecode"] = "email";

                if (_NotesAttachment.Contains("filename"))

                    _EmailAttachment["filename"] = _NotesAttachment.GetAttributeValue<string>("filename");

                if (_NotesAttachment.Contains("documentbody"))

                    _EmailAttachment["body"] = _NotesAttachment.GetAttributeValue<string>("documentbody");

                if (_NotesAttachment.Contains("mimetype"))

                    _EmailAttachment["mimetype"] = _NotesAttachment.GetAttributeValue<string>("mimetype");

                service.Create(_EmailAttachment);

            }

            // Sending email

            SendEmailRequest SendEmail = new SendEmailRequest();

            SendEmail.EmailId = _ResultEntity.Id;

            SendEmail.TrackingToken = "";

            SendEmail.IssueSend = true;

            SendEmailResponse res = (SendEmailResponse)service.Execute(SendEmail);

        }
        private DataCollection<Entity> GetConfiguration(IOrganizationService service)
        {
          
            string fetchXML = "<fetch version='1.0' output-format='xml-platform' no-lock='true'  mapping='logical' distinct='true'>" +
                                    "  <entity name='ldv_configuration'>    " +
                                    "    <attribute name='ldv_name' />    "  +
                                    "    <attribute name='ldv_primarykey' />    " +
                                    "  </entity>" +
                                    "</fetch>";

            EntityCollection values = service.RetrieveMultiple(new FetchExpression(fetchXML));

            return values.Entities;
        }
    }
}
