using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace AllCRrolluptoAccount
{
    class Program
    {
        static IOrganizationService organizationService = null;
        // IOrganizationService organizationService = null;
        static void Main(string[] args)
        {
            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = "FuturedonticsCRMAdmin@dentsplysirona.com";
                clientCredentials.UserName.Password = "XP9h#d71fdmK7&iL2";
                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
                // Copy and Paste Organization Service Endpoint Address URL
                organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri("https://fd-stage.api.crm.dynamics.com/XRMServices/2011/Organization.svc"), null, clientCredentials, null);
                if (organizationService != null)
                {
                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;
                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Established Successfully...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.Message);
            }
            // Console.ReadKey();
            CRrolluponAccount();
        }
        static void CRrolluponAccount()
        {
            string crquery = string.Empty;
            string contact = string.Empty;
            string oppcontact = string.Empty;
            string leadquery = string.Empty;
            Entity campaignResponse = new Entity("campaignresponse");

            crquery = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='200'>" +
                                  "<entity name='campaignresponse'>" +
                                    "<attribute name='subject' />" +
                                    "<attribute name='activityid' />" +
                                    "<attribute name='regardingobjectid' />" +
                                    "<attribute name='responsecode' />" +
                                    "<order attribute='createdon' descending='true' />"+
                                    "<attribute name='fdx_reconversionopportunity' />" +
                                    "<attribute name='fdx_reconversionlead' />" +
                                    "<attribute name='fdx_reconversioncontact' />" +
                                    "<attribute name='customer' />" +
                                    "<attribute name='subject' />" +
                                    "<filter type='and'>" +
                                      "<filter type='or'>" +
                                        "<condition attribute='fdx_reconversioncontact' operator='not-null' />" +
                                        "<condition attribute='fdx_reconversionlead' operator='not-null' />" +
                                        "<condition attribute='fdx_reconversionopportunity' operator='not-null' />" +
                                      "</filter>" +
                                    "</filter>" +
                                  "</entity>" +
                                "</fetch>";
            EntityCollection campaignresponses = organizationService.RetrieveMultiple(new FetchExpression(crquery));
            Guid reconcontact = new Guid();
            Guid reconlead = new Guid();
            Guid reconnopp = new Guid();
            EntityCollection oppcont = new EntityCollection();

            if (campaignresponses.Entities.Count > 0)
            {
                foreach (Entity campresp in campaignresponses.Entities)
                {
                    #region Update Contact CR's to its account
                    if (campresp.Attributes.Contains("fdx_reconversioncontact"))
                    {
                         reconcontact = ((EntityReference)campresp["fdx_reconversioncontact"]).Id;

                        QueryByAttribute AccountQueryBycontactId = new QueryByAttribute("contact");
                        AccountQueryBycontactId.AddAttributeValue("contactid", reconcontact);
                        AccountQueryBycontactId.ColumnSet = new ColumnSet("contactid", "parentcustomerid");
                        EntityCollection contactrecords = organizationService.RetrieveMultiple(AccountQueryBycontactId);
                        EntityReference contacc = new EntityReference("account");

                        foreach (Entity contactrec in contactrecords.Entities)
                        {
                            if (contactrec.Attributes.Contains("parentcustomerid"))
                            {

                                contacc.Id = ((EntityReference)contactrec["parentcustomerid"]).Id;
                                //Update Customer on Campaign response
                                EntityReference ContAccount = new EntityReference("account", contacc.Id);

                                Entity customer = new Entity("activityparty");
                                customer.Attributes["partyid"] = ContAccount;
                                EntityCollection Customerentity = new EntityCollection();
                                Customerentity.Entities.Add(customer);
                                campaignResponse["customer"] = Customerentity;
                                campaignResponse["fdx_crrollup"] = true;
                                campaignResponse["activityid"] = campresp.Id;
                                organizationService.Update(campaignResponse);
                            }
                        }
                    }
                    #endregion
                    #region Update Opp CR's to its Contacts Account if it has Account or else update Opp Cr's to it Account
                    else if (campresp.Attributes.Contains("fdx_reconversionopportunity"))
                    {
                        reconnopp = ((EntityReference)campresp["fdx_reconversionopportunity"]).Id;
                        
                        QueryByAttribute contactquerybyopp = new QueryByAttribute("opportunity");
                        contactquerybyopp.AddAttributeValue("opportunityid", reconnopp);
                        contactquerybyopp.ColumnSet = new ColumnSet("parentaccountid", "parentcontactid");
                        EntityCollection opprecords = organizationService.RetrieveMultiple(contactquerybyopp);
                        EntityReference contacts = new EntityReference("contact");
                        EntityReference account = new EntityReference("account");


                        foreach (Entity opprecord in opprecords.Entities)
                        {
                            if (opprecord.Attributes.Contains("parentcontactid"))
                            {
                                contacts.Id = ((EntityReference)opprecord["parentcontactid"]).Id;

                                QueryByAttribute AccountQueryBycontactId = new QueryByAttribute("contact");
                                AccountQueryBycontactId.AddAttributeValue("contactid", contacts.Id);
                                AccountQueryBycontactId.ColumnSet = new ColumnSet("contactid", "parentcustomerid");
                                EntityCollection contactrecords = organizationService.RetrieveMultiple(AccountQueryBycontactId);

                                foreach (Entity contactrec in contactrecords.Entities)
                                {
                                    if (contactrec.Attributes.Contains("parentcustomerid"))
                                    {
                                        account.Id = ((EntityReference)contactrec["parentcustomerid"]).Id;

                                        EntityReference ContAccount = new EntityReference("account", account.Id);

                                        Entity customer = new Entity("activityparty");
                                        customer.Attributes["partyid"] = ContAccount;
                                        EntityCollection Customerentity = new EntityCollection();
                                        Customerentity.Entities.Add(customer);
                                        campaignResponse["customer"] = Customerentity;
                                        campaignResponse["fdx_crrollup"] = true;
                                        campaignResponse["activityid"] = campresp.Id;
                                        organizationService.Update(campaignResponse);
                                    }
                                    else if (opprecord.Attributes.Contains("parentaccountid"))
                                    {
                                        //get Opp account id here
                                        account.Id = ((EntityReference)contactrec["parentaccountid"]).Id;
                                        EntityReference OppAccount = new EntityReference("account", account.Id);

                                        Entity customer = new Entity("activityparty");
                                        customer.Attributes["partyid"] = OppAccount;
                                        EntityCollection Customerentity = new EntityCollection();
                                        Customerentity.Entities.Add(customer);
                                        campaignResponse["customer"] = Customerentity;
                                        campaignResponse["fdx_crrollup"] = true;
                                        campaignResponse["activityid"] = campresp.Id;
                                        organizationService.Update(campaignResponse);

                                    }
                                }
                            }
                        }
                    }
                    #endregion
                    #region Update Lead CR's to its Account
                    else if (campresp.Attributes.Contains("fdx_reconversionlead"))
                    {
                        reconlead = ((EntityReference)campresp["fdx_reconversionlead"]).Id;

                            QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
                            AccountQueryByleadId.AddAttributeValue("leadid", reconlead);
                            AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
                            EntityCollection leadrecords = organizationService.RetrieveMultiple(AccountQueryByleadId);
                            EntityReference acc = new EntityReference("account");

                            foreach (Entity leadrecord in leadrecords.Entities)
                            {
                                if (leadrecord.Attributes.Contains("parentaccountid"))
                                {
                                    acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;

                                    EntityReference LeadAccount = new EntityReference("account", acc.Id);

                                    Entity customer = new Entity("activityparty");
                                    customer.Attributes["partyid"] = LeadAccount;
                                    EntityCollection Customerentity = new EntityCollection();
                                    Customerentity.Entities.Add(customer);
                                    campaignResponse["customer"] = Customerentity;
                                    campaignResponse["fdx_crrollup"] = true;
                                    campaignResponse["activityid"] = campresp.Id;
                                    organizationService.Update(campaignResponse);
                                }
                            }
                    }
                    #endregion
                }
            }
        }
    }
}
