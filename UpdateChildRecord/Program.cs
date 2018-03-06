using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace UpdateChildRecord
{
    class Program
    {
        static void Main(string[] args)
        {
            ImportData.GetExcelData();
            Console.ReadKey();
            return;
            // Define the fetch attributes.
            // Set the number of records per page to retrieve.
            int fetchCount = 5000;
            // Initialize the page number.
            int pageNumber = 1;
            // Initialize the number of records.
            int recordCount = 0;
            // Specify the current paging cookie. For retrieving the first page, 
            // pagingCookie should be null.
            string pagingCookie = null;
            //set Security Protocol
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            CrmServiceClient crmSvc = new CrmServiceClient(ConfigurationManager.ConnectionStrings["MyCRMServer"].ConnectionString);
            Console.WriteLine(crmSvc.IsReady);
            string usersFetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='systemuser'>
                                <attribute name='fullname' />
                                <attribute name='businessunitid' />
                                <attribute name='title' />
                                <attribute name='address1_telephone1' />
                                <attribute name='positionid' />
                                <attribute name='systemuserid' />
                                <attribute name='address1_stateorprovince' />
                                    <filter type='and'>
                                                  <condition attribute='fullname' operator='eq' value='Amol Gholap' />
                                                </filter>
                                <order attribute='fullname' descending='false' />
                              </entity>
                            </fetch>";
            FetchExpression userQuery = new FetchExpression(usersFetch);
            EntityCollection users = crmSvc.OrganizationServiceProxy.RetrieveMultiple(userQuery);
            Console.WriteLine(string.Format("{0} user found", users.Entities.Count));
            foreach (var user in users.Entities)
            {
                Console.WriteLine("Processing " + user.Id);
                //check if user has value for address1_stateorprovince;
                if(!user.Contains("address1_stateorprovince"))
                {
                    continue;
                }
                //get all the contact where created by is user and 
                string contactWithAccountFetch = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='contact'>
                                                <attribute name='fullname' />
                                                <attribute name='telephone1' />
                                                <attribute name='contactid' />
                                                <attribute name='address1_stateorprovince' />
                                                <order attribute='fullname' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='modifiedby' operator='eq' value='{0}' />
                                                </filter>
                                                <link-entity name='account' from='accountid' to='parentcustomerid' visible='false' link-type='outer' alias='customerId'>
                                                  <attribute name='address1_stateorprovince' />
                                                    <attribute name='accountid' />
                                                </link-entity>
                                              </entity>
                                            </fetch>", user.Id);

                while (true)
                {
                    // Build fetchXml string with the placeholders.
                    string xml = CreateXml(contactWithAccountFetch, pagingCookie, pageNumber, fetchCount);

                    // Excute the fetch query and get the xml result.
                    RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                    {
                        Query = new FetchExpression(xml)
                    };

                    EntityCollection returnCollection = ((RetrieveMultipleResponse)crmSvc.OrganizationServiceProxy.Execute(fetchRequest1)).EntityCollection;
                    Console.WriteLine(string.Format("{0} accounts found", returnCollection.Entities.Count));
                    ProcessContacts(user, returnCollection.Entities, crmSvc.OrganizationServiceProxy);

                    // Check for morerecords, if it returns 1.
                    if (returnCollection.MoreRecords)
                    {
                        // Increment the page number to retrieve the next page.
                        pageNumber++;

                        // Set the paging cookie to the paging cookie returned from current results.                            
                        pagingCookie = returnCollection.PagingCookie;
                    }
                    else
                    {
                        // If no more records in the result nodes, exit the loop.
                        break;
                    }
                }
            }


        }
        public static void ProcessContacts(Entity user,DataCollection<Entity> entities, OrganizationServiceProxy service)
        {
            foreach (var contact in entities)
            {
                if (contact.Contains("address1_stateorprovince") && contact.Contains("customerId.address1_stateorprovince"))
                {
                    if (contact["address1_stateorprovince"] != ((AliasedValue)contact["customerId.address1_stateorprovince"]).Value)
                    {
                        //update account
                        var customer = ((AliasedValue)contact["customerId.accountid"]).Value;                      
                        Entity account = new Entity("account");
                        account.Attributes.Add("accountid", customer);
                        account.Attributes.Add("address1_stateorprovince", user["address1_stateorprovince"]);
                        service.Update(account);
                    }
                }
            }
        }
        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }
    }
}
