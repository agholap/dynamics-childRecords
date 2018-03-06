using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace UpdateChildRecord
{
    public class ImportData
    {
        public static OrganizationServiceProxy _serviceProxy = null;
        static CrmServiceClient crmSvc = null;
        public static OrganizationServiceProxy serviceProxy
        {
            get
            {
                if(crmSvc == null || !crmSvc.IsReady )
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    crmSvc = new CrmServiceClient(ConfigurationManager.ConnectionStrings["MyCRMServer"].ConnectionString);
                    return crmSvc.OrganizationServiceProxy;
                }
                return crmSvc.OrganizationServiceProxy;
            }
        }

        public static void GetExcelData()
        {
            log4net.Config.BasicConfigurator.Configure();
            log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));
            int recordCount = 0;
            string con =
  @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\importdata\SQLExport.xls;" +
  @"Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [IssuesAndProblems$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        var recordId = dr[8];
                        recordCount++;
                        try
                        {
                            var issueSubjectName = dr[0];
                            var categoryName = dr[1];
                            var description = dr[2];
                            var todayDate = dr[3];
                            var responsiblePerson = dr[4];
                            var resolvedOn = dr[5];
                            var statusName = dr[6];
                            var parentId = dr[7];                        
                            var physicianName = dr[9];
                            var practiceId = dr[10];
                            var practiceName = dr[11];

                            Entity issues = new Entity("hoag_issueproblems");
                            if (todayDate != null && !string.IsNullOrEmpty(todayDate.ToString()))
                            {
                                issues.Attributes.Add("hoag_todaysdate", DateTime.Parse(todayDate.ToString()));
                            }
                            if (resolvedOn != null && !string.IsNullOrEmpty(resolvedOn.ToString()))
                            {
                                issues.Attributes.Add("hoag_resolvedon", DateTime.Parse(resolvedOn.ToString()));
                            }
                            if (recordId != null && !string.IsNullOrEmpty(recordId.ToString()))
                            {
                                issues.Attributes.Add("hoag_recordid", recordId.ToString());
                            }
                            if (description != null && !string.IsNullOrEmpty(description.ToString()))
                            {
                                issues.Attributes.Add("hoag_description", System.Security.SecurityElement.Escape(description.ToString()));
                            }
                            if (practiceId != null && !string.IsNullOrEmpty(practiceId.ToString()))
                            {
                                issues.Attributes.Add("hoag_practicegroup", new EntityReference("account", GetLookupValue(recordId,"account", "hoag_recordid", practiceId.ToString(), log)));
                            }
                            if (parentId != null && !string.IsNullOrEmpty(parentId.ToString()))
                            {

                                issues.Attributes.Add("hoag_physician", new EntityReference("contact", GetLookupValue(recordId,"contact", "hoag_recordid", parentId.ToString(), log)));
                            }
                            if (issueSubjectName != null && !string.IsNullOrEmpty(issueSubjectName.ToString()))
                            {
                                issues.Attributes.Add("hoag_subject", new EntityReference("hoag_subject", GetLookupValue(recordId,"hoag_subject", "hoag_name", issueSubjectName.ToString(), log)));
                            }
                            if (categoryName != null && !string.IsNullOrEmpty(categoryName.ToString()))
                            {
                                issues.Attributes.Add("hoag_category", new EntityReference("hoag_issuecategory", GetLookupValue(recordId,"hoag_issuecategory", "hoag_name", categoryName.ToString(), log)));
                            }
                            serviceProxy.Create(issues);

                            Console.Write("\r{0}   ", recordCount);

                            //Console.Write("\r{0}%  records processed!", recordCount);

                            // Console.Write("{0} records processed!", recordCount);

                        }
                        catch(Exception ex)
                        {
                                log.Error("Row Number " + (recordCount+1) + " Record Id : " + recordId + " " + ex.Message.ToString());

                        }
                    }
                }
            }
            Console.WriteLine("{0} records processed!", recordCount);
        }
        static Dictionary<string, Guid> lookups = new Dictionary<string, Guid>();
        public static Guid GetLookupValue(object recordId, string entityName, string relatedAttribute, string matchValue, log4net.ILog log)
        {
            try
            {

                if (lookups.ContainsKey(matchValue))
                {
                    return lookups[matchValue];
                }
                string fetch = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='{0}'>
                                            <attribute name='{1}' />
                                            <filter type='and'>
                                              <condition attribute='{2}' operator='eq' value='{3}' />
                                            </filter>
                                          </entity>
                                        </fetch>", entityName, entityName + "id", relatedAttribute, System.Security.SecurityElement.Escape(matchValue));
                FetchExpression fetchQuery = new FetchExpression(fetch);
                EntityCollection results = serviceProxy.RetrieveMultiple(fetchQuery);
                lookups.Add(matchValue, results.Entities[0].Id);
                return results.Entities[0].Id;
            }
            catch(Exception ex)
            {
                log.Error("Record Id : " + recordId +  " Related to  " + entityName + "  " + " Attribute " + relatedAttribute +" Value " + matchValue + " " + ex.Message.ToString());
                throw ex;
            }
        }
    }
}
