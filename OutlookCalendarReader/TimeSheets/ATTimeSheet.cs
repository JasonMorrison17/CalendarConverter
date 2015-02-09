using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookCalendarReader.AutoTask;
using System.Net;
using System.ServiceModel;

namespace OutlookCalendarReader.TimeSheets
{
    class ATTimeSheet : ITimeSheet
    {
        private String _user { get; set; }

        private String _password { get; set; }

        private ATWSZoneInfo _zoneInfo { get; set; }

        private ATWSSoapClient _client { get; set; }

        public String UserName
        {
            get
            {
                return _user;
            }

            set
            {
                _user = value;
            }
        }

        public String Password
        {
            set
            {
                _password = value;
            }
        }

        public ATTimeSheet(String userName, String password)
        {
            Console.WriteLine("Initializing ATTimeSheet...");
            this._user = userName;
            this._password = password;
            this._client = new ATWSSoapClient();
            this._zoneInfo = this._client.getZoneInfo(_user);
            Console.WriteLine("Attempting to connect...");
            this.connect();
            this.createTimeEntry();
            //this.getAccounts();
        }

        private void connect()
        {
            Console.WriteLine("Creating binding...");
            BasicHttpBinding binding = new BasicHttpBinding();
            binding.Security.Mode = BasicHttpSecurityMode.Transport;
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Basic;

            binding.MaxReceivedMessageSize = 2147483647;

            Console.WriteLine("Creating endpoint...");
            EndpointAddress ea = new EndpointAddress(this._zoneInfo.URL);

            Console.WriteLine("Intializing client and binding authentication...");
            this._client = new ATWSSoapClient(binding, ea);
            this._client.ClientCredentials.UserName.UserName = this._user;
            this._client.ClientCredentials.UserName.Password = this._password;
        }

        private void createTimeEntry()
        {
            //Get the account + Project Name from calendar category
            //Get Task from calendar subject
            //Query for contract
            //String projName = "Project- Mass RMV FileNet Admin (Webject) ";

            // query for any account. This should return all accounts since we are retreiving anything greater than 0.
            StringBuilder sb = new StringBuilder();
            sb.Append("<queryxml><entity>Project</entity>").Append(System.Environment.NewLine);
            sb.Append("<query><field>ProjectName<expression op=\"equals\">Project- Mass RMV FileNet Admin (Webject) </expression></field></query>").Append(System.Environment.NewLine);
            sb.Append("</queryxml>").Append(System.Environment.NewLine);

            var r = this._client.query(new AutotaskIntegrations(), sb.ToString());
            Project proj = null;
            if (r.ReturnCode == 1)
            {
                if (r.EntityResults.Length > 0)
                {
                    proj = (Project)r.EntityResults[0];
                }                
            }
            
            sb = new StringBuilder();
            sb.Append("<queryxml><entity>Contract</entity>").Append(System.Environment.NewLine);
            sb.Append("<query><field>id<expression op=\"equals\">" + proj.ContractID + "</expression></field></query>").Append(System.Environment.NewLine);
            sb.Append("</queryxml>").Append(System.Environment.NewLine);
            r = this._client.query(new AutotaskIntegrations(), sb.ToString());
            Contract contract = null;
            if (r.ReturnCode == 1)
            {
                if (r.EntityResults.Length > 0)
                {
                    contract = (Contract)r.EntityResults[0];
                }
            }

            sb = new StringBuilder();
            sb.Append("<queryxml><entity>Account</entity>").Append(System.Environment.NewLine);
            sb.Append("<query><field>id<expression op=\"equals\">" + proj.AccountID + "</expression></field></query>").Append(System.Environment.NewLine);
            sb.Append("</queryxml>").Append(System.Environment.NewLine);
            r = this._client.query(new AutotaskIntegrations(), sb.ToString());
            Account account = null;
            if (r.ReturnCode == 1)
            {
                if (r.EntityResults.Length > 0)
                {
                    account = (Account)r.EntityResults[0];
                }
            }

            Console.WriteLine("Queries Completed...");
        }
        
        private void getEntityInfo()
        {
            EntityInfo[] ei = this._client.getEntityInfo(new AutotaskIntegrations());
            foreach (EntityInfo e in ei)
            {
                Console.WriteLine(e.Name);
                //Field[] fi = this._client.GetFieldInfo(new AutotaskIntegrations(), e.Name);
                //foreach (Field f in fi)
                //{
                //    Console.WriteLine(f.Name + ": " + f.Description);
                //}
            }
        }

        private void getAccounts()
        {
            // query for any account. This should return all accounts since we are retreiving anything greater than 0.
            StringBuilder sb = new StringBuilder();
            sb.Append("<queryxml><entity>Account</entity>").Append(System.Environment.NewLine);
            sb.Append("<query><field>id<expression op=\"greaterthan\">0</expression></field></query>").Append(System.Environment.NewLine);
            sb.Append("</queryxml>").Append(System.Environment.NewLine);

            // have no clue what this does.
            AutotaskIntegrations at_integrations = new AutotaskIntegrations();

            // this example will not handle the 500 results limitation.
            // Autotask only returns up to 500 results in a response. if there are more you must query again for the next 500.
            var r = this._client.query(at_integrations, sb.ToString());
            Console.WriteLine("response ReturnCode = " + r.ReturnCode);
            if (r.ReturnCode == 1)
            {
                if (r.EntityResults.Length > 0)
                {
                    foreach (var item in r.EntityResults)
                    {
                        Account acct = (Account)item;                        
                        Console.WriteLine("Account Name = " + acct.AccountName);
                        Console.WriteLine("Account number = " + acct.AccountNumber);
                    }
                }
            }
        }
    }
}
