using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EmTrac2SF.EMSC2SF;

namespace EmTrac2SF.Salesforce
{
    public class ForceConnection
    {
        public string SessionID { get; set; }
        public string ServerUrl { get; set; }
		public bool ConnectionFailed = false;
		public string ErrorMessage = "";

        public ForceConnection(string connectionString)
        {
            ForceConnectionStringBuilder connectionBuilder = 
                new ForceConnectionStringBuilder(connectionString);

			if( Login( connectionBuilder.Username, connectionBuilder.Password, connectionBuilder.Token ) )
				return;
			else
				ConnectionFailed = true;
        }

        public ForceConnection(string username, string password, string securityToken)
        {
            ErrorMessage = "";
			ConnectionFailed = false;
			Login(username, password, securityToken);
        }

        private bool Login(string username, string password, string securityToken)
        {
            try
            {
                using (SforceService service = new SforceService())
                {
                    LoginResult loginResult = 
                        service.login(username, String.Concat(password, securityToken));

                    this.SessionID = loginResult.sessionId;
                    this.ServerUrl = loginResult.serverUrl;
                }

                return true;
            }
            catch (Exception excpt)
            {
				ErrorMessage = excpt.Message;
                return false;
            }
        }
    }
}