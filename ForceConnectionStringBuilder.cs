using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace EmTrac2SF.Salesforce
{
    public class ForceConnectionStringBuilder
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string Token { get; set; }

        public string ConnectionString
        {
            get 
            { 
                return String.Format("Username={0}; Password={1}; Token={2};"); 
            }
            set 
            {
                string[] pairs = value.Split(';');

                foreach (string pair in pairs)
                {
                    if (String.IsNullOrWhiteSpace(pair))
                        continue;

                    string[] parts = pair.Split('=');

                    if (parts.Length != 2)
                    {
                        throw new ApplicationException("Malformed connection string parameter.  The connection string should be formated list this: username=value1; password=value2; token=value3;");
                    }

                    string key = parts[0].Trim();
                    string setting = parts[1].Trim();

                    if (String.IsNullOrEmpty(key) || String.IsNullOrEmpty(setting))
                        continue;

                    switch(key.ToLower())
                    {
                        case "username":
                            Username = setting;
                            break;
                        case "password":
                            Password = setting;
                            break;
                        case "token":
                            Token = setting;
                            break;
                        default :
                            throw new ApplicationException(String.Format("Invalid parameter {0}", parts[0]));
                    }
                }
            }
        }
        
        public ForceConnectionStringBuilder()
        {

        }

        public ForceConnectionStringBuilder(string connectionString)
        {
            ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings[connectionString];
            if (settings != null)
            {
                ConnectionString = settings.ConnectionString;
            }
            else
            {
                ConnectionString = connectionString;
            }
        }
    }
}