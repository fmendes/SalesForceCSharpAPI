using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmTrac2SF
{
	public class ADInformation
	{
		public string SAMAccountName;
		public string userAccountControl;
		public string givenName;
		public string sn;
		public string LastName { get { return sn; } set { sn = value; } }
		public string FirstName { get { return givenName; } set { givenName = value; } }
		public string LastFirstName { get{ return string.Concat(sn, ", ", givenName );}}
		public string TelephoneNumber;
		public string mail;
		public string title;
		public string department;
		public string company;
		public string physicalDeliveryOfficeName;
		public string displayName;
		public bool Active;
		public bool Found = false;

		public ADInformation(string strUserId)
		{
			SAMAccountName = strUserId;
			LastName = "";
			FirstName = "";
			Found = GetADInformation();
		}

		public ADInformation( string strLastName, string strFirstName )
		{
			SAMAccountName = "";
			LastName = strLastName;
			FirstName = strFirstName;
			Found = GetADInformation();
		}

		public bool GetADInformation()
		{
			string strUserId = "", strFilter = "";

			if(!SAMAccountName.Equals( "" ))
			{
				strUserId = SAMAccountName;

				if(strUserId.Contains( @"\" ))
					strUserId = strUserId.Substring( 5 );

				// only EmCare/EMSC users
				strFilter = string.Format( "(|(&(objectClass=User)(sAMAccountName={0})(|(company=EmCare*)(company=EMSC*))))", strUserId );
			}

			if(!LastName.Equals( "" ))
				// only EmCare/EMSC users
				strFilter = string.Format( "(|(&(objectClass=User)(givenname={0})(sn={1})(|(company=EmCare*)(company=EMSC*))))", FirstName, LastName );

			string strServer = System.Configuration.ConfigurationManager.AppSettings["EMSC"].ToString();
			string strADUser = System.Configuration.ConfigurationManager.AppSettings["LDAPUID"].ToString();
			string strADPwd = System.Configuration.ConfigurationManager.AppSettings["LDAPPwd"].ToString();

			string sLDAPPath = string.Format("LDAP://{0}/DC=EMSC,DC=root01,DC=org", strServer);
			System.DirectoryServices.DirectoryEntry objDE = null;
			System.DirectoryServices.DirectorySearcher objDS = null;
			try
			{
				objDE = new System.DirectoryServices.DirectoryEntry( sLDAPPath, strADUser, strADPwd, System.DirectoryServices.AuthenticationTypes.Secure );

				objDS = new System.DirectoryServices.DirectorySearcher( objDE );

				// get the LDAP filter string based on selections
				objDS.Filter = strFilter;
				objDS.ReferralChasing = System.DirectoryServices.ReferralChasingOption.None;

				//String strResult = String.Format(
				//"(&(objectClass={0})(givenname={1})(sn={2}))",
				//sLDAPUserObjectClass, sFirstNameSearchFilter, sLastNameSearchFilter);
				//string sFilter = 
				//String.Format("(&(objectclass=user)(MemberOf=CN={0},OU=Groups,DC={1},DC=root01,DC=org))",
				//    strGroupName, strDomain);

				objDS.PropertiesToLoad.Add( "userAccountControl" );
				objDS.PropertiesToLoad.Add( "SAMAccountName" );
				objDS.PropertiesToLoad.Add( "givenName" );
				objDS.PropertiesToLoad.Add( "sn" );
				objDS.PropertiesToLoad.Add( "TelephoneNumber" );
				objDS.PropertiesToLoad.Add( "mail" );
				objDS.PropertiesToLoad.Add( "title" );
				objDS.PropertiesToLoad.Add( "department" );
				objDS.PropertiesToLoad.Add( "company" );
				objDS.PropertiesToLoad.Add( "physicalDeliveryOfficeName" );
				objDS.PropertiesToLoad.Add( "displayName" );

				//start searching
				System.DirectoryServices.SearchResultCollection objSRC = objDS.FindAll();

				try
				{
					if( objSRC.Count != 0 )
					{
						//if(objSRC.Count > 1)
						//    Found = Found;

						// grab the first search result
						System.DirectoryServices.SearchResult objSR = objSRC[ 0 ];

						Found = true;

						displayName = objSR.Properties[ "displayName" ][ 0 ].ToString();
						givenName = objSR.Properties[ "givenName" ][ 0 ].ToString();
						sn = objSR.Properties[ "sn" ][ 0 ].ToString();
						SAMAccountName = objSR.Properties[ "SAMAccountName" ][ 0 ].ToString();

						userAccountControl = objSR.Properties[ "userAccountControl" ][ 0 ].ToString();
						int iInactiveFlag = Convert.ToInt32( userAccountControl );
						iInactiveFlag = iInactiveFlag & 0x0002;
						Active = iInactiveFlag <= 0;

						if( objSR.Properties[ "TelephoneNumber" ].Count > 0 )
							TelephoneNumber = objSR.Properties[ "TelephoneNumber" ][ 0 ].ToString();
						if( objSR.Properties[ "mail" ].Count > 0 )
							mail = objSR.Properties[ "mail" ][ 0 ].ToString();
						if( objSR.Properties[ "title" ].Count > 0 )
							title = objSR.Properties[ "title" ][ 0 ].ToString();
						if( objSR.Properties[ "department" ].Count > 0 )
							department = objSR.Properties[ "department" ][ 0 ].ToString();
						if( objSR.Properties[ "company" ].Count > 0 )
							company = objSR.Properties[ "company" ][ 0 ].ToString();
						if( objSR.Properties[ "physicalDeliveryOfficeName" ].Count > 0 )
							physicalDeliveryOfficeName = objSR.Properties[ "physicalDeliveryOfficeName" ][ 0 ].ToString();
					}
					else
					{
						Found = false;
						return Found;
					}
				}
				catch( Exception )
				{
					// ignore errors
					Found = false;
					return false;
				}
				finally
				{
					objDE.Dispose();
					objSRC.Dispose();
					//objDS.Dispose();
				}
			}
			catch( Exception )
			{
				// ignore errors
				Found = false;
				return false;
			}
			finally
			{
				objDS.Dispose();
			}

			return Found;
		}
	}
}