using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using Company2SF.Salesforce;
using Company2SF.EMSC2SF;
using Metaphone;
using GenericLibrary;
using System.Text;
using System.IO;

namespace Company2SF
{
	public static class SalesForceExtensions
	{
		public static bool IsADuplicateOf( this Contact c, Contact objNewContact )
		{
			// discard if both first/last names don't match
			if( !c.FirstName.Equals( objNewContact.FirstName )
										|| !c.LastName.Equals( objNewContact.LastName ) )
				return false;

			// discard if middle name or SSN are not blank and don't match
			if( !c.Middle_Name__c.IsNullOrBlank() && !objNewContact.Middle_Name__c.IsNullOrBlank()
												&& !c.Middle_Name__c.NullAwareEquals( objNewContact.Middle_Name__c ) )
				return false;
			if( !c.SSN__c.IsNullOrBlank() && !objNewContact.SSN__c.IsNullOrBlank()
										&& !c.SSN__c.NullAwareEquals( objNewContact.SSN__c ) )
				return false;

			// name matches, so we test email, home/mobile phones, birthdate, address, city, ME nbr, work phone
			// (any matches are flagged as duplicate)
			if( c.Email != null && c.Email.NotNullBlankAndEquals( objNewContact.Email ) )
				return true;

			if( c.HomePhone != null && ( c.HomePhone.NotNullBlankAndEquals( objNewContact.HomePhone )
										|| c.HomePhone.NotNullBlankAndEquals( objNewContact.MobilePhone ) ) )
				return true;
			if( c.MobilePhone != null && ( c.MobilePhone.NotNullBlankAndEquals( objNewContact.HomePhone )
										|| c.MobilePhone.NotNullBlankAndEquals( objNewContact.MobilePhone ) ) )
				return true;

			if( c.Birthdate != null && c.Birthdate.NotNullAndEquals( objNewContact.Birthdate ) )
				return true;

			// check the main and other address line #1
			string strOtherStreet = ( objNewContact.OtherStreet ?? "" ).Replace( ".", "" );
			string strAddress1 = objNewContact.Address_Line_1__c.Replace( ".", "" );
			if( c.OtherStreet != null )
			{
				string strCOthStr= c.OtherStreet.Replace( ".", "" );
				if( ( strCOthStr.IsEqualOrPartiallyMatchedTo( strOtherStreet )
											|| strCOthStr.IsEqualOrPartiallyMatchedTo( strAddress1 ) ) )
					return true;
			}
			if( c.Address_Line_1__c != null )
			{
				string strCAddr1 = c.Address_Line_1__c.Replace( ".", "" );
				if( ( strCAddr1.IsEqualOrPartiallyMatchedTo( strAddress1 )
											|| strCAddr1.IsEqualOrPartiallyMatchedTo( strOtherStreet ) ) )
					return true;
			}

			// check main and other city
			string strOtherCity = ( objNewContact.OtherCity ?? "" ).Replace( ".", "" );
			string strCity = objNewContact.City__c.Replace( ".", "" );
			if( c.OtherCity != null )
			{
				string strCOthCty = c.OtherCity.Replace( ".", "" );
				if( ( strCOthCty.IsEqualOrPartiallyMatchedTo( strOtherCity )
						|| strCOthCty.IsEqualOrPartiallyMatchedTo( strCity ) ) )
					return true;
			}
			if( c.City__c != null )
			{
				string strCCty = c.City__c.Replace( ".", "" );
				if( ( strCCty.IsEqualOrPartiallyMatchedTo( strCity )
						|| strCCty.IsEqualOrPartiallyMatchedTo( strOtherCity ) ) )
					return true;
			}

			// check me number
			if( c.MeNumber__c != null && ( c.MeNumber__c.NotNullBlankAndEquals( objNewContact.MeNumber__c )
										|| c.MeNumber__c.NotNullBlankAndEquals( objNewContact.MeNumber__c ) ) )
				return true;

			if( c.Work_Phone__c != null && c.Work_Phone__c.NotNullBlankAndEquals( objNewContact.Work_Phone__c ) )
				return true;

			// if name matches but middle name and ssn cant be compared
			// and neither email, phones and birth dt match, then we cannot conclude it is a duplicate
			return false;
		}

		public static bool IsAMADuplicateOf( this Contact c, Contact objAMAContact )
		{
			// compare MeNumber
			if( !c.MeNumber__c.IsNullOrBlank() )
			{
				// same MeNumber means duplicate
				if( objAMAContact.MeNumber__c.Equals( c.MeNumber__c ) )
					return true;
				else
					return false;
			}

			// compare both first/last names 
			if( !c.FirstName.Equals( objAMAContact.FirstName )
										|| !c.LastName.Equals( objAMAContact.LastName ) )
				return false;

			if( c.HomePhone != null && ( c.HomePhone.NotNullBlankAndEquals( objAMAContact.HomePhone )
										|| c.HomePhone.NotNullBlankAndEquals( objAMAContact.MobilePhone ) ) )
				return true;

			if( c.Birthdate != null && c.Birthdate.NotNullAndEquals( objAMAContact.Birthdate ) )
				return true;

			string strOtherStreet = objAMAContact.OtherStreet.Replace( ".", "" );
			string strAddress1 = objAMAContact.Address_Line_1__c.Replace( ".", "" );
			if( c.OtherStreet != null )
			{
				string strCOthStr= c.OtherStreet.Replace( ".", "" );
				if( ( strCOthStr.IsEqualOrPartiallyMatchedTo( strOtherStreet )
											|| strCOthStr.IsEqualOrPartiallyMatchedTo( strAddress1 ) ) )
					return true;
			}
			if( c.Address_Line_1__c != null )
			{
				string strCAddr1 = c.Address_Line_1__c.Replace( ".", "" );
				if( ( strCAddr1.IsEqualOrPartiallyMatchedTo( strAddress1 )
											|| strCAddr1.IsEqualOrPartiallyMatchedTo( strOtherStreet ) ) )
					return true;
			}

			// if name matches but middle name and ssn cant be compared
			// and neither email, phones and birth dt match, then we cannot conclude it is a duplicate
			return false;
		}
	}

	public class DataLoader
	{
		public ApiService objAPI;
		public ApiService API
		{
			set
			{
				objAPI = value;

				objAPI.OnError = HandleError;
				objAPI.OnReportStatus = ReportStatus;

				// auto configure API according to the setting in the web.config
				strInstance = System.Configuration.ConfigurationManager.AppSettings[ "Instance" ].ToLower();
				switch( strInstance )
				{
					case "dev1":
						objAPI.ConnectionString = "SalesforceDEV1";
						break;
					case "test1":
						objAPI.ConnectionString = "SalesforceLogin";
						break;
					case "prod":
						if( Properties.Settings.Default.Company2SF_EMSC_SF_SforceService.StartsWith( "https://login.salesforce.com" ) )
						{
							objAPI.ConnectionString = "SalesforcePRODUCTION";
						}
						else
							lblError.Text = "ERROR:  In order to connect to Production, please switch the Salesforce URL to https://login.salesforce.com in the web.config.";
						break;
				}
			}
			get { return objAPI; }
		}

		DBAccess objDB;
		string strInstance = "";
		string strAppPath = HttpContext.Current.Request.PhysicalApplicationPath;
		bool bMassDataLoad = false;
		System.Web.UI.WebControls.TextBox tbStatus;
		System.Web.UI.WebControls.Label lblError;

		public DataLoader( string strPath
			, System.Web.UI.WebControls.TextBox tbStatusParam
			, System.Web.UI.WebControls.Label lblErrorParam
			, System.Web.UI.HtmlControls.HtmlGenericControl OpenCSVListParam )
		{
			strAppPath = strPath;	//Request.PhysicalApplicationPath
			tbStatus = tbStatusParam;
			lblError = lblErrorParam;

			objDB = new DBAccess();
			objDB.ErrorLabel = lblErrorParam;

			Company2SFUtils.OpenCSVList = OpenCSVListParam;
		}

		public bool HandleError( string strErrorMessage, string strCommand )
		{
			// report error then save it to a file
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** ERROR:  ", objAPI.ErrorMessage
						, " - Error happened during execution of ", strCommand, " **" );
			Company2SFUtils.SaveStatusCSV( "", tbStatus, true );
			// do not cancel
			return false;
		}

		public bool ReportStatus( params string[] strStatus )
		{
			return Company2SFUtils.ReportStatus( strStatus );
		}

		public void RemoveAMADuplicates()
		{
			string strFirstName = "", strLastName = "";
			bool bContinue = true;
			while( bContinue )
			{
				List<Contact> objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
				, " Birthdate, Email, Address_Line_1__c, City__c, OtherStreet, OtherCity, HomePhone, MobilePhone, MeNumber__c, Work_Phone__c, SSN__C, AMAOnly__c "
				, " FirstName > '" + strFirstName + "' AND LastName > '" + strLastName + "'", " LastName, FirstName, AMAOnly__c LIMIT 2000 " );

				if( objContacts.Count == 0 )
					break;

				strFirstName = objContacts.Last().FirstName;
				strLastName = objContacts.Last().LastName;

				List<Contact> objProvs4Upd = new List<Contact>();

				Contact objPrevious = objContacts.First();
				foreach( Contact objProv in objContacts.Skip( 1 ) )
				{
					if( objProv.LastName.Equals( objPrevious.LastName )
						&& objProv.FirstName.Equals( objPrevious.FirstName ) 
						&& ( objProv.SSN__c.NotNullAndEquals( objPrevious.SSN__c )
							|| objProv.HomePhone.NotNullAndEquals( objPrevious.HomePhone )
							|| objProv.Email.NotNullAndEquals( objPrevious.Email )
							|| objProv.Birthdate.NotNullAndEquals( objPrevious.Birthdate )
						) )
					{
						Contact objAMAContact = null, objCompanyContact = null;
						if( objProv.AMAOnly__c.NotNullAndEquals( "1" ) )
						{
							objAMAContact = objProv;
							objCompanyContact = objPrevious;
						}
						else
							if( objPrevious.AMAOnly__c.NotNullAndEquals( "1" ) )
							{
								objAMAContact = objPrevious;
								objCompanyContact = objProv;
							}
							else 
								// ignore if neither was from AMA
								continue;

						// copy provider to list of contacts for update
						objAMAContact.LastName = objAMAContact.LastName + " (DUPLICATE)";
						objAMAContact.Name = null;
						objProvs4Upd.Add( objAMAContact );

						// copy MENumber if available
						if( !objAMAContact.MeNumber__c.NotNullAndEquals( "" ) && objCompanyContact.MeNumber__c.NotNullAndEquals( "" ) )
						{
							objCompanyContact.MeNumber__c = objAMAContact.MeNumber__c;
							objCompanyContact.Name = null;
							objProvs4Upd.Add( objCompanyContact );
						}

						continue;
					}

					objPrevious = objProv;
				}

				// update providers
				if( objProvs4Upd.Count > 0 )
				{
					SaveResult[] objResults = objAPI.Update( objProvs4Upd.ToArray<sObject>() );
				}

			}

		}

		public void RemoveDuplicateNotes()
		{
			string strSOQL = "SELECT ID, ParentID FROM NOTE WHERE Title like 'SystemName Note:%' "
				+ " order by ParentId , Title DESC";
			List<Note> objNotes = objAPI.Query<Note>( strSOQL );

			// if no notes came back, then we're finished
			if( objNotes.Count == 0 )
				return;

			List<string> objIDList = new List<string>( 4000 );
			string strLastParent = "";
			foreach( Note objNote in objNotes )
				// compare to the previous parent to see if note is duplicate
				if( strLastParent.Equals( objNote.ParentId ) )
				{
					// add duplicate to the list
					objIDList.Add( objNote.Id );

					// every 4k rows, submit deletes
					if( objIDList.Count > 3999 )
					{
						// delete duplicates
						DeleteResult[] objDelResult = objAPI.Delete( objIDList.ToArray() );

						// reset list
						objIDList = new List<string>( 4000 );
					}
				}
				else
					strLastParent = objNote.ParentId;

			// final deletes submission
			if( objIDList.Count > 0 )
			{
				// delete duplicates
				DeleteResult[] objDelResult = objAPI.Delete( objIDList.ToArray() );
			}
		}

		public void LoadCredentialsReportedWithErrors()
		{
			// load Company credentials
			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllCredentials.txt" );

			// load SF agencies
			string strSOQL = 
				"select Id, Name, SystemName_Agency_Match__c, Code__c, City__c from TableStoringAgenciesThatGiveCredentials order by Name";
			List<Credential_Agency__c> objAgencies = objAPI.Query<Credential_Agency__c>( strSOQL );

			// load SF providers (non-AMA)
			List<Contact> objProviders = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
					 , null, " RecruitingID__c != null ", " RecruitingID__c, PhysicianNumber__c DESC " );

			List<Credential__c> objNewCredentials = new List<Credential__c>(); ;

			string strFileName = string.Concat( strAppPath, "_CandidatePendingRows.txt" );
			StreamReader objSR = new StreamReader( strFileName );
			string strLine = "", strPreviousRecrID = "0", strPreviousAgency = "";
			while( ( strLine = objSR.ReadLine() ) != null )
			{
				string[] strData = strLine.Split( '\t' );
				if( strData.Count() < 2 )
					continue;

				if( strData[ 1 ].Contains( ".0 not found" ) )
				{
					// this should be a case of a duplicate provider that was removed
					// so we need to FIND the provider by name and update the credential using the new recr id

					// parse the recruiting id (example "Recr ID: 106532.0 not found")
					string strRecrID = strData[ 1 ].Substring( 9 );
					int iPos = strRecrID.IndexOf( ".0" );
					strRecrID = strRecrID.Substring( 0, iPos );

					// skip repeated recr IDs
					if( strPreviousRecrID.Equals( strRecrID ) )
						continue;

					strPreviousRecrID = strRecrID;

					// get provider name
					string[] strName = strData[ 0 ].Split( "|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries );
					if( strName.Count() < 2 )
					{
						ReportStatus( "Blank provider name: ", strData[0], strData[1] );
						continue;
					}

					string strFirstName = strName[ 0 ];
					string strLastName = strName[ 1 ];

					// if provider not found, skip
					Contact objProv = objProviders.FirstOrDefault( c => c.FirstName.Equals( strFirstName )
											&& c.LastName.Equals( strLastName ) );
					if( objProv == null )
					{
						ReportStatus( "Provider name not found: ", strFirstName, " ", strLastName );
						continue;
					}

					// get the credentials for the recr ID
					DataRow[] objRows = objDT.Select( "RecruitingID = " + strRecrID );
					foreach( DataRow objDR in objRows )
					{
						Credential__c objCred = objDR.ConvertTo<Credential__c>();

						string strCredName = string.Concat( objDR[ "FirstName" ].ToString(), " ", objDR[ "LastName" ].ToString()
							, "-", objCred.Name );

						//string strAgency = objCred.Credential_Agency__c;
						//DataRow[] objAgencyFound = objDTAgencies.Select(
						//            string.Concat( "Company_Agency_Match__c = '", strAgency, "'" ) );
						//if( objAgencyFound.Count() > 0 )
						//    if( !objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString().Equals( "" ) )
						//        strAgency = objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString();
						//objCred.Credential_Agency__c = strAgency;

						// set the external ids in the proper places
						if( !objCred.Credential_Agency__c.Equals( "" ) )
						{
							objCred.Credential_Agency__r = new Credential_Agency__c();
							objCred.Credential_Agency__r.Company_Agency_Match__c = objCred.Credential_Agency__c;
						}
						objCred.Credential_Agency__c = null;

						if( !objCred.Credential_Sub_Type__c.Equals( "" ) )
						{
							objCred.Credential_Sub_Type__r = new Credential_Subtype__c();
							objCred.Credential_Sub_Type__r.Company_SubType_Match__c = objCred.Credential_Sub_Type__c;
						}
						objCred.Credential_Sub_Type__c = null;

						// use the recr id of the provider found with the same name
						objCred.Contact__r = new Contact();
						objCred.Contact__r.RecruitingID__c = objProv.RecruitingID__c;
						objCred.Contact__r.RecruitingID__cSpecified = true;

						// if name too long, reduce to just the credential number and truncate
						if( strCredName.Length > 80 )
							strCredName = objCred.Name;
						objCred.Name = strCredName.Left( 80 );

						objNewCredentials.Add( objCred );

						// each 200 credentials, submit update to SalesForce
						if( objNewCredentials.Count > 200 )
						{
							UpsertResult[] objResults = objAPI.Upsert( "Name", objNewCredentials.ToArray<sObject>() );

							// create CSV file / set the Ids in the list of candidates
							Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNewCredentials );

							objNewCredentials = new List<Credential__c>();
						}
					}

				}

				if( strData[ 1 ].Contains( "Tried to change to contact " ) )
				{
					// this should be a case of a provider with same name or an actual duplicate provider
					// so we need to either CHANGE the credential name or
					// FIND the right credential to update (the one attached to the contact that has physician nbr)

					// parse first, last name - example:  |Jonathan|Hart|-Certification NPDB
					string[] strName = strData[ 0 ].Split( "| ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries );

					if( strName.Count() < 2 )
					{
						ReportStatus( "Blank provider name: ", strData[ 0 ], strData[ 1 ] );
						continue;
					}

					string strFirstName = strName[ 0 ];
					string strLastName = strName[ 1 ];

					// parse cert name/number
					int iPos = strData[ 0 ].IndexOf( "-" );
					string strCertName = strData[ 0 ].Substring( iPos + 1 );
					if( strCertName.Length == 0 )
						continue;

					// parse contact id - example: Tried to change to contact 003E0000006g8a2IAA from  003E0000006VCEIIA4
					string strContactId = strData[ 1 ].ParseFromTo( "contact ", " from " );
					if( strContactId.Length != 18 )
						continue;

					// find the provider by the id
					Contact objProv = objProviders.FirstOrDefault( c => c.Id.Equals( strContactId ) );
					if( objProv == null )
						continue;

					// find the credential
					DataRow[] objRows = objDT.Select( "Name = '" + strCertName 
							+ "' and FirstName = '" + strFirstName + "' and LastName = '" + strLastName + "'"
							+ " and RecruitingID = " + objProv.RecruitingID__c );
					foreach( DataRow objDR in objRows )
					{
						Credential__c objCred = objDR.ConvertTo<Credential__c>();

						string strCredName = string.Concat( objDR[ "FirstName" ].ToString()
							, " ", objDR[ "MiddleName" ].ToString()
							, " ", objDR[ "LastName" ].ToString()
							, "-", objDR[ "Credential_Type__c" ].ToString(), " / ", objCred.Name, " (2)" );
						objCred.Name = strCredName;

						// set the external ids in the proper places
						if( !objCred.Credential_Agency__c.Equals( "" ) )
						{
							objCred.Credential_Agency__r = new Credential_Agency__c();
							objCred.Credential_Agency__r.Company_Agency_Match__c = objCred.Credential_Agency__c;
						}
						objCred.Credential_Agency__c = null;

						if( !objCred.Credential_Sub_Type__c.Equals( "" ) )
						{
							objCred.Credential_Sub_Type__r = new Credential_Subtype__c();
							objCred.Credential_Sub_Type__r.Company_SubType_Match__c = objCred.Credential_Sub_Type__c;
						}
						objCred.Credential_Sub_Type__c = null;

						// use the contact id parsed
						objCred.Contact__c = strContactId;

						// if name too long, reduce to just the credential number and truncate
						if( strCredName.Length > 80 )
							strCredName = objCred.Name;
						objCred.Name = strCredName.Left( 80 );

						objNewCredentials.Add( objCred );

						// each 200 credentials, submit update to SalesForce
						if( objNewCredentials.Count > 200 )
						{
							UpsertResult[] objResults = objAPI.Upsert( "Name", objNewCredentials.ToArray<sObject>() );

							// create CSV file / set the Ids in the list of candidates
							Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNewCredentials );

							objNewCredentials = new List<Credential__c>();
						}
					}
				}

				if( strData[ 1 ].Contains( "Missing agency" ) )
				{
					// this should be a case of a missing agency
					// so we need to reload the credentials after these agencies are created

					// parse the agency id (example:  "Missing agency: board118")
					string strAgencyID = strData[ 1 ].Substring( 16 );
					//int iPos = strAgencyID.IndexOfAny( "0123456789".ToCharArray() );
					//string strAgencyType = strAgencyID.Substring( 0, iPos );
					//strAgencyID = strAgencyID.Substring( iPos, strAgencyID.Length - iPos );

					// skip repeated Agency IDs
					if( strPreviousAgency.Equals( strAgencyID ) )
						continue;

					strPreviousAgency = strAgencyID;

					// note:  license227 doesn't exist

					// get the credentials for the recr ID
					DataRow[] objRows = objDT.Select( "Credential_Agency__c = '" + strAgencyID + "'" );
					foreach( DataRow objDR in objRows )
					{
						Credential__c objCred = objDR.ConvertTo<Credential__c>();

						string strId = objDR[ "RecruitingID" ].ToString();

						string strCredName = string.Concat( objDR[ "FirstName" ].ToString(), " ", objDR[ "LastName" ].ToString()
							, "-", objCred.Name );

						//string strAgency = objCred.Credential_Agency__c;
						//DataRow[] objAgencyFound = objDTAgencies.Select(
						//            string.Concat( "Company_Agency_Match__c = '", strAgency, "'" ) );
						//if( objAgencyFound.Count() > 0 )
						//    if( !objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString().Equals( "" ) )
						//        strAgency = objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString();
						//objCred.Credential_Agency__c = strAgency;

						// set the external ids in the proper places
						if( !objCred.Credential_Agency__c.Equals( "" ) )
						{
							objCred.Credential_Agency__r = new Credential_Agency__c();
							objCred.Credential_Agency__r.Company_Agency_Match__c = objCred.Credential_Agency__c;
						}
						objCred.Credential_Agency__c = null;

						if( !objCred.Credential_Sub_Type__c.Equals( "" ) )
						{
							objCred.Credential_Sub_Type__r = new Credential_Subtype__c();
							objCred.Credential_Sub_Type__r.Company_SubType_Match__c = objCred.Credential_Sub_Type__c;
						}
						objCred.Credential_Sub_Type__c = null;

						// use the recr id of the provider found with the same name
						objCred.Contact__r = new Contact();
						objCred.Contact__r.RecruitingID__c = Convert.ToDouble( strId );
						objCred.Contact__r.RecruitingID__cSpecified = true;

						// if name too long, reduce to just the credential number and truncate
						if( strCredName.Length > 80 )
							strCredName = objCred.Name;
						objCred.Name = strCredName.Left( 80 );

						objNewCredentials.Add( objCred );

						// each 200 credentials, submit update to SalesForce
						if( objNewCredentials.Count > 200 )
						{
							UpsertResult[] objResults = objAPI.Upsert( "Name", objNewCredentials.ToArray<sObject>() );

							// create CSV file / set the Ids in the list of candidates
							Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNewCredentials );

							objNewCredentials = new List<Credential__c>();
						}
					}

				}

			}

			objSR.Dispose();

			// submit last update to SalesForce
			if( objNewCredentials.Count > 0 )
			{
				UpsertResult[] objResults = objAPI.Upsert( "Name", objNewCredentials.ToArray<sObject>() );

				// create CSV file / set the Ids in the list of candidates
				Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNewCredentials );
			}

			return;
		}

		public void LoadMissingCredentials()
		{
			// get list of agencies from Company free from duplicates
			DataTable objDTAgencies = objDB.GetDataTableFromSQLFile( "SQLAllAgencies.txt" );

			Company2SFUtils.FlagDupesByAddressCityName( objDTAgencies );

			bool bContinueProcessing = true;
			string strLastRecrId = "0", strCurrentRecrId = "0";
			int iCount = 0;
			while( bContinueProcessing )
			{
				string strCondition = string.Concat( " AND p.ID > ", strLastRecrId, " " );
				DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllCredentials.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					// interrupt process
					ReportStatus( objDB.ErrorMessage );
					bContinueProcessing = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more Job App rows found." );
					bContinueProcessing = false;
					break;
				}

				// get recr IDs to retrieve equivalent rows from SF
				strLastRecrId = objDT.Rows[ objDT.Rows.Count -1 ][ "RecruitingID" ].ToString();
				strCurrentRecrId = objDT.Rows[ 0 ][ "RecruitingID" ].ToString();

				// retrieve equivalent credentials from SF
				string strSOQL = 
"select Id, Name, Contact__r.RecruitingID__c, Physician_Number__c, TableStoringAgencyGivingCredentials.Company_Agency_Match__c "
+ ", Credential_Number__c , Credential_Type__c, TableStoringSubTypesOfCredentials.Company_SubType_Match__c, State__c from Credential__c "
+ " where Contact__r.RecruitingID__c <= " + strLastRecrId
+ " and Contact__r.RecruitingID__c >= " + strCurrentRecrId
+ " order by Contact__r.RecruitingID__c, Name DESC ";
				List<Credential__c> objCredentials = objAPI.Query<Credential__c>( strSOQL );

				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				//if( objCredentials.Count == 0 )
				//    break;

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " Credential rows retrieved. Total ", iCount.ToString(), " processed.\r\n" );

				List<Credential__c> objNewCredentials = new List<Credential__c>();

				// skip existing credentials and add missing ones *******************
				foreach( DataRow objDR in objDT.Rows )
				{
					Credential__c objCred = objDR.ConvertTo<Credential__c>();

					string strId = objDR[ "RecruitingID" ].ToString();
					string strAgency = objCred.Credential_Agency__c;
					string strNbr = objCred.Credential_Number__c;
					string strType = objCred.Credential_Type__c;
					string strSubType = objCred.Credential_Sub_Type__c;
					string strState = objCred.State__c;
					string strName = string.Concat( objDR[ "FirstName" ].ToString(), " ", objDR[ "LastName" ].ToString()
						, "-", objCred.Name );

					// prepend name to the credential record name
					//objCred.Name = strName;

					// find the non-duplicate equivalent of the agency 
					// (example:  License105 should become 240 Medical Board of CA)
					DataRow[] objAgencyFound = objDTAgencies.Select(
								string.Concat( "Company_Agency_Match__c = '", strAgency, "'" ) );
					if( objAgencyFound.Count() > 0 )
						if( ! objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString().Equals( "" ) )
							strAgency = objAgencyFound[ 0 ][ "DuplicateOfCode" ].ToString();

					objCred.Credential_Agency__c = strAgency;

					// if credential exists, skip it
					Credential__c objFound = objCredentials.FirstOrDefault( c =>
									c.Contact__r.RecruitingID__c.ToString().NotNullAndEquals( strId )
									&& c.Credential_Type__c.NotNullAndEquals( strType ) 
									&& ( c.Credential_Number__c != null ?
										c.Credential_Number__c.NotNullAndEquals( strNbr )
										: strNbr.Equals( "" ) 
										)
									&& ( c.State__c != null ? 
										c.State__c.NotNullAndEquals( strState ) 
										: strState.Equals( "" ) 
										)
									&& ( c.Credential_Agency__r != null ?
										c.Credential_Agency__r.Company_Agency_Match__c.NotNullAndEquals( strAgency ) 
										: strAgency.Equals( "" ) 
										) 
									&& ( c.Credential_Sub_Type__r != null ?
										c.Credential_Sub_Type__r.Company_SubType_Match__c.NotNullAndEquals( strSubType ) 
										: strSubType.Equals( "" ) 
										) 	
									);
					// if found, skip
					if( objFound != null )
						continue;

					// check whether the name of the credential has already been used (example:  licenses for recr Id 22497)
					objFound = objNewCredentials.FirstOrDefault( c => c.Name.Equals( strName ) );
					if( objFound != null )
					{
						// try adding the subtype (it may be different subtype with same nbr - example:  recr Id 22497)
						if( !objCred.Credential_Sub_Type__c.Equals( "" ) )
						{
							strName = string.Concat( objDR[ "FirstName" ].ToString()
												, " ", objDR[ "LastName" ].ToString()
												, "-", objCred.Name, "/", objCred.Credential_Sub_Type__c );

							objFound = objNewCredentials.FirstOrDefault( c => c.Name.Equals( strName ) );
						}

						// if available, insert middle name to avoid duplicates
						if( objFound != null && ! objDR[ "MiddleName" ].IsNullOrBlank() )
						{
							strName = string.Concat( objDR[ "FirstName" ].ToString(), " ", objDR[ "MiddleName" ].ToString()
											, " ", objDR[ "LastName" ].ToString()
											, "-", objCred.Name );

							objFound = objNewCredentials.FirstOrDefault( c => c.Name.Equals( strName ) );
						}

						// if name is still duplicate, add year or " #2"
						if( objFound != null )
						{
							if( objCred.From__c != null )
								strName = string.Concat( strName, " - "
												, ( (DateTime) objCred.From__c ).Year.ToString() );
							else
								strName = string.Concat( strName, " #2" );
						}
					}

					// set the external ids in the proper places
					if( ! objCred.Credential_Agency__c.Equals( "" ) )
					{
						objCred.Credential_Agency__r = new Credential_Agency__c();
						objCred.Credential_Agency__r.Company_Agency_Match__c = objCred.Credential_Agency__c;
					}
					objCred.Credential_Agency__c = null;

					if( ! objCred.Credential_Sub_Type__c.Equals( "" ) )
					{
						objCred.Credential_Sub_Type__r = new Credential_Subtype__c();
						objCred.Credential_Sub_Type__r.Company_SubType_Match__c = objCred.Credential_Sub_Type__c;
					}
					objCred.Credential_Sub_Type__c = null;

					objCred.Contact__r = new Contact();
					objCred.Contact__r.RecruitingID__c = Convert.ToDouble( strId );
					objCred.Contact__r.RecruitingID__cSpecified = true;

					// if name too long, reduce to just the credential number and truncate
					if( strName.Length > 80 )
						strName = objCred.Name;
					objCred.Name = strName.Left( 80 );

					// if credential doesn't exist, add it to the list for upsertion
					objNewCredentials.Add( objCred );
				}

				// upsert missing credentials
				//SaveResult[] objUpdResult = objAPI.Insert( objNewCredentials.ToArray() );

				UpsertResult[] objResults = objAPI.Upsert( "Name", objNewCredentials.ToArray<sObject>() );

				// create CSV file / set the Ids in the list of candidates
				Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNewCredentials );

				// after the last one was processed, stop
				if( objCredentials.Count == 1 )
					bContinueProcessing = false;
			}

			return;
		}

		public void CleanCandidates()
		{

			int iCount = 0, iUpdCount = 0;
			while( iCount > 0 )
			{
				List<Candidate__c> objCandies = objAPI.Query<Candidate__c>(
	"select Id, Name, OwnerId, Contact__c, Facility_Name__c, Contact__r.RecruitingID__c from Candidate__c "
	+ " order by Contact__c, Facility_Name__c" );
				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				if( objCandies.Count == 0 )
					return;
				
				iCount = 0;
				iUpdCount = 0;
				List<string> strIds = new List<string>( 2000 );
				List<Candidate__c> objCandies4Upd = new List<Candidate__c>( 2000 );

				string strLastKey = "";
				Candidate__c objPreviousCandidate = null;
				foreach( Candidate__c objCandidate in objCandies )
				{
					if( iCount == 2000 )
					{
						// if array is full, submit the deletes and reset the array
						DeleteResult[] objDelResult = objAPI.Delete( strIds.ToArray() );
						iCount = 0;
						strIds = new List<string>( 2000 );
					}

					if( iUpdCount == 2000 )
					{
						// if array is full, submit the updates and reset the array
						SaveResult[] objSaveResult = objAPI.Update( objCandies4Upd.ToArray() );
						iUpdCount = 0;
						objCandies4Upd = new List<Candidate__c>( 2000 );
					}

					// check whether this is a duplicate (when contact and facility are the same as previous)
					if( strLastKey.Equals( objCandidate.Contact__c + objCandidate.Facility_Name__c ) )
					{
						// decide whether to delete the current record or the previous
						if( objCandidate.OwnerId.Equals( "005E0000000hMcKIAU" ) ) // "005E0000000hMcKIAU" = Batch Load User
						{
							// add Candidate to the list to delete
							strIds.Add( objCandidate.Id );
							iCount++;

							// correct the name of previous Candidate record
							Candidate__c objCandCopy = new Candidate__c();
							objCandCopy.Id = objPreviousCandidate.Id;
							objCandCopy.Name = objCandidate.Name;
							objCandies4Upd.Add( objCandCopy );
							iUpdCount++;

							continue;
						}
					}

					strLastKey = objCandidate.Contact__c + objCandidate.Facility_Name__c;
					objPreviousCandidate = objCandidate;
				}

				if( iCount > 0 )
				{
					// submit last batch of deletes
					DeleteResult[] objDelResult = objAPI.Delete( strIds.ToArray() );
				}

				if( iUpdCount > 0 )
				{
					// if array is full, submit the updates and reset the array
					SaveResult[] objSaveResult = objAPI.Update( objCandies4Upd.ToArray() );
				}
			}
		}

		public void CleanJobAppsByRecrID()
		{
			bool bContinueProcessing = true;
			string strLastRecrId = "0";
			while( bContinueProcessing )
			{
				List<Job_Application__c> objJobApps = objAPI.Query<Job_Application__c>(
"select Id, Name, Contact__r.FirstName, Contact__r.LastName, Contact__r.RecruitingID__c, Contact__r.Degree__c, LastModifiedDate "
+ "from Job_Application__c "
+ "where Contact__r.RecruitingID__c != null and Contact__r.RecruitingID__c >= " + strLastRecrId
+ " order by Contact__r.RecruitingID__c, LastModifiedDate DESC, Name DESC LIMIT 2000 " );
				// where Name like 'a0iE0%' 
				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				if( objJobApps.Count == 0 )
					break;

				List<string> strIds = new List<string>( 2000 );
				List<Job_Application__c> objJAForUpdate = new List<Job_Application__c>( 2000 );
				DateTime dtCutOff = Convert.ToDateTime( "9/27/2011" );
				Job_Application__c objPrevious = null;
				foreach( Job_Application__c objJA in objJobApps )
				{
					// if record is older than 9/27/11, check if it is a duplicate
					DateTime dtModified = (DateTime) objJA.LastModifiedDate;
					if( dtModified.CompareTo( dtCutOff ) < 0  )
					{
						// if it is a duplicate, collect for removal
						if( objJA.Contact__r.RecruitingID__c.ToString().Equals( strLastRecrId ) )
						{
							// add Job App to the list to delete
							strIds.Add( objJA.Id );

							// correct the Provider Name & Degree on the previous record
							objPrevious.Provider_Name_Degree__c =
								string.Concat( objJA.Contact__r.LastName, ", ", objJA.Contact__r.FirstName
								, ", ", objJA.Contact__r.Degree__c );

							// make a copy of previous record for the update
							// because SalesForce doesn't accept more than one reference in Contact__r f/update
							Job_Application__c objJACopy = new Job_Application__c();
							objJACopy.Id = objPrevious.Id;
							objJACopy.Provider_Name_Degree__c = objPrevious.Provider_Name_Degree__c;
							objJAForUpdate.Add( objJACopy );

							continue;
						}
					}

					#region Previous Clean up
					//// previous clean up
					//// if it is a duplicate that starts with a0iE0, collect for removal
					//if( !objJA.Name.StartsWith( "Job App" ) )
					//{
					//    if( objJA.Contact__r.RecruitingID__c.ToString().Equals( strLastRecrId ) )
					//    {
					//        // add Job App to the list to delete
					//        strIds.Add( objJA.Id );
					//        continue;
					//    }

					//    // it is not a duplicate, so we correct the name
					//    objJA.Name = string.Concat( "Job Application for ", objJA.Contact__r.FirstName, " ", objJA.Contact__r.LastName );

					//    // make a copy of object because SalesForce doesn't accept more than one reference in Contact__r f/update
					//    Job_Application__c objJACopy = new Job_Application__c();
					//    objJACopy.Id = objJA.Id;
					//    objJACopy.Name = objJA.Name;
					//    objJAForUpdate.Add( objJACopy );
					//}
					#endregion

					strLastRecrId = objJA.Contact__r.RecruitingID__c.ToString();

					objPrevious = objJA;
				}

				DeleteResult[] objDelResult = objAPI.Delete( strIds.ToArray() );

				SaveResult[] objUpdResult = objAPI.Update( objJAForUpdate.ToArray() );

				// after the last one was processed, stop
				if( objJobApps.Count == 1 )
					bContinueProcessing = false;
			}

			return;
		}

		public void CleanJobAppsByMENbr()
		{
			ReportStatus( "Starting Job App clean up." );

			bool bContinueProcessing = true;
			int iCount = 0;
			string strLastMENumber = "";
			while( bContinueProcessing )
			{
				List<Job_Application__c> objJobApps = objAPI.Query<Job_Application__c>(
"select Id, Name, Contact__r.FirstName, Contact__r.LastName, Contact__r.MeNumber__c "
+ "from Job_Application__c "
+ "where Contact__r.AMAOnly__c = '1' and Contact__r.MeNumber__c >= '" + strLastMENumber
+ "' order by Contact__r.MeNumber__c, Name DESC LIMIT 2000 " );
				// and ( NOT Name LIKE 'Job App%' ) 
				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				iCount++;
				ReportStatus( "Retrieved Job Apps with ME Number >= ", strLastMENumber, " - # ", iCount.ToString() );

				if( objJobApps.Count == 0 )
					break;

				List<string> strIds = new List<string>( 2000 );
				List<Job_Application__c> objJAForUpdate = new List<Job_Application__c>( 2000 );
				foreach( Job_Application__c objJA in objJobApps )
				{
					// if it is a duplicate that starts with a0iE0, collect for removal
					if( !objJA.Name.StartsWith( "Job App" ) )
					{
						if( objJA.Contact__r.MeNumber__c.ToString().Equals( strLastMENumber ) )
						{
							ReportStatus( "Preparing to remove Job App ", objJA.Name, " with ME# ", strLastMENumber );

							// add Job App to the list to delete
							strIds.Add( objJA.Id );
							continue;
						}

						// it is not a duplicate, so we correct the name
						objJA.Name = string.Concat( "Job Application for ", objJA.Contact__r.FirstName, " ", objJA.Contact__r.LastName );

						ReportStatus( "Preparing to rename Job App with ME# ", strLastMENumber, " to ", objJA.Name );

						// make a copy of object because SalesForce doesn't accept more than one reference in Contact__r f/update
						Job_Application__c objJACopy = new Job_Application__c();
						objJACopy.Id = objJA.Id;
						objJACopy.Name = objJA.Name;
						objJAForUpdate.Add( objJACopy );
					}

					strLastMENumber = objJA.Contact__r.MeNumber__c.ToString();
				}

				ReportStatus( "Removing ", strIds.Count.ToString(), " duplicate Job Apps." );

				DeleteResult[] objDelResult = objAPI.Delete( strIds.ToArray() );

				ReportStatus( "Renaming ", objJAForUpdate.Count.ToString(), " Job Apps." );

				SaveResult[] objUpdResult = objAPI.Update( objJAForUpdate.ToArray() );

				// after the last one was processed, stop
				if( objJobApps.Count == 1 )
					bContinueProcessing = false;
			}

			ReportStatus( "Finished cleaning up Job Apps." );

			return;
		}

		public void BulkLoadJobAppFromAMA()
		{

			// create a job
			string strJobId = objAPI.RESTCreateJob( "upsert", "Job_Application__c", "MeNumber__c", "CSV" );
			if( strJobId.Contains( "ERROR" ) )
				return;

			// create CSV files for Company providers
			int iCount = CreateAMAJobAppCSV( strJobId );

			// close the job after submitting batches
			string strState = objAPI.RESTSetJobState( strJobId, "Closed" );	// to abort, send "Aborted"

		}

		public void BulkLoadJobAppFromCompany()
		{

			// create a job
			string strJobId = objAPI.RESTCreateJob( "upsert", "Job_Application__c", "RecruitingID__c", "CSV" );
			if( strJobId.Contains( "ERROR" ) )
				return;

			// create CSV files for Company providers
			int iCount = CreateCompanyJobAppCSV( strJobId );

			// close the job after submitting batches
			string strState = objAPI.RESTSetJobState( strJobId, "Closed" );	// to abort, send "Aborted"



			//            // *** CSV file containing the contacts goes here ***
			//            string strCSVContent = "";
			////@"FirstName,LastName,AccountId,RecruitingID__c
			////""Test"",""Test"",""001E0000002U9FC"",""0""";	// 001E0000002U9FC = Company

			//            // load CSV file
			//            string strFileName = string.Concat( Company2SFUtils.strPath, "JobApp1.csv" );
			//            StreamReader objSR = new StreamReader( strFileName );
			//            strCSVContent = objSR.ReadToEnd();
			//            objSR.Close();

			//            // create and submit a batch
			//            string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
			//            if( strBatchId.Contains( "ERROR" ) )
			//                return;

			//// verify the batch status
			//string strBatchState = objAPI.RESTCheckBatch( strJobId, strBatchId );
			//if( strBatchState.Contains( "Failed" ) )	// "Completed" if successful
			//    return;

			//string strResult = objAPI.RESTGetBatchResult( strJobId, strBatchId );

			//tbStatus.Text = "Job Id: " + strJobId + " - Batch Id: " + strBatchId + " - State: " + strState + "\r\n" + strResult
			//    + "\r\n" + strBatchState;
		}

		public void BulkLoadProviderFromCompany()
		{
			// create a job
			string strJobId = objAPI.RESTCreateJob( "upsert", "Contact", "RecruitingID__c", "CSV" );
			if( strJobId.Contains( "ERROR" ) )
				return;

			// create CSV files for Company providers
			int iCount = LoadCompanyProviderCSV( strJobId );

			// close the job after submitting batches
			string strState = objAPI.RESTSetJobState( strJobId, "Closed" );	// to abort, send "Aborted"

		}

		public void DeleteExistingProviderNotes( List<Contact> objContacts, string strWhere = "" )
		{
			string strSOQL = "select Id, ParentId from Note where Title like 'Company Note%' ";
			if( !strWhere.Equals( "" ) )
				strSOQL = strSOQL + " " + strWhere;

			// delete existing notes whose Title starts with "Company Note" (only providers have notes like that)
			List<Note> objExistingNotes = objAPI.Query<Note>( strSOQL );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				ReportStatus( objAPI.ErrorMessage );
				return;
			}

			// if list of contacts was specified, create list of note ids for the contacts
			List<string> strNoteIds	= new List<string>( objExistingNotes.Count );
			if( objContacts != null )
			{
				// create a list of ids and delete existing notes for the contacts we just updated
				// copy ids to string list
				foreach( Note objN in objExistingNotes )
				{
					Contact objC = objContacts.FirstOrDefault( p => p.Id == objN.ParentId );
					// if note doesn't belong to any of the contacts we updated, do not delete it
					if( objC == null )
						continue;

					strNoteIds.Add( objN.Id );
				}
			}
			else
				// copy ids to string list
				foreach( Note objN in objExistingNotes )
					strNoteIds.Add( objN.Id );

			// delete the notes for the contacts we just updated
			objAPI.Delete( strNoteIds.ToArray() );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				ReportStatus( objAPI.ErrorMessage );
				return;
			}

			// release memory
			strNoteIds.Clear();
			objExistingNotes.Clear();
			return;
		}

		public void BulkLoadMissingCompanyNotes()
		{
			//Batch Load User :  005E0000000hMcK  Fernando:  005E0000000gyNa

			string strLastRecrId = "0";
			while( true )
			{
				List<Contact> objProviders = objAPI.Query<Contact>(
					"select Id, Name, RecruitingID__c, ( select Id from Notes where Title like 'Company Note%' ) from Contact "
					+ " where RecruitingID__c > " + strLastRecrId + " order by RecruitingID__c limit 4000" );
				if( objProviders.Count == 0 )
					break;

				// get equivalent note rows from Company
				string strCondition = string.Concat( 
					" AND p.ID BETWEEN ", objProviders.First().RecruitingID__c.ToString()
								, " AND ", objProviders.Last().RecruitingID__c.ToString() );
				DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllProvidersNotes.txt", null, strCondition );

				// if no notes, skip it
				if( objDT.Rows.Count == 0 )
					continue;

				// check whether there is a note in 
				foreach( Contact objC in objProviders )
				{
					// skip contacts that already have notes in EmForce
					QueryResult objQRNotes = objC.Notes;
					if( objQRNotes != null && objQRNotes.size > 0 )
						continue;

					// find Company equivalent row
					DataRow[] objRows = objDT.Select( "ParentId = " + objC.RecruitingID__c.ToString() + "" );

					// if no notes in Company, skip it
					if( objRows.Count() == 0 )
						continue;

					string strBody = objRows[ 0 ][ "Body" ].ToString();
					string strTitle =  objRows[ 0 ][ "Title" ].ToString().Left( 80 );

					// if note is empty, skip it
					if( strBody.Length == 0 )
						continue;

					// import note into EmForce
					ReportStatus( "Attempting to add note for recr Id ", objC.RecruitingID__c.ToString() );
						
					Note[] objNoteList = new Note[ objRows.Count() + 1 ];

					string str2ndBody = "";
					if( strBody.Length > 31999 )
					{
						str2ndBody = strBody.Substring( 32000 );
						strBody = strBody.Substring( 0, 31999 );
					}

					Note objNote = new Note();
					objNote.ParentId = objC.Id;
					objNote.Body = strBody;
					objNote.Title = strTitle;
					objNote.IsPrivate = false;						
					objNoteList[0] = objNote;

					if( ! str2ndBody.Equals( "" ) )
					{
						Note obj2ndNote = new Note();
						obj2ndNote.ParentId = objC.Id;
						obj2ndNote.Body = str2ndBody;
						obj2ndNote.Title = strTitle + " (continued)";
						obj2ndNote.IsPrivate = false;
						objNoteList[ 1 ] = obj2ndNote;

					}

					if( objRows.Count() > 1 )
					{
						Note objNote2 = new Note();
						objNote2.ParentId = objC.Id;
						objNote2.Body = strBody;
						objNote2.Title = objRows[ 1 ][ "Title" ].ToString();
						objNote2.IsPrivate = false;
						if( str2ndBody.Equals( "" ) )
							objNoteList[ 1 ] = objNote2;
						else
							objNoteList[ 2 ] = objNote2;
					}

					SaveResult[] objResult = objAPI.Insert( objNoteList );
					if( objResult.Count() > 1 && objResult[ 1 ].errors != null )
						ReportStatus( "ERROR: ", objResult[ 1 ].errors[ 0 ].message
							, " for recr ID ", objC.RecruitingID__c.ToString() );
					if( objResult[ 0 ].errors != null )
						ReportStatus( "ERROR: ", objResult[0].errors[0].message
							, " for recr ID ", objC.RecruitingID__c.ToString() );

				}

				strLastRecrId = objProviders.Last<Contact>().RecruitingID__c.ToString();
			}
		}

		public void BulkLoadProviderNotesFromCompany()
		{
			// update only notes after 2011-08-16T19:07:00.000Z

			// create a job
			string strJobId = objAPI.RESTCreateJob( "insert", "Note", "", "CSV" );
			if( strJobId.Contains( "ERROR" ) )
				return;

			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			// only delete notes that were not modified by another person
			DeleteExistingProviderNotes( new List<Contact>(), " and LastModifiedById = '005E0000000gyNa' " );

			DataTable objDT = null;

			string strColumnList = "ParentId,Body,Title,IsPrivate";

			int iCount = 0, iIterationNbr = 1;
			bool bKeepLoading = true;
			string strLastRecrId = "0";
			List<Contact> objContacts = null;
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of Provider rows with RecrID > ", strLastRecrId );

				string strCondition = string.Concat( " AND p.ID > ", strLastRecrId, " " );
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllProvidersNotes.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more Provider rows found." );
					bKeepLoading = false;
					break;
				}

				// get existing contacts that came from Company
				strCondition = string.Concat( " RecruitingID__c > ", strLastRecrId, " " );
				objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
					, null, strCondition, " RecruitingID__c limit 4000 " );

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " Provider rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// start the CSV data with the header
				StringBuilder strbCSVData = new StringBuilder();
				strbCSVData.AppendLine( strColumnList );

				List<Note> objNotes = ValidateNotes( objDT, objContacts );
				foreach( Note objNote in objNotes )
					strbCSVData.Append( objNote.ToSalesForceCSVString( strColumnList ) );

				// store the last MeNumber for the next iteration
				strLastRecrId = objDT.Rows[ objDT.Rows.Count - 1 ][ "ParentID" ].ToString();

				// save CSV file
				string strFileName = string.Concat( Company2SFUtils.strPath, "ProviderNotes", iIterationNbr.ToString(), ".csv" );

				// save new provider contacts in CSV file
				string strCSVContent = strbCSVData.ToString();
				File.WriteAllText( strFileName, strCSVContent );

				// create and submit a batch
				string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
				if( strBatchId.Contains( "ERROR" ) )
					throw new Exception( strBatchId );

				iIterationNbr++;
			}

			ReportStatus( iCount.ToString(), " Total Contact rows processed. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished Contact CSV file. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );

			// close the job after submitting batches
			string strState = objAPI.RESTSetJobState( strJobId, "Closed" );	// to abort, send "Aborted"
		}

		public List<Note> ValidateNotes( DataTable objDT, List<Contact> objContacts )
		{
			// load comments from datatable to note list
			List<Note> objNotes = new List<Note>( objDT.Rows.Count );
			int iSkipped = 0;
			DateTime	dtCreated	= DateTime.Now;
			foreach( DataRow objDR in objDT.Rows )
			{
				// copy all datatable columns to contact object attributes
				Note objNewNote = objDR.ConvertTo<Note>( true );

				// eliminate special characters and normalize line breaks
				objNewNote.Body = objNewNote.Body.RemoveSpecialCharacters( bReplaceWithSpace: true, bSkipNLCR: false )
					.Replace( "\r", "\n" ).Replace( "\n\n", "\n" ).Replace( "\n\n", "\n" );

				// link note to Contact/Provider using RecruitingID
				double dblRecruitingID = Convert.ToDouble( objNewNote.ParentId );
				string strContactId = "";
				Contact objProvider = objContacts.FirstOrDefault( i => i.RecruitingID__c == dblRecruitingID );
				if( objProvider != null )
					strContactId = objProvider.Id;
				else // skip if provider was not found
				{
					ReportStatus( "Skipped comments for RecID ", dblRecruitingID.ToString()
						, " (could not be found).\r\n" );
					iSkipped++;
					continue;
				}

				objNewNote.ParentId = strContactId;
				if( objProvider.CreatedDate != null )
					dtCreated = (DateTime) objProvider.CreatedDate;
				else
					dtCreated = DateTime.Now;
				objNewNote.Title = string.Concat( "Company Note: ", objProvider.FirstName, " ", objProvider.LastName
							, " on ", dtCreated.ToShortDateString() );

				objNotes.Add( objNewNote );

				ReportStatus( "Processed contact note ", objNewNote.Title
								, " (", objNotes.Count.ToString(), "/", objDT.Rows.Count.ToString(), ")" );
			}

			ReportStatus( iSkipped.ToString(), " comment rows skipped (RecID mismatch).\r\n" );

			return objNotes;
		}

		public int LoadCompanyProviderCSV( string strJobId )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;

			System.Text.RegularExpressions.Regex objRegExZip = new System.Text.RegularExpressions.Regex(
						@"\d{5}([ -]\d{4})*" ); // roughly 12345-1234

			// roughly abc-abc+abc.abc @ abc-abc.abc-abc.abcd
			System.Text.RegularExpressions.Regex objRegExEmail = new System.Text.RegularExpressions.Regex(
						@"\w+([-+.]\w+)*@(\w+([-]\w+)*(\.\w+([-]\w+)*)*)+\.(edu|com|[A-Za-z]{2,4})" );
			//old:  @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" );

			System.Text.RegularExpressions.Regex objRegExPhone = new System.Text.RegularExpressions.Regex(
						@"([\(]*\d*[\)]*[ .-])*\d*[ .-]\d*([ ]*\w{0,4}[ ]?\d+)?" ); // roughly (1234)-1234-1234 x123

			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			List<EMSC2SF.Sub_Region__c> objSubRegions = objAPI.Query<EMSC2SF.Sub_Region__c>(
									"select ID, Name, Recruiting_Manager__c, Sub_Region_Description__c from Sub_Region__c " );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return -1;
			}

			Account objAcct = objAPI.QuerySingle<Account>( "select Id, Name from Account where Name = 'CompanyName'" );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return -1;
			}

			// standard specialties to which to compare the imported specialty
			List<string> objSpecialties = new List<string>( new string[] { "Anesthesiology", "Critical Care"
					, "Emergency Medicine", "Family Practice", "General Surgery", "Internal Medicine"
					, "Occupational Medicine", "Other", "Pediatrics", "Psychiatry", "Radiology" } );
			// keywords to replace in specialty
			string[] strSpecialtySearchFor = { " pediatric ", " surgery ", " primary care ", " med "
												, " urgent care ", " wound care ", " fp "
												, " im " };
			string[] strSpecialtySubstitutions = { " pediatrics ", " general surgery ", " family practice ", " medicine "
												, " emergency medicine ", " emergency medicine ", " family practice "
												, " internal medicine " };

			// standard lead sources to which to compare the imported sources
			List<string> objLeadSources = new List<string>( new string[] { 
				"Acquisition", "Ad", "Agency", "Cold Call", "Conference", "Email Campaign", "EmSource", "Facebook", "Internet"
				, "Job Fair", "Mailing", "Preceptor Program", "Referral", "Residency", "Unknown", "Word of Mouth" } );
			List<string> obj2ndLeadSources = new List<string>( new string[] { 				
				"Acquisition", "Agency", "Cold Call", "Email Campaign", "EmSource", "Facebook", "Mailing", "Preceptor Program"
				, "Residency", "Unknown", "Word of Mouth" 
				, "Advance PA/NP", "Classified Ad", "EM News", "EP Monthly", "IM", "Other", "PA World", "WEM Annals", "ACEP"
				, "ACOEP", "ACP-ASIM", "AOA", "GSACEP", "NAIP", "Other", "SHM", "ACEP/EMRA", "Advance PA/NP", "CareerBuilder"
				, "CareerMD", "EDPhysician.com", "Company Careers", "FAPA", "GasWork", "HealtheCareers", "hospitalistjobs"
				, "hospitalistworking", "MDJobSite", "NAPR", "Other", "Physicianwork", "PracticeLink", "CareerMD", "EMRA"
				, "Incentive Eligible", "Incentive Ineligible", "Non-Company", "Recruiter" } );

			// keywords to replace in lead source
			string[] strSearchFor = { " journal ", " at hospital ", " career ", " commercial ", " referral-hospital "
									, " not incentive eligible ", " advertisement ", " email/fax ", " annals "
									, " web page ", " meeting ", " website ", " newspaper " };

			// keywords that will replace the above in lead source
			string[] strSubstitutions = { " ad ", " acquisition ", " job ", " campaign ", " acquisition "
										, " incentive ineligible ", " ad ", " email ", " conference "
										, " internet ", " conference ", " internet ", " ad " };

			string strColumnList = 
"Birthdate,FirstName,LastName,Middle_Name__c,Suffix__c,Degree__c,Salutation,Gender__c,SSN__c,Address_Line_1__c,"
+ "Address_Line_2__c,City__c,State__c,Zipcode__c,OtherStreet,OtherCity,OtherState,OtherPostalCode,"
+ "Email,Contacts_Email_No_2__c,Contacts_Email_No_3__c,HomePhone,Work_Phone__c,OtherPhone,MobilePhone,"
+ "Fax,US_Citizen__c,Birthplace__c,Foreign_Languages__c,Risk_Category__c,Specialty__c,Specialties_Description__c,"
+ "Spouse_Name__c,Spouse_Occupation__c,Spouse_Home_State__c,Children_Information__c,Hobbies__c,Deceased__c,"
+ "Owning_Region__c,Lead_Source_1__c,Lead_Source_Text__c,Drivers_License_Expiration_Date__c,"
+ "Drivers_License_Number__c,Drivers_License_State__c,PhysicianNumber__c,RecruitingID__c,Active_Military__c,"
+ "MeNumber__c,Directorship_Experience__c,NPI__c,UPIN__c,OwnerId,Metaphone_Name__c,Metaphone_Address__c,Metaphone_City__c,AccountId";

			int iCount = 0, iIterationNbr = 1;
			bool bKeepLoading = true;
			string strLastRecrId = "0";
			List<Contact> objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
, " Birthdate, Email, Address_Line_1__c, City__c, OtherStreet, OtherCity, HomePhone, MobilePhone, MeNumber__c, Work_Phone__c, SSN__C "
, " RecruitingID__c > 0 AND ( NOT ( LastName LIKE 'DELETE***%' ) ) ", " LastName, FirstName " );
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of Provider rows with RecrID > ", strLastRecrId );

				string strCondition = string.Concat( " AND p.ID > ", strLastRecrId, " " );
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllProviders2CSV.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more Provider rows found." );
					bKeepLoading = false;
					break;
				}

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " Provider rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// start the CSV data with the header
				StringBuilder strbCSVData = new StringBuilder();
				strbCSVData.AppendLine( strColumnList );

				// validate and process contact rows
				foreach( DataRow objDR in objDT.Rows )
				{
					Contact objNewContact = ValidateProviderContact( objDT, objRegExZip, objRegExEmail, objRegExPhone, objExistingUsers, objSubRegions, objSpecialties, strSpecialtySearchFor, strSpecialtySubstitutions, objLeadSources, obj2ndLeadSources, strSearchFor, strSubstitutions, objContacts, objDR );

					if( objNewContact != null )
					{
						objNewContact.AccountId = objAcct.Id;

						objContacts.Add( objNewContact );

						strbCSVData.Append( objNewContact.ToSalesForceCSVString( strColumnList ) );
					}
				}

				// store the last MeNumber for the next iteration
				strLastRecrId = objDT.Rows[ objDT.Rows.Count - 1 ][ "RecruitingID__c" ].ToString();

				// save CSV file
				string strFileName = string.Concat( Company2SFUtils.strPath, "Provider", iIterationNbr.ToString(), ".csv" );

				// save new provider contacts in CSV file
				string strCSVContent = strbCSVData.ToString();
				File.WriteAllText( strFileName, strCSVContent );

				ReportStatus( "** Saved file ", strFileName );

				// create and submit a batch
				string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
				if( strBatchId.Contains( "ERROR" ) )
					return -1;

				ReportStatus( "** Submitted batch id: ", strBatchId, " for Job Id ", strJobId );

				iIterationNbr++;
			}

			ReportStatus( iCount.ToString(), " Total Contact rows processed. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished Contact CSV file. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );
			return iIterationNbr - 1;
		}

		public int CreateAMAJobAppCSV( string strJobId )
		{

			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;

			ReportStatus( DateTime.Now.ToString(), " Starting refresh of Job Application. Loading data from AMA.\r\n" );

			//// roughly abc-abc+abc.abc @ abc-abc.abc-abc.abcd
			//System.Text.RegularExpressions.Regex objRegExEmail = new System.Text.RegularExpressions.Regex(
			//            @"\w+([-+.]\w+)*@(\w+([-]\w+)*(\.\w+([-]\w+)*)*)+\.(edu|com|[A-Za-z]{2,4})" );

			// do the load in batches of 10000 rows each			
			int iCount = 0, iIterationNbr = 1;
			bool bKeepLoading = true;
			string strLastMeNumber = "0";
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of Job App rows with MeNumber > ", strLastMeNumber );

				string strCondition = string.Concat( " AND p.MeNumber > '", strLastMeNumber, "' " );
				objDT = objDB.GetDataTableFromSQLFile( "SQL_AMA_JobApplications.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more Job App rows found." );
					bKeepLoading = false;
					break;
				}

				// store the last MeNumber for the next iteration
				strLastMeNumber = objDT.Rows[ objDT.Rows.Count - 1 ][ "MeNumber__c" ].ToString();

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " Job App rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// convert data to upper case initials where appropriate
				foreach( DataRow objDR in objDT.Rows )
				{
					objDR[ "First_Name__c" ] = Util.Capitalize( objDR[ "First_Name__c" ].ToString() );
					objDR[ "Last_Name__c" ] = Util.Capitalize( objDR[ "Last_Name__c" ].ToString() );
					objDR[ "Address_Line_1__c" ] = Util.Capitalize( objDR[ "Address_Line_1__c" ].ToString() );
					objDR[ "Address_Line_2__c" ] = Util.Capitalize( objDR[ "Address_Line_2__c" ].ToString() );
					objDR[ "City__c" ] = Util.Capitalize( objDR[ "City__c" ].ToString() );
				}

				// save CSV file
				//string strFileName = string.Concat( Company2SFUtils.strPath, "AMAJobApp.csv" ); // a million rows - better not save a file

				string strCSVContent = objDT.ToSalesForceCSVString();

				// create and submit a batch
				string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
				if( strBatchId.Contains( "ERROR" ) )
					return -1;

				iIterationNbr++;
			}

			ReportStatus( iCount.ToString(), " Total Job App rows processed. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished Job App CSV file. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );
			return iIterationNbr - 1;
		}

		public int CreateCompanyJobAppCSV( string strJobId )
		{

			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;

			ReportStatus( DateTime.Now.ToString(), " Starting refresh of Job Application. Loading data from Company.\r\n" );

			// get candidate records
			DataTable objCandidatesDT = objDB.GetDataTableFromSQLFile( "SQLAllCandidates.txt" );
			//// set the Recr ID as key for search purposes
			//objCandidatesDT.PrimaryKey	= new DataColumn[] { objCandidatesDT.Columns[ "Contact__c" ]
			//                                                , objCandidatesDT.Columns[ "Site_Code__c" ] };

			ReportStatus( "Loaded ", objCandidatesDT.Rows.Count.ToString(), " Candidate rows." );

			// retrieve all facilities
			List<Facility__c> objFacilities = objAPI.Query<Facility__c>(
				"select ID, Name, Sub_Region__r.Name, Sub_Region__r.Region_Code__r.Name, Site_Code__c from Facility__c order by Site_Code__c" );

			// roughly abc-abc+abc.abc @ abc-abc.abc-abc.abcd
			System.Text.RegularExpressions.Regex objRegExEmail = new System.Text.RegularExpressions.Regex(
						@"\w+([-+.]\w+)*@(\w+([-]\w+)*(\.\w+([-]\w+)*)*)+\.(edu|com|[A-Za-z]{2,4})" );

			// do the load in batches of 10000 rows each
			int iCount = 0, iIterationNbr = 1;
			bool bKeepLoading = true;
			string strLastRecrId = "78686";
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of Job App rows with RecrID > ", strLastRecrId );

				string strCondition = string.Concat( " AND p.ID > ", strLastRecrId, " " );
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllJobApplications.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more Job App rows found." );
					bKeepLoading = false;
					break;
				}

				// store the last MeNumber for the next iteration
				strLastRecrId = objDT.Rows[ objDT.Rows.Count - 1 ][ "RecruitingID__c" ].ToString();

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " Job App rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// validate emails (most common error)
				foreach( DataRow objDR in objDT.Rows )
				{
					// extract and validate email
					string strEmail = objDR[ "Email__c" ].ToString();
					if( strEmail.Equals( "" ) )
						continue;

					strEmail = strEmail.Replace( ",", "." ).Replace( "..", "." ).Replace( " ", "" )
							.Replace( ">", "." ).Replace( "@@", "@" );

					System.Text.RegularExpressions.MatchCollection objMatch = objRegExEmail.Matches( strEmail );
					if( objMatch.Count == 0 )
						objDR[ "Email__c" ] = "";
					else
						objDR[ "Email__c" ] = objMatch[ 0 ].Value;

					// process candidate records
					string strStage = "", strStatus = "", strRegion = "", strFTPT = "";
					DataRow[] objCandidates = objCandidatesDT.Select( 
										"Contact__c = '" + objDR[ "RecruitingID__c" ].ToString() + "'" );
					foreach( DataRow objCandRow in objCandidates )
					{
						// try match 1st 15 characters of facility
						string strName = objCandRow[ "HospitalName" ].ToString().Left( 15 );
						string strSiteCode = objCandRow[ "Facility_Name__c" ].ToString();
						Facility__c objFacility = FindFacilityByNameOrCode( objFacilities, strName, strSiteCode );
						if( objFacility == null )
						{
							tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find hospital matching "
												, strName, " or site code <", strSiteCode );
							continue;
						}

						Sub_Region__c objSubReg = objFacility.Sub_Region__r;
						Region__c objReg = objSubReg != null ? objSubReg.Region_Code__r : null;
						string strSubRegCode = "";
						
						if( objSubReg != null && objReg != null )
							strSubRegCode = objSubReg.Name.Equals( objReg.Name ) ?
												 objSubReg.Name : objReg.Name + " / " + objSubReg.Name;

						string strStatusCode = objCandRow[ "Candidate_Status__c" ].ToString();
						int iPosBeginParenthesis = strStatusCode.IndexOf( "(" );
						if( iPosBeginParenthesis >= 0 )
						{
							int iPosEndParenthesis = strStatusCode.IndexOf( ")" );
							strStatusCode = strStatusCode.Substring( iPosBeginParenthesis + 1
														, iPosEndParenthesis - iPosBeginParenthesis - 1 );
						}

						string strSearchStatus = strStatusCode;
						switch( strSearchStatus )
						{
							case "AW": strSearchStatus = "W"; break;
							case "AR": strSearchStatus = "P"; break;
							case "AC": strSearchStatus = "P"; break;
							case "SAP":
							case "SAR":
							case "SAH": strSearchStatus = "W"; break;
							default: strSearchStatus = "N"; break;
						}

						strSearchStatus = ";" + strSearchStatus + "-" + strSiteCode + "-" + strStatusCode + ";";

						if( ! strStage.Contains( objCandRow[ "Candidate_Stage__c" ].ToString() ) )
							strStage += ";" + objCandRow[ "Candidate_Stage__c" ].ToString() + ";";

						strStatus += strSearchStatus;

						if( ! strRegion.Contains( strSubRegCode ) )
							strRegion += ";" + strSubRegCode + ";";

						if( ! strFTPT.Contains( objCandRow[ "Full_Time_Part_Time__c" ].ToString() ) )
							strFTPT += ";" + objCandRow[ "Full_Time_Part_Time__c" ].ToString() + ";";
					}

					objDR[ "Provider_Name_Degree__c" ] =
						objDR[ "Last_Name__c" ].ToString() + ", " + objDR[ "First_Name__c" ].ToString() + ", " + objDR[ "Degree__c" ].ToString();

					objDR[ "Search_Candidate_Status__c" ] = strStatus.Left( 255 );
					objDR[ "Search_Candidate_Stage__c" ] = strStage.Left( 255 );
					objDR[ "Search_Region__c" ] = strRegion.Left( 255 );
					objDR[ "Search_FullTime_PartTime__c" ] = strFTPT.Left( 255 );
				}

				// remove physician number after done because it is a formula field 
				// and SalesForce errors out instead of politely ignoring it
				objDT.Columns.Remove( "Physician_Number__c" );

				// save CSV file
				string strFileName = string.Concat( Company2SFUtils.strPath, "JobApp", iIterationNbr.ToString(), ".csv" );

				objDT.SaveAsSalesForceCSV( strFileName );

				// load CSV file
				//string strFileName = string.Concat( Company2SFUtils.strPath, "JobApp1.csv" );
				StreamReader objSR = new StreamReader( strFileName );
				string strCSVContent = objSR.ReadToEnd();
				objSR.Close();

				// create and submit a batch
				string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
				if( strBatchId.Contains( "ERROR" ) )
					return -1;

				iIterationNbr++;
			}

			ReportStatus( iCount.ToString(), " Total Job App rows processed. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished Job App CSV file. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );
			return iIterationNbr - 1;
		}

		public void BulkDeleteTest1()
		{
			// create a job to DELETE a contact
			string strJobId = objAPI.RESTCreateJob( "delete", "Contact", "", "CSV" );
			if( strJobId.Contains( "ERROR" ) )
				return;

			// *** CSV file containing the contacts goes here ***
			string strCSVContent = @"Id
""003E0000006V9Tw"""; // delete john test

			// create and submit a batch
			string strBatchId = objAPI.RESTCreateBatch( strJobId, strCSVContent );
			if( strBatchId.Contains( "ERROR" ) )
				return;

			// verify the batch status
			string strBatchState = objAPI.RESTCheckBatch( strJobId, strBatchId );
			if( strBatchState.Contains( "Failed" ) )	// "Completed" if successful
				return;

			// close the job after submitting batches
			string strState = objAPI.RESTSetJobState( strJobId, "Closed" );	// to abort, send "Aborted"

			// get results if available
			string strResult = objAPI.RESTGetBatchResult( strJobId, strBatchId );

			tbStatus.Text = "Job Id: " + strJobId + " - Batch Id: " + strBatchId + " - State: " + strState + "\r\n" + strResult
				+ "\r\n" + strBatchState;
		}

		public class LeadSourcePair
		{
			public string LeadSource1 = "";
			public string LeadSource2 = "";
		}

		public static LeadSourcePair ConvertLeadSource( List<string> objLeadSources, List<string> obj2ndLeadSources
				, string[] strSearchFor, string[] strSubstitutions, string strLeadSource )
		{
			// normalize the lead source using the substitutions and convert to a standard value
			string strLeadSource1 = Company2SFUtils.ConvertToStandardValue( objLeadSources, strSearchFor, strSubstitutions, strLeadSource );
			string strLeadSource2 = Company2SFUtils.ConvertToStandardValue( obj2ndLeadSources, strSearchFor, strSubstitutions, strLeadSource );

			switch( strLeadSource1 )
			{
				case "Acquisition":
				case "Agency":
				case "Cold Call":
				case "Email Campaign":
				case "EmSource":
				case "Facebook":
				case "Mailing":
				case "Preceptor Program":
				case "Residency":
				case "Unknown":
				case "Word of Mouth":
					strLeadSource2 = strLeadSource1; break;

				case "Internet":
					strLeadSource2 = "Other"; break;

				case "Other":
					// attempt to populate lead source 1 deriving from lead source 2
					switch( strLeadSource2 )
					{
						case "Incentive Eligible":
						case "Incentive Ineligible":
						case "Non-Company":
						case "Recruiter":
							strLeadSource1 = "Referral";
							break;

						case "CareerMD":
						case "EMRA":
							strLeadSource1 = "Job Fair";
							break;

						case "ACEP/EMRA":
						case "Advance PA/NP":
						case "CareerBuilder":
						case "EDPhysician.com":
						case "Company Careers":
						case "FAPA":
						case "GasWork":
						case "HealtheCareers":
						case "hospitalistjobs":
						case "hospitalistworking":
						case "MDJobSite":
						case "NAPR":
						case "Physicianwork":
						case "PracticeLink":
							//case "Other":		is duplicate, so don't default
							//case "CareerMD":		is duplicate, so let it default to Job Fair
							strLeadSource1 = "Internet";
							break;

						case "ACEP":
						case "ACOEP":
						case "ACP-ASIM":
						case "AOA":
						case "GSACEP":
						case "NAIP":
						case "SHM":
							//case "Other":		is duplicate, so don't default
							strLeadSource1 = "Conference";
							break;

						case "Classified Ad":
						case "EM News":
						case "EP Monthly":
						case "IM":
						case "Other":
						case "PA World":
						case "WEM Annals":
							//case "Advance PA/NP":		is duplicate, so default to Internet
							strLeadSource1 = "Ad";
							break;

						default:
							strLeadSource1 = "Unknown";
							strLeadSource2 = strLeadSource;
							break;
					}
					break;
			}

			LeadSourcePair objReturnValues = new LeadSourcePair();
			objReturnValues.LeadSource1 = strLeadSource1;
			objReturnValues.LeadSource2 = strLeadSource2;

			return objReturnValues;
		}

		public void MassDataLoad()
		{
			bMassDataLoad = true;

			bool bSkipSubTypes = true;
			bool bSkipAgencies = true;
			bool bSkipCredentials = true;
			bool bSkipInstitutions = true;
			bool bSkipEducExp = true;
			bool bSkipReferences = true;
			bool bSkipResidencyCSV = true;
			int iCount = 0;
			int iTotalCount = 0;
			List<Credential_Subtype__c> objSubTypes = null;
			List<Credential_Agency__c> objAgencies = null;
			List<Credential__c> objCredentials = null;
			List<Reference__c> objReferences = null;
			List<Institution__c> objInstitutions = null;
			List<Education_or_Experience__c> objEducExp = null;
			List<Residency_Program__c> objResidPrograms = null;
			List<Resident__c> objResidents = null;

			// will skip account, settings, provider, notes, hierarchy and facilities
			DateTime dtLoadBegin = DateTime.Now;
			string strStatus = string.Concat( "** Mass Data Load started:  ", dtLoadBegin.ToString() );

			// retrieve Providers from SalesForce
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Contacts...\r\n" );
			List<Contact> objProviders = Company2SFUtils.GetProvidersFromSF( objAPI, lblError,
				" Owning_Region__c, Owning_Candidate_Stage__c, Owning_Candidate_Status__c, Birthdate "
				, null, " RecruitingID__c, Name " );
			Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objProviders, "Providers", tbStatus );

			// load references
			if( !bSkipReferences )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading References...\r\n" );
				objReferences = RefreshReferences( objProviders );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objReferences, "References", tbStatus );
			}

			// load agencies and credentials
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Agencies...\r\n" );
			if( !bSkipCredentials )
			{
				// load sub types
				if( !bSkipSubTypes && !bSkipCredentials )
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Subtypes...\r\n" );
					objSubTypes = RefreshSubtypes();
					Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objSubTypes, "SubTypes", tbStatus );
				}

				if( !bSkipAgencies )
					objAgencies = RefreshAgencies( objProviders );
				else
					objAgencies = objAPI.Query<Credential_Agency__c>(
"Select Id, OwnerId, IsDeleted, Name, CreatedDate, CreatedById, LastModifiedDate, LastModifiedById, SystemModstamp, Address1__c"
+ ", Address2__c, Dept__c, City__c, Zip__c, Phone__c, EXT__c, Fax__c, Contact__c, Title__c, Salutation__c, Credential_Type__c, "
+ "Code__c, Company_Agency_Match__c, State__c, Search_Exclude__c, State_Licensing_Agency__c, Search_Commonly_Used__c, "
+ "Metaphone_Address__c, Metaphone_City__c, Metaphone_Name__c from TableWithAgenciesGivingCredential" );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objAgencies, "Agencies", tbStatus );

				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Credentials...\r\n" );
				objCredentials = RefreshCredentials( objProviders, objAgencies, objSubTypes );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objCredentials, "Credentials", tbStatus );
			}

			// load institutions and education/experience (and residencies)
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Institutions...\r\n" );
			if( bSkipInstitutions )
				objInstitutions = objAPI.Query<Institution__c>(
"Select Id, OwnerId, IsDeleted, Name, CreatedDate, CreatedById, LastModifiedDate, LastModifiedById, SystemModstamp, "
+ "Address1__c, Address2__c, City__c, Code__c, State__c, Contact__c, Credential_Type__c, Dept__c, Company_Agency_Match__c, "
+ "EXT__c, Fax__c, Phone__c, Salutation__c, Title__c, Zip__c, Provider_Contract__c, Metaphone_Address__c, Metaphone_City__c, "
+ "Metaphone_Name__c, Acuity__c from Institution__c" );
			else
				objInstitutions = RefreshInstitutions( objProviders );
			Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objInstitutions, "Institutions", tbStatus );

			if( !bSkipEducExp )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Education/Experiences...\r\n" );
				objEducExp = RefreshEducationExperience( objProviders, objInstitutions );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objEducExp, "Education/Experiences", tbStatus );
			}

			// load candidates (it will query facilities by itself)
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Candidates...\r\n" );
			List<Candidate__c> objCandidates = RefreshCandidates( objProviders );
			Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objCandidates, "Candidates", tbStatus );

			if( !bSkipResidencyCSV )
			{
				// load residency data from CSV files
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Residency Programs...\r\n" );
				objResidPrograms = RefreshResidencyPrograms( objInstitutions );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objResidPrograms, "Residency Programs", tbStatus );

				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nLoading Residents...\r\n" );
				objResidents = RefreshResidents( objResidPrograms );
				Company2SFUtils.UpdateStatus( ref strStatus, ref iCount, ref iTotalCount, objResidents, "Residents", tbStatus );
			}

			TimeSpan tsDuration = DateTime.Now.Subtract( dtLoadBegin );
			strStatus = string.Concat( strStatus, "\r\n** Mass Data Load finished: \t", DateTime.Now.ToString() );
			strStatus = string.Concat( strStatus, "\r\n** Records loaded: \t", iTotalCount.ToString() );
			strStatus = string.Concat( strStatus, "\r\n** Duration: \t", tsDuration.Hours.ToString(), " hours and "
				, tsDuration.Minutes.ToString(), " minutes" );
			Company2SFUtils.SaveStatusCSV( strStatus, tbStatus );

			tbStatus.Text = string.Concat( tbStatus.Text, strStatus );
			bMassDataLoad = false;
		}

		public void InitializeAccountAndSettings()
		{

			Account objAcct = objAPI.QuerySingle<Account>( "select Id, Name from Account where Name = 'Company'" );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return;
			}

			if( objAcct == null )
			{
				objAcct = new Account();
				objAcct.Name = "Company";
				SaveResult[] objResults = objAPI.Insert( new Account[] { objAcct } );
				if( objResults != null )
					if( objResults[ 0 ].success )
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nAccount Company added." );
					else
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nError:  ", objResults[ 0 ].errors[ 0 ].message
										, " - ", objResults[ 0 ].errors[ 0 ].statusCode );
			}
			else
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nAccount Company already exists!" );

			List<Company_Hub_Settings__c> objSettings = objAPI.Query<Company_Hub_Settings__c>(
				"select Id, Name, Disable_Validation_Rules__c from Company_Hub_Settings__c" );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return;
			}

			if( objSettings == null || objSettings.Count == 0 )
			{
				// add settings for company
				Organization objOrg = objAPI.QuerySingle<Organization>( "select Id from Organization" );
				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				Company_Hub_Settings__c objOrgSettings = new Company_Hub_Settings__c();
				objOrgSettings.Disable_Validation_Rules__c = false;
				objOrgSettings.Disable_Validation_Rules__cSpecified = true;
				objOrgSettings.SetupOwnerId = objOrg.Id;
				objSettings.Add( objOrgSettings );

				// add settings for user
				string strUser = System.Configuration.ConfigurationManager.AppSettings[ "User" ].ToLower();
				User objUser = objAPI.QuerySingle<User>( string.Concat( "select Id from User where Name = '", strUser, "'" ) );
				if( !objAPI.ErrorMessage.Equals( "" ) )
				{
					lblError.Text = objAPI.ErrorMessage;
					return;
				}

				Company_Hub_Settings__c objUserSettings = new Company_Hub_Settings__c();
				objUserSettings.Disable_Validation_Rules__c = true;
				objUserSettings.Disable_Validation_Rules__cSpecified = true;
				objUserSettings.SetupOwnerId = objUser.Id;
				objSettings.Add( objUserSettings );

				SaveResult[] objSettingsResults = objAPI.Insert( objSettings.ToArray() );
				if( objSettingsResults != null )
				{
					if( objSettingsResults[ 0 ].success )
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nOrganization Settings added." );
					else
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nError:  ", objSettingsResults[ 0 ].errors[ 0 ].message
										, " - ", objSettingsResults[ 0 ].errors[ 0 ].statusCode );

					if( objSettingsResults[ 1 ].success )
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nUser Settings added." );
					else
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nError:  ", objSettingsResults[ 1 ].errors[ 1 ].message
										, " - ", objSettingsResults[ 1 ].errors[ 1 ].statusCode );
				}
			}
			else
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nSettings already exist!" );

			//if(objSettingsResults[ 0 ].errors != null)
			//    tbStatus.Text = string.Concat( tbStatus.Text, objSettingsResults[ 0 ].errors[ 0 ].message );

			//if(objSettingsResults[ 1 ].errors != null)
			//    tbStatus.Text = string.Concat( tbStatus.Text, objSettingsResults[ 1 ].errors[ 0 ].message );

		}

		public void UpdateUsersEmail( bool bDisplayOnly = true )
		{
			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError, bForUpdate: true );

			string strEmailList = "jane.doe@Company.com,john.doe@Company.com";
			strEmailList = strEmailList.ToLower();

			List<EMSC2SF.User> objUsersToUpdate	= new List<User>();

			foreach( EMSC2SF.User objUser in objExistingUsers )
			{
				if( objUser.IsActive != null && objUser.Email != null && !(bool) objUser.IsActive )
					continue;

				// remove suffixes
				string strEmail = objUser.Email.Replace( ".test1", "" ).Replace( ".dev1", "" );

				// check if user is in the list
				if( !strEmailList.Contains( strEmail ) )
					continue;

				// change email
				objUser.Email = strEmail;

				// separate the user record for update
				objUsersToUpdate.Add( objUser );
			}

			// update only the users who have changed
			UpsertResult[] objUsersResults = null;
			if( !bDisplayOnly )
				objUsersResults = objAPI.Upsert( "Id", objUsersToUpdate.ToArray<sObject>() );

			// create CSV file / set the Ids in the list of users
			Company2SFUtils.SetIdsReportErrors( objUsersToUpdate, objUsersResults, tbStatus );

			return;
		}

		public List<Contact> RefreshAMAProviders( bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;

			ReportStatus( DateTime.Now.ToString(), " Starting refresh of AMA Providers. Loading existing providers from EmForce.\r\n" );

			// retrieve SalesForce provider records to perform duplicate detection during the loop
			// (only retrieve non-AMA records)
			List<Contact> objExistingContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
					, " HomePhone, Birthdate, OtherStreet, Address_Line_1__c, MeNumber__c "
					, " AMAOnly__c <> '1' " );

			ReportStatus( objExistingContacts.Count.ToString(), " AMA provider rows retrieved.\r\n" );

			Account objAcct = objAPI.QuerySingle<Account>( "select Id, Name from Account where Name = 'Company'" );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return null;
			}

			// get the last creation date of AMA providers in order to only bring recent AMA providers
			List<Contact> objLastAMA = objAPI.Query<Contact>(
				"select Id, CreatedDate from Contact where AMAOnly__c <> '1' order by CreatedDate DESC limit 1" );
			string strLastAMADt = DateTime.Today.ToShortDateString();
			if( objLastAMA.Count == 1 )
				strLastAMADt = ( (DateTime) objLastAMA[ 0 ].CreatedDate ).ToShortDateString();

			//strLastAMADt = "1/1/2000";

			// do the load in batches of 10000 rows each			
			int iCount = 0;
			bool bKeepLoading = true;
			string strLastMeNumber = "0";
			List<Contact> objContacts = new List<Contact>( 10000 );
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of AMA providers with MeNumber > ", strLastMeNumber );

				string strCondition = string.Concat(
							" AND p.CompanyUpdateDate > '", strLastAMADt, "' AND p.MeNumber > '", strLastMeNumber, "' " );
				objDT = objDB.GetDataTableFromSQLFile( "SQL_AMA_Providers.txt", null, strCondition );

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );
					//tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more AMA provider rows found." );
					bKeepLoading = false;
					break;
				}

				// store the last MeNumber for the next iteration
				strLastMeNumber = objDT.Rows[ objDT.Rows.Count - 1 ][ "MeNumber__c" ].ToString();

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " AMA provider rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// load providers from datatable to contact list
				objContacts = new List<Contact>( objDT.Rows.Count );
				foreach( DataRow objDR in objDT.Rows )
				{
					// copy all datatable columns to contact object attributes
					Contact objNewContact = objDR.ConvertTo<Contact>( true );

					// capitalize initials
					objNewContact.FirstName = Util.Capitalize( objNewContact.FirstName );
					objNewContact.LastName = Util.Capitalize( objNewContact.LastName );
					objNewContact.OtherStreet = Util.Capitalize( objNewContact.OtherStreet );
					objNewContact.Address_Line_1__c = Util.Capitalize( objNewContact.Address_Line_1__c );
					objNewContact.Address_Line_2__c = Util.Capitalize( objNewContact.Address_Line_2__c );
					objNewContact.City__c = Util.Capitalize( objNewContact.City__c );
					objNewContact.OtherCity = Util.Capitalize( objNewContact.OtherCity );

					// find duplicate record in the existing EmForce records
					Contact objContactFound = objExistingContacts.FirstOrDefault( c =>
											c.IsAMADuplicateOf( objNewContact ) );
					if( objContactFound != null )
					{
						ReportStatus( "* Skipping AMA duplicate for provider ", objNewContact.FirstName
								, " ", objNewContact.LastName
								, " - MeNumber = ", objNewContact.MeNumber__c
								, " (duplicate RecrID ", objContactFound.RecruitingID__c.ToString(), ")" );

						// keep the 1st record entered or the Company record and skip this one
						continue;
					}

					// capitalize properly for cases such as "Macon Ga" and "N Kansas City Al"
					// (should become Macon GA and N Kansas City AL)
					objNewContact.Birthplace__c = Util.CapitalizeWithStateCode( objNewContact.Birthplace__c );

					// add metaphone values
					string strName = string.Concat( objNewContact.FirstName, " ", objNewContact.LastName );

					objNewContact.Metaphone_Name__c = strName.ToNormalizedMetaphone();
					objNewContact.Metaphone_Address__c = objNewContact.Address_Line_1__c.ToNormalizedMetaphone().Left( 50 );
					objNewContact.Metaphone_City__c = objNewContact.City__c.ToNormalizedMetaphone();

					// fix time zone bug
					objNewContact.Birthdate = Company2SFUtils.FixTimeZoneBug( objNewContact.Birthdate );

					// for some reason, this is required for the upsert to work
					objNewContact.RecruitingID__cSpecified = true;
					objNewContact.PhysicianNumber__cSpecified = true;
					objNewContact.CreatedDateSpecified = true;
					objNewContact.Active_Military__cSpecified = true;
					objNewContact.Deceased__cSpecified = true;
					objNewContact.Directorship_Experience__cSpecified = true;
					objNewContact.Drivers_License_Expiration_Date__cSpecified = true;
					objNewContact.BirthdateSpecified = ( objNewContact.Birthdate != null );

					objNewContact.AccountId = objAcct.Id;

					objContacts.Add( objNewContact );

					ReportStatus( "Processed contact ", objNewContact.FirstName, " ", objNewContact.LastName
						, " (", objContacts.Count.ToString(), "/", objDT.Rows.Count.ToString(), ")" );
				}

				// use MeNumber to upsert existing AMA records in EmForce
				UpsertResult[] objResults = null;
				if( !bDisplayOnly )
					objResults = objAPI.Upsert( "MeNumber__c", objContacts.ToArray<sObject>() );

				Company2SFUtils.ReportErrorsToHistoryFile( objResults, objContacts );
			}

			ReportStatus( iCount.ToString(), " Total AMA provider rows loaded. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished loading AMA contacts. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );

			return objContacts;
		}

		public List<Education_or_Experience__c> RefreshAMAProviderCredentials( bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;

			ReportStatus( DateTime.Now.ToString()
				, " Starting refresh of AMA Provider Credentials.\r\n" );

			// get record types and the ids for Education and Board credential
			List<RecordType> objRecTypes = objAPI.Query<RecordType>( "select Id, Name from RecordType" );
			string strEducationTypeId = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Degree/ Education" ).Id;
			string strBoardTypeId = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Board" ).Id;

			// bring Board subtypes for credential
			List<Credential_Subtype__c> objSubTypes = objAPI.Query<Credential_Subtype__c>(
				"select Id, Name from Credential_Subtype__c where Credential_Type__c = 'Board' " );

			// get the last creation date of AMA providers in order to only bring recent AMA providers
			List<Contact> objLastAMA = objAPI.Query<Contact>(
				"select Id, CreatedDate from Contact where AMAOnly__c <> '1' order by CreatedDate DESC limit 1" );
			string strLastAMADt = DateTime.Today.ToShortDateString();
			if( objLastAMA.Count == 1 )
				strLastAMADt = ( (DateTime) objLastAMA[ 0 ].CreatedDate ).ToShortDateString();


			strLastAMADt = "1/1/2001";



			// load institutions list
			List<Institution__c> objInstitutions = objAPI.Query<Institution__c>(
"select Id, Name, Company_Agency_Match__c, City__c, Metaphone_Name__c, Metaphone_City__c from Institution__c where Metaphone_City__c != null and Metaphone_Name__c != null order by Metaphone_City__c, Metaphone_Name__c" );

			ReportStatus( "Loaded ", objInstitutions.Count.ToString(), " institutions." );

			// do the load in batches of 10000 rows each
			int iCount = 0, iSkipped = 0;
			bool bKeepLoading = true;
			string strLastMeNumber = "0";
			List<Education_or_Experience__c> objEducations = null;
			List<Credential__c> objCredentials = null;
			while( bKeepLoading )
			{
				// get next batch of rows
				ReportStatus( "Loading next batch of AMA providers with MeNumber > ", strLastMeNumber );

				string strCondition = string.Concat(
							" AND p.CompanyUpdateDate > '", strLastAMADt, "' AND p.MeNumber > '", strLastMeNumber, "' " );
				objDT = objDB.GetDataTableFromSQLFile( "SQL_AMA_Provider_Credentials.txt", null, strCondition );

				//strCondition = string.Concat(
				//            " AMAOnly__c = '1' AND MeNumber__c > '", strLastMeNumber, "' " );
				//List<Contact> objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError, "MeNumber__c"
				//                , strCondition, " MeNumber__c limit 10000 " );	// not sure this will bring the right records

				if( !objDB.ErrorMessage.Equals( "" ) )
				{
					ReportStatus( objDB.ErrorMessage );
					//tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );

					// interrupt process
					bKeepLoading = false;
					break;
				}

				if( objDT.Rows.Count == 0 )
				{
					// interrupt process because there are no more rows
					ReportStatus( "No more AMA provider rows found." );
					bKeepLoading = false;
					break;
				}

				// store the last MeNumber for the next iteration
				strLastMeNumber = objDT.Rows[ objDT.Rows.Count - 1 ][ "MeNumber__c" ].ToString();

				iCount += objDT.Rows.Count;
				ReportStatus( objDT.Rows.Count.ToString(), " AMA provider rows retrieved. Total ", iCount.ToString(), " loaded.\r\n" );

				// load providers from datatable to contact list
				objEducations = new List<Education_or_Experience__c>( objDT.Rows.Count );
				objCredentials = new List<Credential__c>( objDT.Rows.Count );
				foreach( DataRow objDR in objDT.Rows )
				{
					// capitalize initials
					string strFirstName = Util.Capitalize( objDR[ "FirstName" ].ToString() );
					string strLastName = Util.Capitalize( objDR[ "LastName" ].ToString() );
					string strMeNumber = objDR[ "MeNumber__c" ].ToString();

					// add Board Certifications if present
					string strBoardCert = objDR[ "BoardCert" ] != null ? objDR[ "BoardCert" ].ToString() : "";
					if( !strBoardCert.Equals( "" ) )
						iSkipped = AddBoardCertification( strBoardTypeId, objSubTypes, iSkipped, objCredentials, strFirstName, strLastName, strMeNumber, strBoardCert );
					strBoardCert = objDR[ "BoardCert1" ] != null ? objDR[ "BoardCert1" ].ToString() : "";
					if( !strBoardCert.Equals( "" ) )
						iSkipped = AddBoardCertification( strBoardTypeId, objSubTypes, iSkipped, objCredentials, strFirstName, strLastName, strMeNumber, strBoardCert );
					strBoardCert = objDR[ "BoardCert2" ] != null ? objDR[ "BoardCert2" ].ToString() : "";
					if( !strBoardCert.Equals( "" ) )
						iSkipped = AddBoardCertification( strBoardTypeId, objSubTypes, iSkipped, objCredentials, strFirstName, strLastName, strMeNumber, strBoardCert );

					Education_or_Experience__c objNewEducation = new Education_or_Experience__c();

					string strInstitution = objDR[ "GraduateEducationInstitution" ].ToString();
					string strCity = objDR[ "InstitutionCity" ].ToString().Trim();

					string strInstitNormalized = Util.NormalizeAMAInstitution( strInstitution );

					string strMetaphoneInstitution = strInstitution.ToMetaphone();
					string strMetaphoneCity = strCity.ToMetaphone();

					// convert institution code into lookup id
					Institution__c objInstitutionFound = objInstitutions.FirstOrDefault(
								i => ( i.City__c.Equals( strCity )
											|| i.Metaphone_City__c.Equals( strMetaphoneCity ) )
									&& ( i.Name.IsEqualOrPartiallyMatchedTo( strInstitNormalized )
											|| i.Metaphone_Name__c.Equals( strMetaphoneInstitution ) ) );
					//Institution__c objInstitutionFound = objInstitutions.FirstOrDefault( 
					//                        i => ( i.Name.IsMetaphoneMatchedTo( strInstitution )
					//                                || i.Name.IsEqualOrPartiallyMatchedTo( strInstitNormalized ) )
					//                        && i.City__c.IsMetaphoneMatchedTo( strCity ) );
					if( objInstitutionFound != null )
						objNewEducation.Institution__c = objInstitutionFound.Id;
					else
					{
						// create the institution
						objInstitutionFound = new Institution__c();
						objInstitutionFound.Name = strInstitution;
						objInstitutionFound.Metaphone_Name__c = strInstitution.ToNormalizedMetaphone();
						objInstitutionFound.City__c = strCity;
						objInstitutionFound.Metaphone_City__c = strCity.ToNormalizedMetaphone();
						objInstitutionFound.State__c = objDR[ "InstitutionState" ].ToString().Trim();
						objInstitutionFound.Credential_Type__c = "Institution";
						objInstitutionFound.Code__c = objDR[ "InstitutionID" ].ToString().Trim();

						// use MeNumber to upsert existing AMA records in EmForce
						SaveResult[] objSaveResult = null;
						if( !bDisplayOnly )
							objSaveResult = objAPI.Insert( new sObject[] { objInstitutionFound } );

						if( objSaveResult[ 0 ].errors != null )
						{
							ReportStatus( "Could not create institution ", strInstitution,
							   " for the credentials of provider ", strFirstName, " ", strLastName
							   , " - ", objSaveResult[ 0 ].errors[ 0 ].message );
							iSkipped++;
							continue;
						}

						objInstitutionFound.Id = objSaveResult[ 0 ].id;
						objNewEducation.Institution__c = objInstitutionFound.Id;

						objInstitutions.Add( objInstitutionFound );
						ReportStatus( "Created institution ", strInstitution, " with id ", objNewEducation.Institution__c
							, " for the credentials of provider ", strFirstName, " ", strLastName );
					}

					//Contact objProvider = objContacts.FirstOrDefault( i => i.MeNumber__c.Equals( strContactId ) );
					//if( objProvider != null )
					//    objNewEducation.Contact__c = objProvider.Id;
					//else // skip if provider was not found
					//{
					//    ReportStatus( "Could not find ME Number ", strContactId,
					//        " for the credentials of provider ", strFirstName, " ", strLastName );
					//    iSkipped++;
					//    continue;
					//}

					// create a contact object to let SalesForce do the relationship
					Contact objEducProvider = new Contact();
					objEducProvider.MeNumber__c = strMeNumber;
					objNewEducation.Contact__r = objEducProvider;

					if( !Convert.IsDBNull( objDR[ "GraduationFromYear" ] ) )
					{
						DateTime dtValue = new DateTime( Convert.ToInt16( objDR[ "GraduationFromYear" ] ), 1, 1 );
						objNewEducation.From__c = dtValue;
						// fix time zone bug
						objNewEducation.From__c = Company2SFUtils.FixTimeZoneBug( objNewEducation.From__c );
					}
					if( !Convert.IsDBNull( objDR[ "GraduationToYear" ] ) )
					{
						DateTime dtValue = new DateTime( Convert.ToInt16( objDR[ "GraduationToYear" ] ), 1, 1 );
						objNewEducation.To__c = dtValue;
						// fix time zone bug
						objNewEducation.To__c = Company2SFUtils.FixTimeZoneBug( objNewEducation.To__c );
					}

					objNewEducation.Name = string.Concat( strFirstName, " ", strLastName
										, "-Medical School at ", strInstitution ).Left( 80 );

					objNewEducation.Description__c = "Degree/ Education";
					objNewEducation.Type__c = "Medical School";
					objNewEducation.RecordTypeId = strEducationTypeId;

					objNewEducation.From__cSpecified = ( objNewEducation.From__c != null );
					objNewEducation.To__cSpecified = ( objNewEducation.To__c != null );

					objEducations.Add( objNewEducation );

					ReportStatus( "Processed graduate education ", objNewEducation.Name
						, " (", objEducations.Count.ToString(), "/", objDT.Rows.Count.ToString(), ")" );
				}

				// use MeNumber to upsert existing AMA records in EmForce
				UpsertResult[] objResults = null;
				if( !bDisplayOnly )
				{
					objResults = objAPI.Upsert( "Name", objEducations.ToArray<sObject>() );

					Company2SFUtils.ReportErrorsToHistoryFile( objResults, objEducations );

					objResults = null;
					objResults = objAPI.Upsert( "Name", objCredentials.ToArray<sObject>() );

					Company2SFUtils.ReportErrorsToHistoryFile( objResults, objCredentials );
				}
			}

			ReportStatus( iCount.ToString(), " Total AMA credential rows loaded. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished loading AMA credentials. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );

			return objEducations;
		}

		public int AddBoardCertification( string strBoardTypeId, List<Credential_Subtype__c> objSubTypes
			, int iSkipped, List<Credential__c> objCredentials, string strFirstName, string strLastName
			, string strMeNumber, string strBoardCert )
		{
			Credential__c objNewCredential = new Credential__c();
			Credential_Subtype__c objSubType = objSubTypes.FirstOrDefault( s => s.Name.Equals( strBoardCert ) );
			if( objSubType != null )
			{
				// if found subtype, add credential to be upserted
				objNewCredential.Credential_Sub_Type__c = objSubType.Id;
				objNewCredential.RecordTypeId = strBoardTypeId;
				objNewCredential.Name = string.Concat( strFirstName, " ", strLastName
										, "-", strBoardCert, " Board Certification" ).Left( 80 );
				objNewCredential.Comments__c = "Imported from AMA data.";
				objNewCredential.Certification_Type__c = "General";

				// create a contact object to let SalesForce do the relationship
				Contact objProvider = new Contact();
				objProvider.MeNumber__c = strMeNumber;
				objNewCredential.Contact__r = objProvider;

				objCredentials.Add( objNewCredential );

				ReportStatus( "Created Board Certification credential ", objNewCredential.Name );
			}
			else
			{
				ReportStatus( "Could not find credential ", strBoardCert,
				   " for the provider ", strFirstName, " ", strLastName );
				iSkipped++;
			}
			return iSkipped;
		}

		public List<Contact> RefreshProviders( bool bImportNotes = true, bool bImportAllRecords = false, bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();
			System.Diagnostics.Stopwatch objWatchOverall = new System.Diagnostics.Stopwatch();
			objWatchOverall.Start();

			DataTable objDT = null;
			if( bImportAllRecords )
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllProviders.txt" );
			else
				objDT = objDB.GetDataTableFromSQLFile( "SQLProvider.txt" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				ReportStatus( "\r\n", objDB.ErrorMessage );
				return null;
			}

			ReportStatus( objDT.Rows.Count.ToString(), " provider rows retrieved.\r\n" );

			System.Text.RegularExpressions.Regex objRegExZip = new System.Text.RegularExpressions.Regex(
						@"\d{5}([ -]\d{4})*" ); // roughly 12345-1234

			// roughly abc-abc+abc.abc @ abc-abc.abc-abc.abcd
			System.Text.RegularExpressions.Regex objRegExEmail = new System.Text.RegularExpressions.Regex(
						@"\w+([-+.]\w+)*@(\w+([-]\w+)*(\.\w+([-]\w+)*)*)+\.(edu|com|[A-Za-z]{2,4})" );
			//old:  @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" );

			System.Text.RegularExpressions.Regex objRegExPhone = new System.Text.RegularExpressions.Regex(
						@"([\(]*\d*[\)]*[ .-])*\d*[ .-]\d*([ ]*\w{0,4}[ ]?\d+)?" ); // roughly (1234)-1234-1234 x123

			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			List<EMSC2SF.Sub_Region__c> objSubRegions = objAPI.Query<EMSC2SF.Sub_Region__c>(
									"select ID, Name, Recruiting_Manager__c, Sub_Region_Description__c from Sub_Region__c " );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return null;
			}

			Account objAcct = objAPI.QuerySingle<Account>( "select Id, Name from Account where Name = 'Company'" );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return null;
			}

			// standard specialties to which to compare the imported specialty
			List<string> objSpecialties = new List<string>( new string[] { "Anesthesiology", "Critical Care"
					, "Emergency Medicine", "Family Practice", "General Surgery", "Internal Medicine"
					, "Occupational Medicine", "Other", "Pediatrics", "Psychiatry", "Radiology" } );
			// keywords to replace in specialty
			string[] strSpecialtySearchFor = { " pediatric ", " surgery ", " primary care ", " med "
												, " urgent care ", " wound care ", " fp "
												, " im " };
			string[] strSpecialtySubstitutions = { " pediatrics ", " general surgery ", " family practice ", " medicine "
												, " emergency medicine ", " emergency medicine ", " family practice "
												, " internal medicine " };

			// standard lead sources to which to compare the imported sources
			List<string> objLeadSources = new List<string>( new string[] { 
				"Acquisition", "Ad", "Agency", "Cold Call", "Conference", "Email Campaign", "EmSource", "Facebook", "Internet"
				, "Job Fair", "Mailing", "Preceptor Program", "Referral", "Residency", "Unknown", "Word of Mouth" } );
			List<string> obj2ndLeadSources = new List<string>( new string[] { 				
				"Acquisition", "Agency", "Cold Call", "Email Campaign", "EmSource", "Facebook", "Mailing", "Preceptor Program"
				, "Residency", "Unknown", "Word of Mouth" 
				, "Advance PA/NP", "Classified Ad", "EM News", "EP Monthly", "IM", "Other", "PA World", "WEM Annals", "ACEP"
				, "ACOEP", "ACP-ASIM", "AOA", "GSACEP", "NAIP", "Other", "SHM", "ACEP/EMRA", "Advance PA/NP", "CareerBuilder"
				, "CareerMD", "EDPhysician.com", "Company Careers", "FAPA", "GasWork", "HealtheCareers", "hospitalistjobs"
				, "hospitalistworking", "MDJobSite", "NAPR", "Other", "Physicianwork", "PracticeLink", "CareerMD", "EMRA"
				, "Incentive Eligible", "Incentive Ineligible", "Non-Company", "Recruiter" } );

			// keywords to replace in lead source
			string[] strSearchFor = { " journal ", " at hospital ", " career ", " commercial ", " referral-hospital "
									, " not incentive eligible ", " advertisement ", " email/fax ", " annals "
									, " web page ", " meeting ", " website ", " newspaper " };

			// keywords that will replace the above in lead source
			string[] strSubstitutions = { " ad ", " acquisition ", " job ", " campaign ", " acquisition "
										, " incentive ineligible ", " ad ", " email ", " conference "
										, " internet ", " conference ", " internet ", " ad " };

			// load providers from datatable to contact list
			List<Contact> objContacts = new List<Contact>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				Contact objNewContact = ValidateProviderContact( objDT, objRegExZip, objRegExEmail, objRegExPhone, objExistingUsers, objSubRegions, objSpecialties, strSpecialtySearchFor, strSpecialtySubstitutions, objLeadSources, obj2ndLeadSources, strSearchFor, strSubstitutions, objContacts, objDR );

				objNewContact.AccountId = objAcct.Id;

				if( objNewContact != null )
					objContacts.Add( objNewContact );
			}

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "RecruitingID__c", objContacts.ToArray<sObject>() );

			Company2SFUtils.ReportErrorsToHistoryFile( objResults, objContacts );

			ReportStatus( objContacts.Count.ToString(), " Total providers loaded. \r\n" );

			//ReportStatus( "Reporting errors for contacts" );

			//// create CSV file / set the Ids in the list of facilities
			//Company2SFUtils.SetIdsReportErrors( objContacts, objResults, tbStatus );
			objResults = null;

			// populate the Provider record name (because it is read only for Contacts and we don't get it in UpsertResults)
			foreach( Contact objProv in objContacts )
				objProv.Name = string.Concat( objProv.FirstName, " ", objProv.LastName );

			//ShowData(objContacts);

			// if we don't need to import notes/comments, return immediately
			if( !bImportNotes )
				return objContacts;

			objWatch.Stop();
			ReportStatus( "Finished loading contacts. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours ", objWatch.Elapsed.Minutes.ToString(), " minutes." );

			RefreshProviderNotes( objContacts );

			objWatchOverall.Stop();
			ReportStatus( "Finished loading contacts + notes. Duration: "
							, objWatchOverall.Elapsed.Hours.ToString(), " hours "
							, objWatchOverall.Elapsed.Minutes.ToString(), " minutes." );

			return objContacts;
		}

		public Contact ValidateProviderContact( DataTable objDT, System.Text.RegularExpressions.Regex objRegExZip, System.Text.RegularExpressions.Regex objRegExEmail, System.Text.RegularExpressions.Regex objRegExPhone, List<EMSC2SF.User> objExistingUsers, List<EMSC2SF.Sub_Region__c> objSubRegions, List<string> objSpecialties, string[] strSpecialtySearchFor, string[] strSpecialtySubstitutions, List<string> objLeadSources, List<string> obj2ndLeadSources, string[] strSearchFor, string[] strSubstitutions, List<Contact> objContacts, DataRow objDR )
		{
			// copy all datatable columns to contact object attributes
			Contact objNewContact = objDR.ConvertTo<Contact>( true );

			objNewContact.Middle_Name__c = objNewContact.Middle_Name__c.RemoveSpecialCharacters();

			// validate zip
			System.Text.RegularExpressions.MatchCollection objMatchZip = objRegExZip.Matches( objNewContact.Zipcode__c );
			if( objMatchZip.Count > 0 )
				objNewContact.Zipcode__c = objMatchZip[ 0 ].Value;

			objNewContact.Zipcode__c = Company2SFUtils.ValidatePostalCode( objNewContact.Zipcode__c );
			objNewContact.OtherPostalCode = Company2SFUtils.ValidatePostalCode( objNewContact.OtherPostalCode );

			// validate ssn
			if( objNewContact.SSN__c != null && objNewContact.SSN__c.Length != 9 )
				objNewContact.SSN__c = "";

			// extract and validate email
			string strUnformattedEmail = objDR[ "UnformattedEmail" ].ToString();
			strUnformattedEmail = strUnformattedEmail.Replace( ",", "." ).Replace( "..", "." ).Replace( " ", "" )
					.Replace( ">", "." ).Replace( "@@", "@" );
			System.Text.RegularExpressions.MatchCollection objMatch = objRegExEmail.Matches( strUnformattedEmail );
			if( objMatch.Count > 0 )
				objNewContact.Email = objMatch[ 0 ].Value;
			if( objMatch.Count > 1 )
				objNewContact.Contacts_Email_No_2__c = objMatch[ 1 ].Value;
			if( objMatch.Count > 2 )
				objNewContact.Contacts_Email_No_3__c = objMatch[ 2 ].Value;
			if( objMatch.Count == 0 )
				objNewContact.Email = "";

			// validate phones
			objNewContact.HomePhone = Company2SFUtils.ValidPhone( objRegExPhone, objNewContact.HomePhone );
			objNewContact.OtherPhone = Company2SFUtils.ValidPhone( objRegExPhone, objNewContact.OtherPhone );
			objNewContact.Work_Phone__c = Company2SFUtils.ValidPhone( objRegExPhone, objNewContact.Work_Phone__c );
			objNewContact.MobilePhone = Company2SFUtils.ValidPhone( objRegExPhone, objNewContact.MobilePhone );
			objNewContact.Fax = Company2SFUtils.ValidPhone( objRegExPhone, objNewContact.Fax );

			// set the owning region using appointments, if blank, use the region in the provider's record
			string strOwningRegion = objDR[ "Default_Owning_Region" ].ToString();
			switch( strOwningRegion )
			{
				case "CEN":
				case "6":
				case "14":
				case "2":
					strOwningRegion = "SWED";
					break;
				case "SE":
				case "7":
					strOwningRegion = "SED";
					break;
				case "NE":
					strOwningRegion = "NED";
					break;
				case "PW":
					strOwningRegion = "PWED";
					break;
				case "DMS":
					strOwningRegion = "PWDMS";
					break;
				case "SANJ":
					strOwningRegion = "PWSANJ";
					break;
				case "4":
				case "1":
				case "9":
				case "11":
				case "13":
					strOwningRegion = "";
					break;
				case "":
					strOwningRegion = objDR[ "Less_Reliable_Region" ].ToString();
					break;
				default: // don't change it
					break;
			}
			objNewContact.Owning_Region__c = strOwningRegion;

			// convert specialty
			objNewContact.Specialty__c = Company2SFUtils.ConvertToStandardValue( objSpecialties, strSpecialtySearchFor, strSpecialtySubstitutions
				, objNewContact.Specialty__c );

			// convert lead source
			LeadSourcePair objLSResults = ConvertLeadSource( objLeadSources, obj2ndLeadSources
				, strSearchFor, strSubstitutions, objNewContact.Lead_Source_1__c );
			objNewContact.Lead_Source_1__c = objLSResults.LeadSource1;
			objNewContact.Lead_Source_Text__c = objLSResults.LeadSource2;

			// convert recruiter usercode to Salesforce User ID
			if( !objDR[ "Recruiter" ].IsNullOrBlank() )
			{
				string strRecruiter = objDR[ "Recruiter" ].ToString();
				EMSC2SF.User objFound = objExistingUsers.FirstOrDefault(
							u => u.Company_Usercode__c != null
								&& u.Company_Usercode__c.Equals( strRecruiter )
								&& u.IsActive == true );
				if( objFound != null )
					objNewContact.OwnerId = objFound.Id;
				else
				{
					// find the recruiting manager for the region
					Sub_Region__c objFoundSR = objSubRegions.FirstOrDefault( i => i.Name.Equals( strOwningRegion ) );
					if( objFoundSR != null )
					{
						// check whether recruiting manager/user is active
						objFound = objExistingUsers.FirstOrDefault(
									u => u.Id.Equals( objFoundSR.Recruiting_Manager__c )
										&& u.IsActive == true );
						if( objFound != null )
							objNewContact.OwnerId = objFound.Id;
					}

					// many providers don't have a region (or the region manager is inactive)
					// so the owner will be whoever does the data load
				}
			}

			// fix time zone bug
			objNewContact.Birthdate = Company2SFUtils.FixTimeZoneBug( objNewContact.Birthdate );
			objNewContact.Drivers_License_Expiration_Date__c =
				Company2SFUtils.FixTimeZoneBug( objNewContact.Drivers_License_Expiration_Date__c );

			// for some reason, this is required for the upsert to work
			objNewContact.RecruitingID__cSpecified = true;
			objNewContact.PhysicianNumber__cSpecified = true;
			objNewContact.CreatedDateSpecified = true;
			objNewContact.Active_Military__cSpecified = true;
			objNewContact.Deceased__cSpecified = true;
			objNewContact.Directorship_Experience__cSpecified = true;
			objNewContact.Drivers_License_Expiration_Date__cSpecified = true;
			objNewContact.BirthdateSpecified = ( objNewContact.Birthdate != null );

			// copy other address into main address if main is empty
			if( objNewContact.Address_Line_1__c.IsNullOrBlank() & objNewContact.City__c.IsNullOrBlank()
				&& !objNewContact.OtherStreet.IsNullOrBlank() )
			{
				objNewContact.Address_Line_1__c = objNewContact.OtherStreet;
				objNewContact.Address_Line_2__c = "";
				objNewContact.City__c = objNewContact.OtherCity;
				objNewContact.State__c = objNewContact.OtherState;
				objNewContact.Zipcode__c = objNewContact.OtherPostalCode;

				objNewContact.OtherStreet = "";
				objNewContact.OtherCity = "";
				objNewContact.OtherState = "";
				objNewContact.OtherPostalCode = "";
			}

			// add metaphone values
			string strName = objNewContact.FirstName;
			if( objNewContact.Middle_Name__c.IsNullOrBlank() )
				strName = string.Concat( strName, " ", objNewContact.LastName );
			else
				strName = string.Concat( strName, " ", objNewContact.Middle_Name__c, " ", objNewContact.LastName );

			objNewContact.Metaphone_Name__c = strName.ToNormalizedMetaphone();
			objNewContact.Metaphone_Address__c = objNewContact.Address_Line_1__c.ToNormalizedMetaphone().Left( 50 );
			objNewContact.Metaphone_City__c = objNewContact.City__c.ToNormalizedMetaphone();

			// find duplicate record in the list by matching metaphone name, phone and email
			Contact objContactFound = objContacts.FirstOrDefault( c =>
									c.IsADuplicateOf( objNewContact ) );

			if( objContactFound != null )
			{
				ReportStatus( "\r\nFound duplicate for provider ", objNewContact.FirstName, " ", objNewContact.LastName
						, " - RecrID = ", objNewContact.RecruitingID__c.ToString() );

				// if found record doesn't have a physician number and the current one does
				// then skip the duplicate instead of adding it to the list
				if( objNewContact.PhysicianNumber__c == null
					|| ( !objContactFound.PhysicianNumber__c.Equals( 0.0 )
						&& objNewContact.PhysicianNumber__c.Equals( 0.0 ) ) )
					return null;

				// check whether the found record is better than the current, if it is, skip the current (don't add)
				if( objNewContact.SSN__c.IsNullOrBlank() && !objContactFound.SSN__c.IsNullOrBlank() )
					return null;
				if( objNewContact.Email.IsNullOrBlank() && !objContactFound.Email.IsNullOrBlank() )
					return null;
				if( objNewContact.Birthdate.IsNullOrBlank() && !objContactFound.Birthdate.IsNullOrBlank() )
					return null;
				if( objNewContact.Address_Line_1__c.IsNullOrBlank() && !objContactFound.Address_Line_1__c.IsNullOrBlank() )
					return null;
				if( objNewContact.OtherStreet.IsNullOrBlank() && !objContactFound.OtherStreet.IsNullOrBlank() )
					return null;
				if( objNewContact.HomePhone.IsNullOrBlank() && !objContactFound.HomePhone.IsNullOrBlank() )
					return null;
				if( objNewContact.Work_Phone__c.IsNullOrBlank() && !objContactFound.Work_Phone__c.IsNullOrBlank() )
					return null;

				ReportStatus( "\r\n* Removing duplicate for provider ", objNewContact.FirstName, " ", objNewContact.LastName
						, " - RecrID = ", objNewContact.RecruitingID__c.ToString()
						, " (duplicate RecrID ", objContactFound.RecruitingID__c.ToString(), ")" );

				// found record is not better so we keep this and delete the old one
				objContacts.Remove( objContactFound );
			}

			ReportStatus( "Processed contact ", objNewContact.FirstName, " ", objNewContact.LastName
				, " (", objContacts.Count.ToString(), "/", objDT.Rows.Count.ToString(), ")" );

			return objNewContact;
		}

		public List<Note> RefreshProviderNotes( List<Contact> objContacts = null, bool bImportAllRecords = false, bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			if( objContacts == null )
			{
				// need Id, Name, FirstName, LastName, RecruitingID__c (default) and CreatedDate
				objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError, "CreatedDate" );
			}

			// refresh the notes/comments for each provider
			DataTable objDT = null;
			if( objContacts != null )
				objDT = objDB.GetDataTableFromSQLFile( "SQLProviderNotes.txt", Company2SFUtils.CreateProviderScript( objContacts ) );
			else
				if( bImportAllRecords )
					objDT = objDB.GetDataTableFromSQLFile( "SQLAllProvidersNotes.txt" );
				else
					objDT = objDB.GetDataTableFromSQLFile( "SQLProviderNotes.txt", Company2SFUtils.CreateProviderScript( objContacts ) );

			ReportStatus( "\r\n", objDT.Rows.Count.ToString(), " provider comment rows retrieved.\r\n" );

			DeleteExistingProviderNotes( objContacts );

			List<Note> objNotes = ValidateNotes( objDT, objContacts );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Id", objNotes.ToArray<sObject>() );

			ReportStatus( "Reporting errors for contact notes" );

			Company2SFUtils.ReportErrorsToHistoryFile( objResults, objNotes );

			//// create CSV file / set the Ids in the list of facilities
			//Company2SFUtils.SetIdsReportErrors( objNotes, objResults, tbStatus );
			//objResults = null;

			ReportStatus( objNotes.Count.ToString(), " Total provider notes loaded. \r\n" );

			objWatch.Stop();
			ReportStatus( "Finished loading notes. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours ", objWatch.Elapsed.Minutes.ToString(), " minutes." );

			return objNotes;
		}

		public List<Sub_Region__c> RefreshHierarchy( bool bDisplayOnly = true )
		{
			List<Division__c> objDivisions = new List<Division__c>();

			// read divisions from CSV file
			string strFileName = string.Concat( strAppPath, "CSV_Divisions.csv" );
			objDivisions.ReadFile<Division__c>( strFileName, "Name,Division_Description__c" );

			UpsertResult[] objDivisionResults = null;
			if( !bDisplayOnly )
				objDivisionResults = objAPI.Upsert( "Name", objDivisions.ToArray<sObject>() );
			//ReportStatus(objDivisionResults);

			// create CSV file / set the Ids in the list of facilities
			Company2SFUtils.SetIdsReportErrors( objDivisions, objDivisionResults, tbStatus );

			List<Region__c> objRegions = new List<Region__c>();

			// read regions from CSV file
			strFileName = string.Concat( strAppPath, "CSV_Regions.csv" );
			objRegions.ReadFile<Region__c>( strFileName, "Name,Region_Description__c,Division_Code__c" );

			// set the relationship between Regions and Divisions
			foreach( Region__c objReg in objRegions )
			{
				// convert division code into lookup id
				string strDivisionCode = objReg.Division_Code__c;
				string strDivisionId = "";
				Division__c objFoundDivision = objDivisions.FirstOrDefault( i => i.Name == strDivisionCode );
				if( objFoundDivision != null )
					strDivisionId = objFoundDivision.Id;
				objReg.Division_Code__c = strDivisionId;
			}

			UpsertResult[] objRegionResults = null;
			if( !bDisplayOnly )
				objRegionResults = objAPI.Upsert( "Name", objRegions.ToArray<sObject>() );
			//ReportStatus(objRegionResults);

			// create CSV file / set the Ids in the list of facilities
			Company2SFUtils.SetIdsReportErrors( objRegions, objRegionResults, tbStatus );

			List<Sub_Region__c> objSub_Regions = new List<Sub_Region__c>();

			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			// read regions from CSV file
			strFileName = string.Concat( strAppPath, "CSV_SubRegions.csv" );
			objSub_Regions.ReadFile<Sub_Region__c>( strFileName, "Name,Sub_Region_Description__c,Region_Code__c,OwnerId" );

			foreach( Sub_Region__c objSubReg in objSub_Regions )
			{
				// convert Owner name into User id
				string strOwner = objSubReg.OwnerId;
				User objFound = objExistingUsers.FirstOrDefault( i => i.Name.Equals( strOwner ) );
				if( objFound != null )
					objSubReg.Recruiting_Manager__c = objFound.Id;
				objSubReg.OwnerId = "";

				// convert region code into lookup id
				string strRegionCode = objSubReg.Region_Code__c;
				string strRegionId = "";
				Region__c objFoundRegion = objRegions.FirstOrDefault( i => i.Name == strRegionCode );
				if( objFoundRegion != null )
					strRegionId = objFoundRegion.Id;
				objSubReg.Region_Code__c = strRegionId;

				objSubReg.SubRegionCode__c = objSubReg.Name;
			}

			UpsertResult[] objSub_RegionResults = null;
			if( !bDisplayOnly )
				objSub_RegionResults = objAPI.Upsert( "Name", objSub_Regions.ToArray<sObject>() );
			//ReportStatus(objSub_RegionResults);

			// create CSV file / set the Ids in the list of facilities
			Company2SFUtils.SetIdsReportErrors( objSub_Regions, objSub_RegionResults, tbStatus );

			return objSub_Regions;
		}

		public List<Facility__c> RefreshFacilities( bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			ReportStatus( DateTime.Now.ToString(), " Starting refresh of facilities.\r\n" );

			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLFacility.txt" );

			ReportStatus( objDT.Rows.Count.ToString(), " facility rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				ReportStatus( objDB.ErrorMessage );

				return null;
			}

			// get a list of recruiters, credentialers, schedulers per hospital
			DataTable objUsersPerHospitalDT = objDB.GetDataTableFromSQLFile( "SQLUsersPerHospital.txt" );

			ReportStatus( objUsersPerHospitalDT.Rows.Count.ToString(), " users per facility rows retrieved.\r\n" );

			List<Sub_Region__c> objSubRegions = RefreshHierarchy();
			//objAPI.Query<Sub_Region__c>("select Id, SubRegionCode__c from Sub_Region__c");objUsersPerHospitalDT

			ReportStatus( objSubRegions.Count.ToString(), " sub regions loaded.\r\n" );

			// get record types and the ids for each service line
			List<RecordType> objRecTypes = objAPI.Query<RecordType>( "select Id, Name from RecordType" );
			string strInpatientServicesType = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Inpatient Services" ).Id;
			string strEmergencyDepartmentType = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Emergency Department" ).Id;
			string strAnesthesiologyType = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Anesthesiology" ).Id;
			string strRadiologyType = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Radiology" ).Id;

			// get users list
			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			ReportStatus( objExistingUsers.Count.ToString(), " SF users retrieved.\r\n" );

			// load providers from datatable to contact list
			List<Facility__c> objFacilities = new List<Facility__c>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				// copy all datatable columns to contact object attributes
				Facility__c objNewFacility = objDR.ConvertTo<Facility__c>( true );

				//// detect and handle potential duplicates (check city, street, sub region, shortname, ...)
				//if (Convert.ToInt32(objDR["DupeFlag"]) == 1)
				//{
				//    string strShortName = objNewFacility.Short_Name__c;
				//    foreach (DataRow objMatchDR in objDT.Rows)
				//    {
				//        if( objMatchDR["Short_Name__c"].ToString())
				//    }
				//}
				string strServiceLine = objNewFacility.Service_line__c;
				switch( strServiceLine )
				{
					case "Inpatient": objNewFacility.RecordTypeId = strInpatientServicesType; break;
					case "Radiology": objNewFacility.RecordTypeId = strRadiologyType; break;
					case "Anesthesia": objNewFacility.RecordTypeId = strAnesthesiologyType; break;
					default: objNewFacility.RecordTypeId = strEmergencyDepartmentType; break;	// Emergency Department
				}


				// check if the hospital has hospitalists in it that need to be separated from non-hospitalists
				bool bCreateHospitalist = false;
				if( Convert.ToInt16( objDR[ "HasHospitalists" ] ) > 0 && Convert.ToInt16( objDR[ "HasNonHospitalists" ] ) > 0 )
				{
					// see if there is another hospital row for hospitalist with the same address
					DataRow[] objFound = objDT.Select( string.Concat(
						"Address_1__c = '", objNewFacility.Address_1__c, "' AND Sub_Region__c = 'IPS' AND HasNonHospitalists = 0" ) );
					if( objFound.Count() == 0 )
					{
						// try again using name and sub region
						objFound = objDT.Select( string.Concat(
						"Name LIKE '", objNewFacility.Name.Replace( "'", "''" ), "%' AND Sub_Region__c = 'IPS' AND HasNonHospitalists = 0 " ) );
					}
					if( objFound.Count() == 0 )
					{
						// try again using address and short name
						objFound = objDT.Select( string.Concat(
						"Address_1__c = '", objNewFacility.Address_1__c, "' AND Short_Name__c LIKE '%Hospitalist%' AND HasNonHospitalists = 0 " ) );
					}
					// if the hospitalist equivalent doesn't exist, 
					// create an extra hospital with the same data but different code, region and service line
					if( objFound.Count() == 0 )
						bCreateHospitalist = true;
				}

				// fetch and populate hospital's recruiter, credentialer and scheduler
				DataRow[] objUserHospital = objUsersPerHospitalDT.Select( string.Concat(
								"HospCode = '", objNewFacility.Site_Code__c, "'" ) );
				if( objUserHospital.Count() > 0 )
				{
					// populate user roles in hospital record
					objNewFacility.Recruiter__c = objUserHospital[ 0 ][ "Recruiter" ].ToString();
					objNewFacility.Recruiter_Backup__c = objUserHospital[ 0 ][ "RecruiterBackup" ].ToString();
					objNewFacility.Credentialer__c = objUserHospital[ 0 ][ "Credentialer" ].ToString();
					objNewFacility.Credentialer_Backup__c = objUserHospital[ 0 ][ "CredentialerBackup" ].ToString();
					objNewFacility.Scheduler__c = objUserHospital[ 0 ][ "Scheduler" ].ToString();
					objNewFacility.Scheduler_MLP__c = objUserHospital[ 0 ][ "SchedulerMLP" ].ToString();

					// populate email addresses
					objNewFacility.Recruiter_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "RecruiterEmail" ].ToString() );
					objNewFacility.Recruiter_Backup_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "RecruiterBackupEmail" ].ToString() );
					objNewFacility.Credentialer_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "CredentialerEmail" ].ToString() );
					objNewFacility.Credentialer_Backup_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "CredentialerBackupEmail" ].ToString() );
					objNewFacility.Scheduler_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "SchedulerEmail" ].ToString() );
					objNewFacility.Scheduler_MLP_Email__c = Company2SFUtils.FilterValidEmail( objUserHospital[ 0 ][ "SchedulerMLPEmail" ].ToString() );

					// convert user names into User ids
					objNewFacility.Recruiter__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Recruiter__c );
					objNewFacility.Recruiter_Backup__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Recruiter_Backup__c );
					objNewFacility.Credentialer__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Credentialer__c );
					objNewFacility.Credentialer_Backup__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Credentialer_Backup__c );
					objNewFacility.Scheduler__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Scheduler__c );
					objNewFacility.Scheduler_MLP__c = Company2SFUtils.ConvertToUserId( tbStatus, objExistingUsers, objNewFacility.Scheduler_MLP__c );

					// default MLP scheduler to the physician scheduler if omitted
					if( objNewFacility.Scheduler_MLP__c.IsNullOrBlank() )
						objNewFacility.Scheduler_MLP__c = objNewFacility.Scheduler__c;
					if( objNewFacility.Scheduler_MLP_Email__c.IsNullOrBlank() )
						objNewFacility.Scheduler_MLP_Email__c = objNewFacility.Scheduler_Email__c;

				}

				// convert sub regions code into lookup id
				string strSubRegionId = "";
				string strSubRegionCode = objNewFacility.Sub_Region__c;
				Sub_Region__c objFoundSubRegion = null;
				if( objNewFacility.Name.ToLower().Contains( "hospitalist" ) )
				{
					// try assigning hospital to the equivalent hospitalist sub region if the name has the word hospitalist
					strSubRegionCode = strSubRegionCode + "IPS";

					objFoundSubRegion = objSubRegions.FirstOrDefault( i => i.SubRegionCode__c == strSubRegionCode );
					if( objFoundSubRegion == null )
						// attempt to assign to IPS if default
						objFoundSubRegion = objSubRegions.FirstOrDefault( i => i.SubRegionCode__c == "IPS" );
				}
				else
					objFoundSubRegion = objSubRegions.FirstOrDefault( i => i.SubRegionCode__c == strSubRegionCode );

				if( objFoundSubRegion != null )
					strSubRegionId = objFoundSubRegion.Id;
				objNewFacility.Sub_Region__c = strSubRegionId;

				// add metaphone values
				objNewFacility.Metaphone_Name__c = objNewFacility.Name.ToNormalizedMetaphone();
				if( objNewFacility.Address_1__c.HasNumbers() )
					objNewFacility.Metaphone_Address__c = objNewFacility.Address_1__c.ToNormalizedMetaphone().Left( 50 );
				else
					objNewFacility.Metaphone_Address__c = objNewFacility.Address_2__c.ToNormalizedMetaphone().Left( 50 );
				objNewFacility.Metaphone_City__c = objNewFacility.City__c.ToNormalizedMetaphone();

				// fix time zone bug
				objNewFacility.Contract_Start_Date__c = Company2SFUtils.FixTimeZoneBug( objNewFacility.Contract_Start_Date__c );
				objNewFacility.Contract_End_Date__c = Company2SFUtils.FixTimeZoneBug( objNewFacility.Contract_End_Date__c );

				// set flags for contract dates
				objNewFacility.Contract_Start_Date__cSpecified = ( objNewFacility.Contract_Start_Date__c != null );
				objNewFacility.Contract_End_Date__cSpecified = ( objNewFacility.Contract_End_Date__c != null );

				objFacilities.Add( objNewFacility );

				if( bCreateHospitalist )
				{
					// create an extra hospitalist facility
					// copy all datatable columns to contact object attributes
					objNewFacility = objDR.ConvertTo<Facility__c>( true );
					objNewFacility.Site_Code__c = string.Concat( objNewFacility.Site_Code__c, "H" );
					objNewFacility.Sub_Region__c = "IPS";
					objNewFacility.Service_line__c = "InPatient Services"; 
					objNewFacility.RecordTypeId = strInpatientServicesType;
					objNewFacility.Name = string.Concat( objNewFacility.Name, "-Hospitalist" );
					objNewFacility.Short_Name__c = string.Concat( objNewFacility.Short_Name__c, "-Hospitalist" );
					objNewFacility.Short_Name__c = objNewFacility.Short_Name__c.Left( 20 );

					// convert sub regions code into lookup id
					objFoundSubRegion = objSubRegions.FirstOrDefault( i => i.SubRegionCode__c.Equals( "IPS" ) );
					if( objFoundSubRegion != null )
						strSubRegionId = objFoundSubRegion.Id;
					objNewFacility.Sub_Region__c = strSubRegionId;

					objFacilities.Add( objNewFacility );
				}

				ReportStatus( "Processed facility ", objNewFacility.Name, " - ", objNewFacility.Site_Code__c
					, " (", objFacilities.Count.ToString(), "/", objDT.Rows.Count.ToString(), ")" );
			}

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Site_Code__c", objFacilities.ToArray<sObject>() );

			Company2SFUtils.ReportErrorsToHistoryFile( objResults, objFacilities );

			//// create CSV file / set the Ids in the list of facilities
			//Company2SFUtils.SetIdsReportErrors( objFacilities, objResults, tbStatus );

			//ShowData(objFacility);

			ReportStatus( objFacilities.Count.ToString(), " Total facilities loaded. \r\n" );

			objWatch.Stop();
			ReportStatus( DateTime.Now.ToString(), " - Finished loading Facilities. Duration: "
							, objWatch.Elapsed.Hours.ToString(), " hours "
							, objWatch.Elapsed.Minutes.ToString(), " minutes." );

			return objFacilities;
		}

		public List<Institution__c> RefreshInstitutions( List<Contact> objProviders = null, bool bDisplayOnly = true )
		{
			DataTable objDT = null;
			if( objProviders != null )
				objDT = objDB.GetDataTableFromSQLFile( "SQLInstitution.txt", Company2SFUtils.CreateProviderScript( objProviders ) );
			else
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllInstitutions.txt" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDT.Rows.Count, " institution rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// add copy of main column to the table to help with duplicate detection
			Company2SFUtils.FlagDupesByAddressCityName( objDT );

			// load agencies from datatable to agency list
			List<Institution__c> objInstitutions = new List<Institution__c>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				//// detect duplicates and flag them in the DuplicateOfCode column
				//FlagDuplicates( objDR );

				// skip the duplicate institutions
				if( !objDR[ "DuplicateOfCode" ].IsNullOrBlank() )
					continue;

				// copy all datatable columns to credential agency object attributes
				Institution__c objNewInstitution = objDR.ConvertTo<Institution__c>();

				// add metaphone values
				string strName = objDR[ "OriginalName" ].ToString();
				objNewInstitution.Metaphone_Name__c = strName.ToNormalizedMetaphone();
				if( objNewInstitution.Address1__c.HasNumbers() )
					objNewInstitution.Metaphone_Address__c = objNewInstitution.Address1__c.ToNormalizedMetaphone().Left( 50 );
				else
					objNewInstitution.Metaphone_Address__c = objNewInstitution.Address2__c.ToNormalizedMetaphone().Left( 50 );
				objNewInstitution.Metaphone_City__c = objNewInstitution.City__c.ToNormalizedMetaphone();

				// set the name after duplication removal
				objNewInstitution.Name = objDR[ "ModifiedName" ].IsNullOrBlank() ?
										objDR[ "OriginalName" ].ToString() : objDR[ "ModifiedName" ].ToString();

				objInstitutions.Add( objNewInstitution );
			}

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Company_Agency_Match__c", objInstitutions.ToArray<sObject>() );
			//ReportStatus(objResults);

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objInstitutions, objResults, tbStatus );

			//ShowData(objInstitutions);

			return objInstitutions;
		}

		public List<Credential_Agency__c> RefreshAgencies( List<Contact> objProviders = null, bool bDisplayOnly = true )
		{
			DataTable objDT = null;
			if( objProviders != null )
				objDT = objDB.GetDataTableFromSQLFile( "SQLAgency.txt", Company2SFUtils.CreateProviderScript( objProviders ) );
			else
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllAgencies.txt" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDT.Rows.Count, " agency rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// add copy of main column to the table to help with duplicate detection
			//CreateMatchColumn( objDT, "OriginalName", true );

			Company2SFUtils.FlagDupesByAddressCityName( objDT );

			// add agencies not in Company
			List<Credential_Agency__c> objAgencies = new List<Credential_Agency__c>( objDT.Rows.Count );
			string strFileName = string.Concat( strAppPath, "CSV_AgenciesNotInCompanyApp.csv" );
			objAgencies.ReadFile( strFileName, "Code__c,Name,State_Licensing_Agency__c,Address1__c,Address2__c,City__c,State__c,"
+ "Zip__c,Phone__c,Ext__c,Fax__c,Contact__c,Title__c,Salutation__c,Credential_Type__c,Company_Agency_Match__c"
							, true );

			// load agencies from datatable to agency list
			foreach( DataRow objDR in objDT.Rows )
			{
				//// detect duplicates and flag them in the DuplicateOfCode column
				//FlagDuplicates( objDR );

				// skip the duplicate institutions
				if( !objDR[ "DuplicateOfCode" ].IsNullOrBlank() )
					continue;

				// copy all datatable columns to credential agency object attributes
				Credential_Agency__c objNewAgency = objDR.ConvertTo<Credential_Agency__c>();

				// add metaphone values
				string strName = objDR[ "OriginalName" ].ToString();
				objNewAgency.Metaphone_Name__c = strName.ToNormalizedMetaphone();
				if( objNewAgency.Address1__c.HasNumbers() )
					objNewAgency.Metaphone_Address__c = objNewAgency.Address1__c.ToNormalizedMetaphone().Left( 50 );
				else
					objNewAgency.Metaphone_Address__c = objNewAgency.Address2__c.ToNormalizedMetaphone().Left( 50 );
				objNewAgency.Metaphone_City__c = objNewAgency.City__c.ToNormalizedMetaphone();

				// set the name after duplication removal
				objNewAgency.Name = objDR[ "ModifiedName" ].IsNullOrBlank() ?
										objDR[ "OriginalName" ].ToString() : objDR[ "ModifiedName" ].ToString();

				objAgencies.Add( objNewAgency );
			}

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Company_Agency_Match__c", objAgencies.ToArray<sObject>() );
			//ReportStatus(objResults);

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objAgencies, objResults, tbStatus );

			//ShowData(objAgencies);

			return objAgencies;
		}

		public List<Credential_Agency__c> ReportDuplicateAgencies()
		{
			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllAgencies.txt" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDT.Rows.Count, " agency rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// add copy of main column to the table to help with duplicate detection
			Company2SFUtils.FlagDupesByAddressCityName( objDT );
			//CreateMatchColumn( objDT, "OriginalName,Address1__c", true, true );

			// create header row
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\nDUPLICATE ADDRESSES:\r\n" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCode:\tDuplicate Name:\tAddress:\tMatch Name:\t"
					, "Duplicate of Code:\tDuplicate of Name:\tRelated Address:" );

			// load agencies from datatable to agency list
			List<Credential_Agency__c> objNonDuplicateAgencies = new List<Credential_Agency__c>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				// skip the duplicate institutions
				if( !objDR[ "DuplicateOfCode" ].IsNullOrBlank() )
				{
					// report duplicates
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\n"
							, objDR[ "Company_Agency_Match__c" ]
							, "\t", objDR[ "OriginalName" ]
							, "\t", objDR[ "Address1__c" ]
							, "\t", objDR[ "City__c" ]
							, "\t", objDR[ "MatchOriginalName" ]
							, "\t", objDR[ "DuplicateOfCode" ]
							, "\t", objDR[ "DuplicateOfName" ]
							, "\t", ( (DataRow) objDR[ "DuplicateOfRowNumber" ] )[ "Address1__c" ] );
					//tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDR[ "OriginalName" ]
					//        , "\t", objDR[ "ModifiedName" ]
					//        , "\t", objDR[ "MatchOriginalName" ]
					//        , "\t", objDR[ "Company_Agency_Match__c" ]
					//        , "\t", objDR[ "DuplicateOfCode" ]
					//        , "\t", objDR[ "DuplicateOfName" ] );
					continue;
				}


				// copy all datatable columns to credential agency object attributes
				Credential_Agency__c objNewAgency = objDR.ConvertTo<Credential_Agency__c>();

				// set the name after duplication removal
				objNewAgency.Name = objDR[ "ModifiedName" ].IsNullOrBlank() ?
										objDR[ "OriginalName" ].ToString() : objDR[ "ModifiedName" ].ToString();

				objNonDuplicateAgencies.Add( objNewAgency );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\nNON-DUPLICATE AGENCIES:\r\n" );
			Company2SFUtils.SetIdsReportErrors( objNonDuplicateAgencies, null, tbStatus );

			//// match by address now
			//tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\nDUPLICATE ADDRESSES:\r\n" );


			//// create header row
			//tbStatus.Text = string.Concat( tbStatus.Text, "\r\nDuplicate Name:\tAddress:\tMatch Name:\tCode:\t"
			//        , "Duplicate of Code:\tDuplicate of Name:\tRelated Address:" );

			//// add copy of main column to the table to help with duplicate detection
			//CreateMatchColumn( objDT, "Address1__c", true, true );

			//// check agencies with duplicate addresses
			//foreach( DataRow objDR in objDT.Rows )
			//    // skip the duplicate institutions
			//    if( !objDR[ "DuplicateOfCode" ].IsNullOrBlank() )
			//        // report duplicates
			//        tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDR[ "OriginalName" ]
			//                , "\t", objDR[ "Address1__c" ]
			//                , "\t", objDR[ "MatchOriginalName" ]
			//                , "\t", objDR[ "Company_Agency_Match__c" ]
			//                , "\t", objDR[ "DuplicateOfCode" ]
			//                , "\t", objDR[ "DuplicateOfName" ]
			//                , "\t", ((DataRow) objDR[ "DuplicateOfRowNumber" ] )[ "Address1__c" ] );

			return objNonDuplicateAgencies;
		}

		public List<Institution__c> ReportDuplicateInstitutions()
		{
			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllInstitutions.txt" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDT.Rows.Count, " institution rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// add copy of main column to the table to help with duplicate detection
			Company2SFUtils.FlagDupesByAddressCityName( objDT );
			//CreateMatchColumn( objDT, "OriginalName,Address1__c", true, true );

			// create header row
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\nDUPLICATE ADDRESSES:\r\n" );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCode:\tDuplicate Name:\tAddress:\tMatch Name:\t"
					, "Duplicate of Code:\tDuplicate of Name:\tRelated Address:" );

			// load agencies from datatable to institution list
			List<Institution__c> objNonDuplicateInstitutions = new List<Institution__c>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				// skip the duplicate institutions
				if( !objDR[ "DuplicateOfCode" ].IsNullOrBlank() )
				{
					// report duplicates
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\n"
							, objDR[ "Company_Agency_Match__c" ]
							, "\t", objDR[ "OriginalName" ]
							, "\t", objDR[ "Address1__c" ]
							, "\t", objDR[ "City__c" ]
							, "\t", objDR[ "MatchOriginalName" ]
							, "\t", objDR[ "DuplicateOfCode" ]
							, "\t", objDR[ "DuplicateOfName" ]
							, "\t", ( (DataRow) objDR[ "DuplicateOfRowNumber" ] )[ "Address1__c" ] );
					continue;
				}


				// copy all datatable columns to credential institution object attributes
				Institution__c objNewInstitution = objDR.ConvertTo<Institution__c>();

				// set the name after duplication removal
				objNewInstitution.Name = objDR[ "ModifiedName" ].IsNullOrBlank() ?
										objDR[ "OriginalName" ].ToString() : objDR[ "ModifiedName" ].ToString();

				objNonDuplicateInstitutions.Add( objNewInstitution );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\nNON-DUPLICATE INSTITUTIONS:\r\n" );
			Company2SFUtils.SetIdsReportErrors( objNonDuplicateInstitutions, null, tbStatus );

			return objNonDuplicateInstitutions;
		}

		public List<Credential_Subtype__c> RefreshSubtypes( bool bDisplayOnly = true )
		{
			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLSubtype.txt" );
			tbStatus.Text = string.Concat( objDT.Rows.Count, " subtype rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// create field to hold Metaphone code
			DataColumn objCol = new DataColumn( "Metaphone", typeof( string ) );
			objDT.Columns.Add( objCol );
			objCol = new DataColumn( "PhoneticValue", typeof( int ) );
			objDT.Columns.Add( objCol );

			// load agencies from datatable to agency list
			List<Credential_Subtype__c> objSubTypes = new List<Credential_Subtype__c>( objDT.Rows.Count );
			foreach( DataRow objDR in objDT.Rows )
			{
				// copy all datatable columns to credential agency object attributes
				Credential_Subtype__c objNewSubType = objDR.ConvertTo<Credential_Subtype__c>();

				//// calculate metaphone to help detect duplicates
				//string strName = objDR[ "Name" ].ToString();
				//NameSimilarity objNS = new NameSimilarity();
				//objNS.Name = strName;
				//objDR[ "Metaphone" ] = objNS.MetaphoneKey;
				//objDR[ "PhoneticValue" ] = objNS.Value;

				objSubTypes.Add( objNewSubType );
			}

			//foreach( DataRow objDR in objDT.Rows )
			//    tbStatus.Text = string.Concat(
			//        tbStatus.Text, "\r\n", objDR[ "Metaphone" ].ToString(), ",", objDR[ "Name" ].ToString()
			//        , ",", objDR[ "Credential_Type__c" ].ToString()
			//        , ",", objDR[ "PhoneticValue" ].ToString() );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Company_SubType_Match__c", objSubTypes.ToArray<sObject>() );
			//ReportStatus(objResults);

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objSubTypes, objResults, tbStatus );

			//ShowData(objSubTypes);

			return objSubTypes;
		}

		public List<Credential__c> RefreshCredentials( List<Contact> objProviders = null
					, List<Credential_Agency__c> objAgencies = null, List<Credential_Subtype__c> objSubTypes = null, bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;
			if( objProviders != null )
			{
				if( bMassDataLoad )
					objDT = objDB.GetDataTableFromSQLFile( "SQLAllCredentials.txt" );
				else
					objDT = objDB.GetDataTableFromSQLFile( "SQLCredential.txt", Company2SFUtils.CreateProviderScript( objProviders ) );
			}
			else
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllCredentials.txt" );
			tbStatus.Text = string.Concat( objDT.Rows.Count, " credential rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// get list of Contacts
			List<Contact> objContacts = null;
			if( objProviders != null )
				objContacts = objProviders;
			else
				objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError, "", " PhysicianNumber__c > 0 " );

			//// get list of subtypes if needed
			//if (objSubTypes == null)
			//    objSubTypes = objAPI.Query<Credential_Subtype__c>(
			//        "select ID, Company_SubType_Match__c from Credential_Subtype__c ORDER BY Company_SubType_Match__c");

			// get list of agencies if needed
			if( objAgencies == null )
				objAgencies = objAPI.Query<Credential_Agency__c>(
					"select ID, Company_Agency_Match__c from Credential_Agency__c ORDER BY Company_Agency_Match__c " );

			// get record types and the ids for Education and Experience
			List<RecordType> objRecTypes = objAPI.Query<RecordType>( "select Id, Name from RecordType" );

			// get list of agencies from Company free from duplicates
			DataTable objDTAgencies = objDB.GetDataTableFromSQLFile( "SQLAllAgencies.txt" );

			Company2SFUtils.FlagDupesByAddressCityName( objDTAgencies );

			// create a sorted list to help avoid duplicate names
			SortedSet<string> objNames = new SortedSet<string>();

			// load credentials from datatable to credential list
			List<Credential__c> objCredentials = new List<Credential__c>();
			int iSkipped = 0;
			foreach( DataRow objDR in objDT.Rows )
			{
				// copy all datatable columns to credential object attributes
				Credential__c objNewCredential = objDR.ConvertTo<Credential__c>();
				//string strFirstName = objDR[ "FirstName" ].ToString();
				//string strLastName = objDR[ "LastName" ].ToString();					 

				// convert agency code into lookup id
				string strAgencyCode = objNewCredential.Credential_Agency__c;
				objNewCredential.Credential_Agency__c = Company2SFUtils.ConvertToAgencyId( objAgencies, strAgencyCode );

				// if agency code could not be converted into lookup id, this is a possible duplicate
				// so we attempt to find the related unique code
				if( objNewCredential.Credential_Agency__c.IsNullOrBlank() )
				{
					DataRow[] objFound = objDTAgencies.Select(
								string.Concat( "Company_Agency_Match__c = '", strAgencyCode, "'" ) );
					if( objFound.Count() > 0 )
						// attempt again to convert agency code into lookup id
						objNewCredential.Credential_Agency__c = Company2SFUtils.ConvertToAgencyId( objAgencies
											, objFound[ 0 ][ "DuplicateOfCode" ].ToString() );
				}

				// convert subtype code into lookup id
				string strSubTypeCode = objNewCredential.Credential_Sub_Type__c;
				//string strSubTypeId = "";
				//Credential_Subtype__c objFoundSubType = objSubTypes.FirstOrDefault( i => i.Company_SubType_Match__c == strSubTypeCode );
				//if (objFoundSubType != null)
				//    strSubTypeId = objFoundSubType.Id;
				//objNewCredential.Credential_Sub_Type__c = strSubTypeId;

				// set sub type without having to find sub type object
				Credential_Subtype__c objSubType = new Credential_Subtype__c();
				objSubType.Company_SubType_Match__c = strSubTypeCode;
				objNewCredential.Credential_Sub_Type__r = objSubType;
				objNewCredential.Credential_Sub_Type__c = "";

				// link credential to Contact/Provider using PhysicianNumber
				// or RecruitingID if State License
				if( objNewCredential.Physician_Number__c != "0"
						|| objNewCredential.Name.Contains( "State License " ) )
				{
					double dblPhysicianID = Convert.ToDouble( objNewCredential.Physician_Number__c );
					Contact objProvider = null;
					if( objNewCredential.Name.Contains( "State License " ) )
						objProvider = objContacts.FirstOrDefault( i => i.RecruitingID__c == dblPhysicianID );
					else
						objProvider = objContacts.FirstOrDefault( i => i.PhysicianNumber__c == dblPhysicianID );

					// skip if provider was not found
					if( objProvider == null )
					{
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find physician nbr "
							, objNewCredential.Physician_Number__c, " for ", objNewCredential.Name );
						iSkipped++;
						continue;
					}

					objNewCredential.Contact__c = objProvider.Id;

					//Contact objProvider = new Contact();
					//if( objNewCredential.Name.Contains( "State License " ) )
					//{
					//    objProvider.RecruitingID__c = dblPhysicianID;
					//    objProvider.RecruitingID__cSpecified = true;
					//}
					//else
					//{
					//    objProvider.PhysicianNumber__c = dblPhysicianID;
					//    objProvider.PhysicianNumber__cSpecified = true;
					//}
					//objNewCredential.Contact__r = objProvider;
					//objNewCredential.Physician_Number__c = "";

					// attempting to make the credential name unique
					string strNewName = objNewCredential.Name;
					if( objProvider.LastName != null && objProvider.FirstName != null )
						strNewName = string.Concat( objProvider.FirstName, " ", objProvider.LastName
										, "-", objNewCredential.Name );

					// if name is duplicate, add middle name
					//Credential__c objDupe = objCredentials.FirstOrDefault( c => c.Name.Equals( strNewName ) );
					if( objNames.Contains( strNewName ) ) //objDupe != null )
					{
						if( !objProvider.Middle_Name__c.IsNullOrBlank() )
						{
							strNewName = string.Concat( objProvider.FirstName, " "
											, objProvider.Middle_Name__c, " ", objProvider.LastName
											, "-", objNewCredential.Name );

							//objDupe = objCredentials.FirstOrDefault( c => c.Name.Equals( strNewName ) );
						}

						// if name is still duplicate, add region
						if( objNames.Contains( strNewName ) ) //objDupe != null )
						{
							if( !objProvider.Owning_Region__c.IsNullOrBlank() )
							{
								strNewName = string.Concat( objProvider.FirstName, " ", objProvider.LastName
												, " (", objProvider.Owning_Region__c
												, ")-", objNewCredential.Name );

								// if name is still duplicate, add year or " #2"
								//objDupe = objCredentials.FirstOrDefault( c => c.Name.Equals( strNewName ) );
								if( objNames.Contains( strNewName ) ) //objDupe != null )
									if( objNewCredential.From__c != null )
										strNewName = string.Concat( strNewName, " - "
														, ( (DateTime) objNewCredential.From__c ).Year.ToString() );
									else
										strNewName = string.Concat( strNewName, " #2" );
							}
						}
					}

					objNewCredential.Name = strNewName;

					// add name to the list to be checked in the next iteration
					objNames.Add( strNewName );
				}
				else // skip if physician number is zero
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nPhysician nbr is zero for ", objNewCredential.Name );
					iSkipped++;
					continue;
				}

				// fix time zone bug
				objNewCredential.From__c = Company2SFUtils.FixTimeZoneBug( objNewCredential.From__c );
				objNewCredential.To__c = Company2SFUtils.FixTimeZoneBug( objNewCredential.To__c );

				// avoid sending blank from/to dates
				objNewCredential.From__cSpecified = ( objNewCredential.From__c != null );
				objNewCredential.To__cSpecified = ( objNewCredential.To__c != null );

				// set the record type id according to the credential type
				string strType = objNewCredential.Credential_Type__c;
				string strRecordTypeId = objRecTypes.FirstOrDefault<RecordType>( i => i.Name.Equals( strType ) ).Id;
				objNewCredential.RecordTypeId = strRecordTypeId;

				objCredentials.Add( objNewCredential );

				ReportStatus( string.Concat( "Processed credential ", objNewCredential.Name, " ", objNewCredential.Credential_Type__c
					, " (", objCredentials.Count, "/", objDT.Rows.Count, ")" ) );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " rows skipped (Institution mismatch)." );

			ReportStatus( "Reporting errors for credentials" );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Name", objCredentials.ToArray<sObject>() );

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objCredentials, objResults, tbStatus );

			objWatch.Stop();
			ReportStatus( string.Concat( "Finished loading credentials. Duration: "
							, objWatch.Elapsed.Hours, " hours ", objWatch.Elapsed.Minutes, " minutes." ) );

			return objCredentials;
		}

		public List<Education_or_Experience__c> RefreshEducationExperience( List<Contact> objContacts = null
					, List<Institution__c> objInstitutions = null, bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null;
			string strProviderScript = "";
			if( objContacts != null )
			{
				if( bMassDataLoad )
					objDT = objDB.GetDataTableFromSQLFile( "SQLAllEducationsExperiences.txt" );
				else
				{
					strProviderScript = Company2SFUtils.CreateProviderScript( objContacts );
					objDT = objDB.GetDataTableFromSQLFile( "SQLEducationExperience.txt", strProviderScript );
				}
			}
			else
			{
				objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError ); //, "", " PhysicianNumber__c > 0 " );
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllEducationsExperiences.txt" );
			}

			tbStatus.Text = string.Concat( objDT.Rows.Count, " education/experience rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			// get list of agencies from SalesForce and from Company and get rid of duplicates from Company
			DataTable objDTInstitutions;
			if( objInstitutions == null )
			{
				objDTInstitutions = objDB.GetDataTableFromSQLFile( "SQLAllInstitutions.txt" );
				objInstitutions = objAPI.Query<Institution__c>( "select Id, Name, Company_Agency_Match__c from Institution__c order by Company_Agency_Match__c" );
			}
			else
			{
				if( bMassDataLoad )
					objDTInstitutions = objDB.GetDataTableFromSQLFile( "SQLAllInstitutions.txt" );
				else
				{
					if( strProviderScript.Equals( "" ) )
						strProviderScript = Company2SFUtils.CreateProviderScript( objContacts );
					objDTInstitutions = objDB.GetDataTableFromSQLFile( "SQLInstitution.txt", strProviderScript );
				}
			}

			// add copy of main column to the table to help with duplicate detection
			Company2SFUtils.FlagDupesByAddressCityName( objDTInstitutions );

			// create a sorted list to help avoid duplicate names
			SortedSet<string> objNames = new SortedSet<string>();

			// get record types and the ids for Education and Experience
			List<RecordType> objRecTypes = objAPI.Query<RecordType>( "select Id, Name from RecordType" );
			string strEducationTypeId = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Degree/ Education" ).Id;
			// changed from "Graduate Education").Id;
			string strExperienceTypeId = objRecTypes.FirstOrDefault<RecordType>( i => i.Name == "Experience" ).Id;

			// load credentials from datatable to credential list
			List<Education_or_Experience__c> objEducExper = new List<Education_or_Experience__c>();
			List<Resident__c> objResidents = new List<Resident__c>();
			int iSkipped = 0;
			string strLastRecordName = "";
			foreach( DataRow objDR in objDT.Rows )
			{
				// copy all datatable columns to credential object attributes
				Education_or_Experience__c objNewEducExp = objDR.ConvertTo<Education_or_Experience__c>();

				// convert institution code into lookup id
				Institution__c objInstitutionFound = objInstitutions.FirstOrDefault(
									i => i.Company_Agency_Match__c.NotNullAndEquals( objNewEducExp.Institution__c ) );
				if( objInstitutionFound != null )
					objNewEducExp.Institution__c = objInstitutionFound.Id;
				else
				{
					// if institution code could not be converted into lookup id, this is a possible duplicate
					// so we attempt to find the related unique code
					DataRow[] objFound = objDTInstitutions.Select( string.Concat(
									"Company_Agency_Match__c = '", objNewEducExp.Institution__c, "'" ) );
					if( objFound.Count() > 0 )
					{
						// attempt again to convert institution code into lookup id
						string strInstitution = objFound[ 0 ][ "DuplicateOfCode" ].ToString();
						objInstitutionFound = objInstitutions.FirstOrDefault( i => i.Company_Agency_Match__c.Equals( strInstitution ) );
						if( objInstitutionFound != null )
							objNewEducExp.Institution__c = objInstitutionFound.Id;
					}

					if( objNewEducExp.Institution__c.IsNullOrBlank() )
					{
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find institution "
							, objNewEducExp.Institution__c );
						iSkipped++;
						continue;
					}
				}

				// link credential to Contact/Provider using PhysicianNumber or RecruitingID
				double dblPhysicianNumber = Convert.ToDouble( objNewEducExp.Contact__c );
				string strContactId = "";
				Contact objProvider = objContacts.FirstOrDefault( i =>
					i.PhysicianNumber__c == dblPhysicianNumber || i.RecruitingID__c == dblPhysicianNumber );
				if( objProvider != null )
					strContactId = objProvider.Id;
				else // skip if provider was not found
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find physician nbr "
						, objNewEducExp.Contact__c, " for ", objNewEducExp.Name );
					iSkipped++;
					continue;
				}
				objNewEducExp.Contact__c = strContactId;

				// fix time zone bug
				objNewEducExp.From__c = Company2SFUtils.FixTimeZoneBug( objNewEducExp.From__c );
				objNewEducExp.To__c = Company2SFUtils.FixTimeZoneBug( objNewEducExp.To__c );

				// attempting to make the credential name unique
				string strNewName = objNewEducExp.Name;
				if( objProvider.LastName != null && objProvider.FirstName != null )
					strNewName = string.Concat( objProvider.FirstName, " ", objProvider.LastName
									, "-", objNewEducExp.Name );

				// if name is duplicate, add middle name
				if( objNames.Contains( strNewName ) )
				{
					if( !objProvider.Middle_Name__c.IsNullOrBlank() )
					{
						strNewName = string.Concat( objProvider.FirstName, " "
										, objProvider.Middle_Name__c, " ", objProvider.LastName
										, "-", objNewEducExp.Name );

						//objDupe = objEducExper.FirstOrDefault( c => c.Name.Equals( strNewName ) );
					}

					// if name is still duplicate, add region
					if( objNames.Contains( strNewName ) )
					{
						if( !objProvider.Owning_Region__c.IsNullOrBlank() )
						{
							strNewName = string.Concat( objProvider.FirstName, " ", objProvider.LastName
											, " (", objProvider.Owning_Region__c
											, ")-", objNewEducExp.Name );
						}
					}

					// if name is still duplicate, add month/year or " #2"
					if( objNames.Contains( strNewName ) )
					{
						string strSuffix = " #2";
						// append from date or to date, if both null, add #2
						if( objNewEducExp.From__c != null )
							strSuffix = string.Concat( " - ", ( (DateTime) objNewEducExp.From__c ).Month.ToString()
												, "/", ( (DateTime) objNewEducExp.From__c ).Year.ToString() );
						else
							if( objNewEducExp.To__c != null )
								strSuffix = string.Concat( " - ", ( (DateTime) objNewEducExp.To__c ).Month.ToString()
													, "/", ( (DateTime) objNewEducExp.To__c ).Year.ToString() );
						strNewName = string.Concat( strNewName, strSuffix );
					}
				}

				objNewEducExp.Name = strNewName.Left( 80 );

				// add name to list to be checked in the next iteration
				objNames.Add( strNewName );

				// if the record is about Residency, add it to the Residency list to insert them later
				// and skip the record in this loop (that is, do not insert it as an Experience)
				if( objNewEducExp.Type__c.Contains( "Residen" ) && objInstitutionFound != null )
				{
					Resident__c objNewResident	= new Resident__c();

					objNewResident.From__c = objNewEducExp.From__c;
					objNewResident.From__cSpecified = ( objNewEducExp.From__c != null );
					objNewResident.To__c = objNewEducExp.To__c;
					objNewResident.To__cSpecified = ( objNewEducExp.To__c != null );
					objNewResident.Contact__c = objNewEducExp.Contact__c;
					objNewResident.Name = strNewName.Left( 80 );	// limit to 80 characters

					switch( objNewEducExp.Type__c )
					{
						case "Residency": objNewResident.Type__c = "Residency";
							break;
						case "Internship & Residency": objNewResident.Type__c = "Internship and Residency";
							break;
						case "Chief Resident": objNewResident.Type__c = "Chief Resident";
							break;
					}

					// store institution Id in residency program to convert it to the Residency Program ID later
					objNewResident.Residency_Program__c = string.Concat( objInstitutionFound.Id, "|", objInstitutionFound.Name );

					objResidents.Add( objNewResident );

					// skip this Education Experience (record will be stored as Residency instead)
					continue;
				}

				//// if the record is Experience about Residency, 
				//// switch it to Graduate Education and set the proper picklist value
				//if(objNewEducExp.Description__c.Equals("Experience") 
				//    && objNewEducExp.Name.Contains( "Residen" )) // look for Resident or Residency
				//{
				//    objNewEducExp.Description__c = "Graduate Education";
				//    objNewEducExp.Type__c = "Residency";
				//}

				// set the appropriate record type
				if( objNewEducExp.Description__c.Equals( "Degree/ Education" ) )
					objNewEducExp.RecordTypeId = strEducationTypeId;
				else
					objNewEducExp.RecordTypeId = strExperienceTypeId;

				// check whether the comments give a hint of radiology training/fellowship
				if( objNewEducExp.Type__c.Equals( "Fellowship" ) )
				{
					if( objNewEducExp.Comments__c.Contains( "radiology" ) )
						objNewEducExp.Type__c = "Radiology Diagnostic Fellowship";

					if( objNewEducExp.Comments__c.Contains( "Musculoskeletal Radiology" ) )
					{
						objNewEducExp.Type__c = "Radiology Diagnostic Fellowship";
						objNewEducExp.Sub_Type__c = "Musculoskeletal Imaging";
					}

					if( objNewEducExp.Comments__c.Contains( "Neuroradiology" ) )
					{
						objNewEducExp.Type__c = "Radiology Diagnostic Fellowship";
						objNewEducExp.Sub_Type__c = "Neuroradiology";
					}
				}

				objNewEducExp.From__cSpecified = ( objNewEducExp.From__c != null );
				objNewEducExp.To__cSpecified = ( objNewEducExp.To__c != null );

				objEducExper.Add( objNewEducExp );

				// store name to detect and remediate duplicate names 
				// (for cases when the same provider had same experience at same hospital but at different dates)
				strLastRecordName = objNewEducExp.Name;

				ReportStatus( string.Concat( "Processed degree/experience ", objNewEducExp.Name, " ", objNewEducExp.Type__c
							, " (", objEducExper.Count, "/", objDT.Rows.Count, ")" ) );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " rows skipped (Institution or physician nbr mismatch)." );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Name", objEducExper.ToArray<sObject>() );

			ReportStatus( "Reporting errors for degree/experience" );

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objEducExper, objResults, tbStatus );

			// collect resid. programs to tie them to the residents
			List<Residency_Program__c> objResidPrograms = objAPI.Query<Residency_Program__c>(
				"select Id, Name, Program_Institution__c, Program_Speciality__c, Service_Line__c from Residency_Program__c" );

			// get facilities to lookup service line/division
			List<Facility__c> objFacilities = objAPI.Query<Facility__c>(
				"select Id, Name, Sub_Region__c, Service_line__c, Sub_Region__r.Region_Code__r.Division_Code__c from Facility__c order by Name " );
			//// replace this with a list with all facilities instead of hitting SF every time
			//objFac = objAPI.QuerySingle<Facility__c>( string.Concat(
			//         "select Id, Name, Sub_Region__c, Service_line__c, "
			//         , "Sub_Region__r.Region_Code__r.Division_Code__c from Facility__c where Name like '"
			//         , strInstitName, "%' limit 1" ) );

			// convert Institution id to Resident Program id in the new resident records
			List<Resident__c> objResidentsToUpsert = new List<Resident__c>();
			foreach( Resident__c objResident in objResidents )
			{
				// split id from name inside Residency_Program__c
				string strInstitutionId = objResident.Residency_Program__c;
				string[] strSplitIdName = strInstitutionId.Split( '|' );
				strInstitutionId = strSplitIdName[ 0 ];
				string strInstitName = strSplitIdName[ 1 ];
				// with the Institution id, find or create a Resident Program and assign its id
				objResident.Residency_Program__c = GetResidentProgram(
					objInstitutions, objContacts, objResidPrograms, objFacilities
					, strInstitutionId, strInstitName, objResident.Contact__c );

				// skip the residents for whom a program could not be created/retrieved
				if( objResident.Residency_Program__c.Equals( "" ) )
					continue;

				objResidentsToUpsert.Add( objResident );

				ReportStatus( string.Concat( "Processed degree/experience ", objResident.Name, " "
							, " (", objResidentsToUpsert.Count, "/", objResidents.Count, ")" ) );
			}

			UpsertResult[] objResidentResults = null;
			if( !bDisplayOnly )
				objResidentResults = objAPI.Upsert( "Name", objResidentsToUpsert.ToArray<sObject>() );

			ReportStatus( "Reporting errors for residents" );

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objResidentsToUpsert, objResidentResults, tbStatus );

			objWatch.Stop();
			ReportStatus( string.Concat( "Finished loading degree/experience. Duration: "
							, objWatch.Elapsed.Hours, " hours ", objWatch.Elapsed.Minutes, " minutes." ) );

			return objEducExper;
		}

		public string GetResidentProgram( List<Institution__c> objInstitutions, List<Contact> objContacts
				, List<Residency_Program__c> objResidPrograms, List<Facility__c> objFacilities
				, string strInstitutionId, string strInstitName, string strContactId )
		{
			// attempt to find the resid. program tied to the institution
			Residency_Program__c objResidPrg = objResidPrograms.FirstOrDefault(
						rp => rp.Program_Institution__c != null && rp.Program_Institution__c.Equals( strInstitutionId ) );

			if( objResidPrg != null )
				// return the new id
				return objResidPrg.Id;
			else
			{
				//// retrieve institution data
				//Institution__c objInstit = objInstitutions.FirstOrDefault( i => i.Id.Equals( strInstitutionId ) );
				//Institution__c objInstit = objAPI.QuerySingle<Institution__c>( string.Concat(
				//		"select Id, Name from Institution__c where Id = '", strInstitutionId, "'" ) );

				// create resid. program tied to the institution
				objResidPrg = new Residency_Program__c();
				objResidPrg.Program_Institution__c = strInstitutionId;

				// see if there is a facility for this institution
				if( strInstitName.Contains( '\'' ) )
					strInstitName = strInstitName.Replace( "'", "\\\'" );
				if( strInstitName == null )
					strInstitName = "";

				Facility__c objFac = null;

				try
				{
					objFac = objFacilities.FirstOrDefault( f => f.Name.Contains( strInstitName ) );
					//// replace this with a list with all facilities instead of hitting SF every time
					//objFac = objAPI.QuerySingle<Facility__c>( string.Concat(
					//         "select Id, Name, Sub_Region__c, Service_line__c, "
					//         , "Sub_Region__r.Region_Code__r.Division_Code__c from Facility__c where Name like '"
					//         , strInstitName, "%' limit 1" ) );
				}
				catch( Exception excpt )
				{
					// make cursed error go away and let this dog of a process continue
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\n** ERROR: ", excpt.Message, " - "
						, excpt.Source, " - Stack Trace:  ", excpt.StackTrace );
				}

				if( objFac != null )
				{
					// and region, service line from facility (if available)
					objResidPrg.Assigned_Region__c = objFac.Sub_Region__r.Region_Code__r.Division_Code__c; // has to get division of facility
					objResidPrg.Service_Line__c = objFac.Service_line__c;
				}
				else
					objResidPrg.Service_Line__c = "Emergency Department";

				string strName = string.Concat( strInstitName, " - Resid. Program" );

				if( !objResidPrg.Program_Speciality__c.IsNullOrBlank() )
					strName = string.Concat( strName, "-", objResidPrg.Program_Speciality__c );
				objResidPrg.Name = strName.Left( 80 );

				// attempt to get Specialty__c from Contact
				Contact objProvider = objContacts.FirstOrDefault( c => c.Id.Equals( strContactId ) );
				if( objProvider != null )
					objResidPrg.Program_Speciality__c = objProvider.Specialty__c;

				// attempt to save the new resid. program and return the id
				SaveResult[] objResult = objAPI.Insert( new Residency_Program__c[] { objResidPrg } );
				if( objResult != null )
					if( objResult[ 0 ].success )
					{
						// store the Id and add the new program to the list
						objResidPrg.Id = objResult[ 0 ].id;
						objResidPrograms.Add( objResidPrg );
						return objResult[ 0 ].id;
					}
			}

			return "";	// if resid program was not created, return empty string
		}

		// define the sequence/hierarchy of stages
		public enum CandidateStage
		{
			Not_Hired = 0, Inactive, Lead, Phone_Screen, Applicant, Interview, Contract_Negotiation
			, Credentialing___SAP, Credentialing___SAR, Credentialing___SAH, Credentialing___Locums_Privileges
			, Credentialing___Active_Privileges, Credentialing___Never_Privileged, Credentialing, Scheduled
		};

		// define the sequence/hierarchy of statuses
		public enum CandidateStatus
		{
			Terminated_Working_TW = 0, Inactive_Recruiting_IR
			, Inactive_Credentialing_IC, Information_Only_IO, Actively_Recruiting_AR
			, Actively_Credentialing_AC, Actively_Working_AW, Proceed_With_Caution_PWC
			, Do_Not_Recruit_DNR, Duplicate_DUP, Bad_Address_BA
		};

		public List<Candidate__c> RefreshCandidates( List<Contact> objProviders = null
							, List<Facility__c> objFacilities = null, bool bDisplayOnly = true )
		{
			System.Diagnostics.Stopwatch objWatch = new System.Diagnostics.Stopwatch();
			objWatch.Start();

			DataTable objDT = null, objContractsDT = null, objInterviewsDT = null;

			if( objProviders == null || bMassDataLoad )
			{
				if( !bMassDataLoad )
					objProviders = Company2SFUtils.GetProvidersFromSF( objAPI, lblError
						, " Owning_Region__c, Owning_Candidate_Stage__c, Owning_Candidate_Status__c  "
						, " RecruitingID__c > 0 ", " RecruitingID__c " );

				objDT = objDB.GetDataTableFromSQLFile( "SQLAllCandidates.txt" );
				objContractsDT = objDB.GetDataTableFromSQLFile( "SQLAllContracts.txt" );
				objInterviewsDT = objDB.GetDataTableFromSQLFile( "SQLAllInterviews.txt" );
			}
			else
			{
				string strProviderScript = Company2SFUtils.CreateProviderScript( objProviders, true );
				objDT = objDB.GetDataTableFromSQLFile( "SQLCandidate.txt", strProviderScript );
				objContractsDT = objDB.GetDataTableFromSQLFile( "SQLContracts.txt", strProviderScript );
				objInterviewsDT = objDB.GetDataTableFromSQLFile( "SQLAllInterviews.txt", strProviderScript );
			}

			tbStatus.Text = string.Concat( objDT.Rows.Count, " candidate rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			if( objFacilities == null )
				objFacilities = objAPI.Query<Facility__c>(
						"select ID, Name, Site_Code__c, Region__c, Sub_Region__c, Sub_Region__r.SubRegionCode__c from Facility__c order by Site_Code__c" );

			List<EMSC2SF.User> objUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			// load candidates from datatable to agency list
			List<Candidate__c> objCandidates = new List<Candidate__c>( objDT.Rows.Count );
			List<Interview__c> objInterviews = new List<Interview__c>( objDT.Rows.Count );
			List<Contact> objProvidersToUpdate = new List<Contact>();
			List<Provider_Contract__c> objContracts = new List<Provider_Contract__c>();
			int iSkipped = 0;
			foreach( DataRow objDR in objDT.Rows )
			{
				string strSiteCode = "";
				bool bSkipThisRecord = false;
				double dblRecruitingID = 0;

				// generate candidate rows
				Candidate__c objNewCandidate = ProcessCandidateRecord( objProviders, objFacilities, objDT
					, objCandidates, objProvidersToUpdate
					, ref dblRecruitingID, ref strSiteCode
					, ref iSkipped, objDR, ref bSkipThisRecord );
				if( bSkipThisRecord )
					continue;

				// generate provider contract rows
				VerifyProviderContracts( objContractsDT, objContracts, objNewCandidate.Name, strSiteCode, dblRecruitingID );

				// generate interview rows
				VerifyInterviews( objInterviewsDT, objInterviews, objNewCandidate.Name, strSiteCode, dblRecruitingID );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " candidate rows skipped (Site Code or RecID mismatch).\r\n" );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Name", objCandidates.ToArray<sObject>() );
			//ReportStatus(objResults);

			ReportStatus( "Reporting errors for candidates" );

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objCandidates, objResults, tbStatus );

			// this is to avoid errors since Name is a formula (SalesForce is dumb!)
			foreach( Contact objProvider2Upd in objProvidersToUpdate )
				objProvider2Upd.Name = null;

			// update providers owning stage/status
			SaveResult[] objProviderResults = null;
			if( !bDisplayOnly )
				objProviderResults = objAPI.Update( objProvidersToUpdate.ToArray<sObject>() );
			// sometimes we get error of "can't update name"

			Company2SFUtils.ReportErrors( objProviderResults, tbStatus );

			// now that we have Candidate__c ids, hook the interviews to them
			iSkipped = 0;
			foreach( Interview__c objInterview in objInterviews )
			{
				string strCandidate = objInterview.Candidate__c;
				Candidate__c objFound = objCandidates.FirstOrDefault<Candidate__c>( i => i.Name == strCandidate );
				if( objFound != null )
					objInterview.Candidate__c = objFound.Id;
				else
					iSkipped++;

				ReportStatus( string.Concat( "Processed interview ", objInterview.Name ) );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, iSkipped, " interview rows skipped (Candidate record name mismatch).\r\n" );

			UpsertResult[] objInterviewResults = null;
			if( !bDisplayOnly )
				objInterviewResults = objAPI.Upsert( "Name", objInterviews.ToArray<sObject>() );

			ReportStatus( "Reporting errors for interviews" );

			// create CSV file / set the Ids in the list of interviews
			Company2SFUtils.SetIdsReportErrors( objInterviews, objInterviewResults, tbStatus );

			// we have Candidate__c ids, hook the contracts to them
			iSkipped = 0;
			SortedSet<string> strNames = new SortedSet<string>();
			foreach( Provider_Contract__c objContract in objContracts )
			{
				string strCandidate = objContract.Candidate__c;
				Candidate__c objFound = objCandidates.FirstOrDefault<Candidate__c>( i => i.Name == strCandidate );
				if( objFound == null )
				{
					iSkipped++;
					continue;
				}

				objContract.Candidate__c = objFound.Id;

				// keep record names unique
				if( strNames.Contains( objContract.Name ) )
					objContract.Name = string.Concat( objContract.Name, " #2" );
				strNames.Add( objContract.Name );

				ReportStatus( string.Concat( "Processed contract ", objContract.Name ) );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, iSkipped, " contract rows skipped (Candidate record name mismatch).\r\n" );

			UpsertResult[] objContractResults = null;
			if( !bDisplayOnly )
				objContractResults = objAPI.Upsert( "Name", objContracts.ToArray<sObject>() );

			ReportStatus( "Reporting errors for contracts" );

			// create CSV file / set the Ids in the list of interviews
			Company2SFUtils.SetIdsReportErrors( objContracts, objContractResults, tbStatus );

			objWatch.Stop();
			ReportStatus( string.Concat( "Finished loading candidates. Duration: "
							, objWatch.Elapsed.Hours, " hours ", objWatch.Elapsed.Minutes, " minutes." ) );

			return objCandidates;
		}

		public Facility__c FindFacilityByNameOrCode( List<Facility__c> objFacilities, string strName, string strSiteCode )
		{
			Facility__c objFound = null;
			if( ! strName.IsNullOrBlank() )
				objFound = objFacilities.FirstOrDefault( f => f.Name.IndexOf( strName ) >= 0 );
			if( objFound == null )
				objFound = objFacilities.FirstOrDefault(
									i => i.Site_Code__c != null && i.Site_Code__c.Equals( strSiteCode ) );

			return objFound;
		}

		private Candidate__c ProcessCandidateRecord( List<Contact> objProviders, List<Facility__c> objFacilities
			, DataTable objDT, List<Candidate__c> objCandidates
			, List<Contact> objProvidersToUpdate
			, ref double dblRecruitingID, ref string strSiteCode
			, ref int iSkipped, DataRow objDR, ref bool bSkipThisRecord )
		{
			// copy all datatable columns to credential agency object attributes
			Candidate__c objNewCandidate = objDR.ConvertTo<Candidate__c>();

			// try match 1st 15 characters
			string strName = objDR[ "HospitalName" ].ToString().Left( 15 );
			strSiteCode = objNewCandidate.Facility_Name__c;
			Facility__c objFacility = FindFacilityByNameOrCode( objFacilities, strName, strSiteCode );
			if( objFacility == null )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find hospital matching "
									, objDR[ "HospitalName" ].ToString(), " or site code <" 
									, strSiteCode, "> for ", objNewCandidate.Name );
				iSkipped++;
				//continue;
				bSkipThisRecord = true;
				return objNewCandidate;
			}
			objNewCandidate.Facility_Name__c = objFacility.Id;
			strSiteCode = objFacility.Site_Code__c;

			// link candidate to Contact/Provider using RecruitingID
			dblRecruitingID = Convert.ToDouble( objNewCandidate.Contact__c );
			double dblRecrID = dblRecruitingID;
			string strContactId = "";
			Contact objProvider = objProviders.FirstOrDefault( i => i.RecruitingID__c == dblRecrID );
			if( objProvider != null )
				strContactId = objProvider.Id;
			else // skip if provider was not found
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find recruiting ID ", objNewCandidate.Contact__c, " for ", objNewCandidate.Name );
				iSkipped++;
				//continue;
				bSkipThisRecord = true;
				return objNewCandidate;
			}
			objNewCandidate.Contact__c = strContactId;

			// correct the record name using the Contact and Facility names in order to avoid duplicates
			objNewCandidate.Name = string.Concat( objProvider.Name, " at ", objFacility.Name ).Left( 80 );

			// if candidate record name is duplicate (due to facilities with same name), make it different
			// by appending the facility short name and/or site code instead
			Candidate__c objDupeFound = objCandidates.FirstOrDefault( c => c.Name.Equals( objNewCandidate.Name ) );
			if( objDupeFound != null )
				if( objFacility.Short_Name__c != null )
					objNewCandidate.Name = string.Concat( objProvider.Name, " at ", objFacility.Short_Name__c, "-", objFacility.Site_Code__c ).Left( 80 );
				else
					objNewCandidate.Name = string.Concat( objProvider.Name, " at ", objFacility.Site_Code__c ).Left( 80 );

			//
			// The code above probably could be changed to append the site code AFTER the facility name
			// if the limit of 80 characters allows
			//

			// if the facility's region is the same as the provider's owning region, turn on the owning region flag
			if( objFacility.Sub_Region__r != null && objFacility.Sub_Region__r.SubRegionCode__c != null )
				objNewCandidate.Owning_Region__c = 
					( objFacility.Sub_Region__r.SubRegionCode__c.Equals( objProvider.Owning_Region__c ) );
			else
				objNewCandidate.Owning_Region__c = false;
			objNewCandidate.Owning_Region__cSpecified = true;

			if( (bool) objNewCandidate.Owning_Region__c )
			{
				string strProviderOwningStage = objProvider.Owning_Candidate_Stage__c;
				if( strProviderOwningStage != null )
				{
					if( strProviderOwningStage.Equals( "Credentialing SAH" ) )
						strProviderOwningStage = "Credentialing - SAH";
					if( strProviderOwningStage.Equals( "Credentialing SAP" ) )
						strProviderOwningStage = "Credentialing - SAP";
					if( strProviderOwningStage.Equals( "Credentialing SAR" ) )
						strProviderOwningStage = "Credentialing - SAR";
				}

				// check whether the candidate stage is higher than the current stage in the provider record
				CandidateStage eProviderStage = 0;
				if( strProviderOwningStage != null )
					eProviderStage = (CandidateStage) Enum.Parse( typeof( CandidateStage )
								, strProviderOwningStage.Replace( ' ', '_' ).Replace( '-', '_' ) );

				CandidateStage eCandidateStage = (CandidateStage) Enum.Parse( typeof( CandidateStage )
								, objNewCandidate.Candidate_Stage__c.Replace( ' ', '_' ).Replace( '-', '_' ) );

				if( eProviderStage < eCandidateStage || objProvider.Owning_Candidate_Stage__c == null )
					// set the owning stage if candidate stage is higher
					objProvider.Owning_Candidate_Stage__c = objNewCandidate.Candidate_Stage__c;

				// check whether the candidate status is higher than the current status in the provider record
				CandidateStatus eProviderStatus = 0;
				if( objProvider.Owning_Candidate_Status__c != null )
					Enum.TryParse<CandidateStatus>(
							objProvider.Owning_Candidate_Status__c.Replace( ' ', '_' ).Replace( "(", "" ).Replace( ")", "" )
							, out eProviderStatus );
				//eProviderStatus = (CandidateStatus) Enum.Parse( typeof( CandidateStatus )
				//         , objProvider.Owning_Candidate_Status__c.Replace( ' ', '_' ).Replace( "(", "" ).Replace( ")", "" ) );

				CandidateStatus eCandidateStatus = 0;
				Enum.TryParse<CandidateStatus>(
							objNewCandidate.Candidate_Status__c.Replace( ' ', '_' ).Replace( "(", "" ).Replace( ")", "" )
							, out eCandidateStatus );
				//eCandidateStatus = (CandidateStatus) Enum.Parse( typeof( CandidateStatus )
				//             , objNewCandidate.Candidate_Status__c.Replace( ' ', '_' ).Replace( "(", "" ).Replace( ")", "" ) );

				// set the owning status if candidate status is higher
				if( eProviderStatus < eCandidateStatus || objProvider.Owning_Candidate_Status__c == null )
					objProvider.Owning_Candidate_Status__c = objNewCandidate.Candidate_Status__c;

				// update or add to the list
				if( objProvidersToUpdate.Contains( objProvider ) )
					objProvidersToUpdate[ objProvidersToUpdate.IndexOf( objProvider ) ] = objProvider;
				else
					objProvidersToUpdate.Add( objProvider );
			}

			// fix time zone bug
			objNewCandidate.Application_Received__c = Company2SFUtils.FixTimeZoneBug( objNewCandidate.Application_Received__c );
			objNewCandidate.Application_Sent__c = Company2SFUtils.FixTimeZoneBug( objNewCandidate.Application_Sent__c );
			objNewCandidate.Interview_Actual_Date__c = null;// Company2SFUtils.FixTimeZoneBug( objNewCandidate.Interview_Actual_Date__c );
			objNewCandidate.First_Shift_Date__c = Company2SFUtils.FixTimeZoneBug( objNewCandidate.First_Shift_Date__c );
			objNewCandidate.Termination_Resignation_Effective_Date__c = Company2SFUtils.FixTimeZoneBug( objNewCandidate.Termination_Resignation_Effective_Date__c );

			// if dates are not null, flag them
			objNewCandidate.Application_Received__cSpecified = ( objNewCandidate.Application_Received__c != null );
			objNewCandidate.Application_Sent__cSpecified = ( objNewCandidate.Application_Sent__c != null );
			objNewCandidate.Interview_Actual_Date__cSpecified = ( objNewCandidate.Interview_Actual_Date__c != null );
			objNewCandidate.First_Shift_Date__cSpecified = ( objNewCandidate.First_Shift_Date__c != null );
			objNewCandidate.Termination_Resignation_Effective_Date__cSpecified = ( objNewCandidate.Termination_Resignation_Effective_Date__c != null );

			objCandidates.Add( objNewCandidate );

			ReportStatus( string.Concat( "Processed candidate ", objNewCandidate.Name
						, " (", objCandidates.Count, "/", objDT.Rows.Count, ")" ) );

			return objNewCandidate;
		}

		private static void VerifyInterviews( DataTable objInterviewsDT, List<Interview__c> objInterviews
			, string strCandidateName, string strSiteCode, double dblRecruitingID )
		{

			// add Provider Contract records to be created, if existing
			DataRow[] objInterviewsDRs = objInterviewsDT.Select( string.Concat(
							"RecruitingID__c = ", dblRecruitingID.ToString()
							, " AND Site_Code__c = '", strSiteCode, "'" ) );
			foreach( DataRow objInterviewDR in objInterviewsDRs )
			{
				Interview__c objNewInterview = null;

				// create interview row for director
				string strDate = objInterviewDR[ "InterviewSetupDIR" ].ToString();
				if( !strDate.Equals( "" ) )
				{
					objNewInterview = new Interview__c();

					objNewInterview.Interview_Actual_Date__c = Company2SFUtils.FixTimeZoneBug( Convert.ToDateTime( strDate ) );
					objNewInterview.Interview_Actual_Date__cSpecified = true;
					objNewInterview.Declined__c = objInterviewDR[ "Declined__c" ].ToString();
					objNewInterview.Interviewing_with_Role__c = "Medical Director";
					//objNewInterview.Position__c = objInterviewDR[ "Position__c" ].ToString();
					objNewInterview.Post_Interview_Comments__c = objInterviewDR[ "Post_Interview_Comments__c" ].ToString();
					objNewInterview.Name = string.Concat( strCandidateName, " on ", strDate ).Left( 80 );
					objNewInterview.Candidate__c = strCandidateName; // store name temporarily, get ids later

					objInterviews.Add( objNewInterview );
				}

				// create interview row for regional director
				strDate = objInterviewDR[ "InterviewSetupRMD" ].ToString();
				if( !strDate.Equals( "" ) )
				{
					objNewInterview = new Interview__c();
					objNewInterview.Interview_Actual_Date__c = Company2SFUtils.FixTimeZoneBug( Convert.ToDateTime( strDate ) );
					objNewInterview.Interview_Actual_Date__cSpecified = true;
					objNewInterview.Declined__c = objInterviewDR[ "Declined__c" ].ToString();
					objNewInterview.Interviewing_with_Role__c = "Regional Medical Director";
					//objNewInterview.Position__c = objInterviewDR[ "Position__c" ].ToString();
					objNewInterview.Post_Interview_Comments__c = objInterviewDR[ "Post_Interview_Comments__c" ].ToString();
					objNewInterview.Name = string.Concat( strCandidateName, " on ", strDate, " w/RMD" ).Left( 80 );
					objNewInterview.Candidate__c = strCandidateName; // store name temporarily, get ids later

					objInterviews.Add( objNewInterview );
				}
			}
		}

		private static void VerifyProviderContracts( DataTable objContractsDT, List<Provider_Contract__c> objContracts
			, string strCandidateName, string strSiteCode, double dblRecruitingID )
		{


			// add Provider Contract records to be created, if existing
			DataRow[] objContractDRs = objContractsDT.Select( string.Concat(
							"RecruitingID__c = ", dblRecruitingID.ToString()
							, " AND Site_Code__c = '", strSiteCode, "'" ) );
			foreach( DataRow objContractDR in objContractDRs )
			{
				Provider_Contract__c objNewContract = new Provider_Contract__c();
				DateTime dtContract = DateTime.Today;

				if( !objContractDR[ "Contract_Received__c" ].IsNullOrBlank() )
				{
					objNewContract.Date_Contract_Received__c = Convert.ToDateTime( objContractDR[ "Contract_Received__c" ] );
					dtContract = (DateTime) objNewContract.Date_Contract_Received__c;
				}

				if( !objContractDR[ "EffectiveDate" ].IsNullOrBlank() )
					objNewContract.Contract_Effective_Date__c = Convert.ToDateTime( objContractDR[ "EffectiveDate" ] );

				if( !objContractDR[ "Contract_Sent__c" ].IsNullOrBlank() )
				{
					objNewContract.Date_Contract_Sent__c = Convert.ToDateTime( objContractDR[ "Contract_Sent__c" ] );
					dtContract = (DateTime) objNewContract.Date_Contract_Sent__c;
				}

				// fix time zone bug
				objNewContract.Date_Contract_Received__c = Company2SFUtils.FixTimeZoneBug( objNewContract.Date_Contract_Received__c );
				objNewContract.Date_Contract_Sent__c = Company2SFUtils.FixTimeZoneBug( objNewContract.Date_Contract_Sent__c );
				objNewContract.Contract_Effective_Date__c = Company2SFUtils.FixTimeZoneBug( objNewContract.Contract_Effective_Date__c );
				dtContract = (DateTime) Company2SFUtils.FixTimeZoneBug( (DateTime?) dtContract );

				objNewContract.Date_Contract_Received__cSpecified = true;
				objNewContract.Date_Contract_Sent__cSpecified = true;
				objNewContract.Contract_Effective_Date__cSpecified = true;

				if( !objContractDR[ "NumHours" ].IsNullOrBlank() )
					objNewContract.Contracted_Hours__c = objContractDR[ "NumHours" ].ToString();

				// use contract sent date to make name distinct
				string strDate = dtContract.ToShortDateString();
				objNewContract.Name = string.Concat( strCandidateName, " on ", strDate ).Left( 80 );

				objNewContract.Candidate__c = strCandidateName; // store name temporarily, get ids later

				objContracts.Add( objNewContract );
			}
		}

		public List<EMSC2SF.User> RefreshUsers( bool bDisplayOnly = true )
		{
			DataTable objDT = objDB.GetDataTableFromSQLFile( "SQLAllUsers.txt" );

			tbStatus.Text = string.Concat( objDT.Rows.Count, " user rows retrieved.\r\n" );

			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			UserRole objRole = objAPI.QuerySingle<UserRole>( "select ID from UserRole where Name = 'Recruiter' " );
			Profile objProfile = objAPI.QuerySingle<Profile>( "select ID from Profile where Name = 'Recruiter Profile' " );

			List<EMSC2SF.User> objExistingUsers = Company2SFUtils.GetSFUserList( objAPI, lblError );

			List<EMSC2SF.User> objUsers = new List<EMSC2SF.User>();

			UpsertResult[] objUsersResults = null;

			EMSC2SF.User objFoundInSF = null;
			//EMSC2SF.User objU = new EMSC2SF.User();
			//objU.Phone = "214 912 0000";
			//objU.Title = "Recruiter";
			//objU.LastName = "Recro";
			//objU.FirstName = "Robot";
			//objU.Email = "recrobot@mailinator.com";
			//objU.Alias = "recrobot";
			//objU.CommunityNickname = "MegaRecruiter9000";
			//objU.UserRoleId = objRole.Id; // "00EE0000000cKzUMAU" = Recruiter in Prod
			//objU.Username = "recrobot@mailinator.com"; // has to be an email address
			//objU.EmailEncodingKey = "ISO-8859-1";
			//objU.TimeZoneSidKey = "America/Los_Angeles";
			//objU.LanguageLocaleKey = "en_US";
			//objU.LocaleSidKey = "en_US";
			//objU.ProfileId = objProfile.Id; // "00eE0000000gtmIIAQ" = Recruiter Profile in Prod
			//objUsers.Add(objU);
			foreach( DataRow objDR in objDT.Rows )
			{
				EMSC2SF.User objU = objDR.ConvertTo<EMSC2SF.User>(); // new EMSC2SF.User();

				// get first region in the list (it is the default)
				string strRegion	= objDR[ "Regions" ].ToString().Split( ' ' ).First();
				switch( strRegion )
				{
					case "1":
					case "4":
					case "9":
					case "11":
					case "13":
					case "304":
					case "008":
					case "19":
					case "068":
						objU.Sub_Region__c = ""; break;
					case "2":
					case "6":
					case "14":
					case "": // blanks go to Central
						objU.Sub_Region__c = "WEST"; break;
					case "7":
						objU.Sub_Region__c = "SOUTH"; break;
					default:
						objU.Sub_Region__c = strRegion; break;
				}

				// try finding user by username in Active Directory
				ADInformation objAD = new ADInformation( objU.Alias );
				if( !objAD.Found )
					// try finding by 1st and last name
					objAD = new ADInformation( objU.LastName, objU.FirstName );
				if( !objAD.Found )
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nUser not found: "
							, objU.Alias, " (", objU.LastName, ", ", objU.FirstName, ")" );
					continue;
				}

				// find the user in the list from SF
				objFoundInSF = objExistingUsers.FirstOrDefault( i => i.FirstName == objU.FirstName && i.LastName == objU.LastName );

				// skip non active users
				if( !objAD.Active )
				{
					// deactivate the user (can't delete)
					objU.IsActive = false;

					// if the user is in SF and deactivate the user
					// instead of skipping
					if( objFoundInSF == null || objFoundInSF.IsActive == false )
					{
						// user is not in SF and is inactive in Company, so we report and process next user
						tbStatus.Text = string.Concat( tbStatus.Text, "\r\nUser inactive: ", objDR.ToTabString() );
						continue;
					}
				}

				// if the user doesn't exist in SF, we can overwrite everything
				// otherwise, skip overwriting Alias, Username and other minor defaults
				if( objFoundInSF == null )
				{
					objU.UserPermissionsSFContentUser = true;
					objU.UserPermissionsSFContentUserSpecified = true;
					objU.UserRoleId = objRole.Id;					// "00EE0000000cKzUMAU" = Recruiter in Prod
					objU.EmailEncodingKey = "ISO-8859-1";
					objU.TimeZoneSidKey = "America/Los_Angeles";
					objU.LanguageLocaleKey = "en_US";
					objU.LocaleSidKey = "en_US";
					objU.ProfileId = objProfile.Id;					// "00eE0000000gtmIIAQ" = Recruiter Profile in Prod

					string strAlias = objAD.SAMAccountName;
					if( strAlias.Length > 8 )
					{
						// try first initial + last name no spaces
						strAlias = string.Concat( objAD.FirstName.First(), objAD.LastName.Replace( " ", "" ) );

						if( strAlias.Length > 8 )
							// if alias still too big, make the Alias with only initials + consonants (ex.:Gina Parzanese = GPrzns)
							strAlias = string.Concat( objAD.FirstName.First()
								, objAD.LastName.Replace( "a", "" ).Replace( "e", "" ).Replace( "i", "" )
											.Replace( "o", "" ).Replace( "u", "" ).Replace( " ", "" ) );

						strAlias = objU.Alias.Left( 8 );
					}
					else
						strAlias = objAD.SAMAccountName;

					objU.Alias = strAlias;

					objU.Username = objAD.mail;		// in SF, username is the email address
				}
				else
				{
					// copy SF required user settings from existing user record
					objU.UserPermissionsSFContentUser = objFoundInSF.UserPermissionsSFContentUser;
					objU.UserPermissionsSFContentUserSpecified = objFoundInSF.UserPermissionsSFContentUserSpecified;
					objU.UserRoleId = objFoundInSF.UserRoleId;
					objU.EmailEncodingKey = objFoundInSF.EmailEncodingKey;
					objU.TimeZoneSidKey = objFoundInSF.TimeZoneSidKey;
					objU.LanguageLocaleKey = objFoundInSF.LanguageLocaleKey;
					objU.LocaleSidKey = objFoundInSF.LocaleSidKey;
					objU.ProfileId = objFoundInSF.ProfileId;

					objU.Alias = objFoundInSF.Alias;
					objU.Username = objFoundInSF.Username;
				}

				if( objU.Username == null )
				{
					// username null can't be added
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nUsername is null for: ", objDR.ToTabString() );
					continue;
				}

				objU.Phone = objAD.TelephoneNumber.IsNullOrBlank() ? objU.Phone : objAD.TelephoneNumber;

				objU.Title = objAD.title.IsNullOrBlank() ? objU.Title : objAD.title;

				objU.Email = objAD.mail.IsNullOrBlank() ? objU.Email : objAD.mail;

				objU.Department = objAD.department;
				objU.LastName = objAD.LastName;
				objU.FirstName = objAD.FirstName;
				objU.CommunityNickname = objAD.LastFirstName;
				objU.CompanyName = objAD.company;

				// decide what user category to use depending on the instance
				switch( strInstance )
				{
					case "prod":
						objU.User_Category__c = "Normal";
						break;
					case "test1":
						objU.User_Category__c = "UAT";
						break;
					case "dev1":
						objU.User_Category__c = "Test";
						break;
				}

				// append .dev1 or .test1 to email if needed
				if( !strInstance.Equals( "prod" ) && !objU.Username.Contains( strInstance ) )
				{
					objU.Username = string.Concat( objU.Username, ".", strInstance );
					objU.Email = string.Concat( objU.Email, ".", strInstance );
				}

				objUsers.Add( objU );

				// send upsert in batches of 20 records
				if( ( objUsers.Count % 20 ) == 0 )
				{
					if( !bDisplayOnly )
						objUsersResults = objAPI.Upsert( "Company_Username__c", objUsers.ToArray<sObject>() );

					// create CSV file / set the Ids in the list of users
					Company2SFUtils.SetIdsReportErrors( objUsers, objUsersResults, tbStatus );

					objUsers.Clear();
				}

				//// for the status report
				//objDR[ "Email__c" ] = objU.Email;
			}


			if( !bDisplayOnly )
				objUsersResults = objAPI.Upsert( "Company_Username__c", objUsers.ToArray<sObject>() );

			// create CSV file / set the Ids in the list of users
			Company2SFUtils.SetIdsReportErrors( objUsers, objUsersResults, tbStatus );

			return objUsers;
		}

		public static int CompareInstitutions( Institution__c p1, Institution__c p2 )
		{
			// order by name, then city, then address
			if( p1.Name.Equals( p2.Name ) )
				if( p1.City__c.Equals( p2.City__c ) )
					return p1.Address1__c.CompareTo( p2.Address1__c );
				else
					return p1.City__c.CompareTo( p2.City__c );

			return p1.Name.CompareTo( p2.Name );
		}

		public List<Residency_Program__c> RefreshResidencyPrograms( List<Institution__c> objInstitutions = null, bool bDisplayOnly = true )
		{
			List<Residency_Program__c> objResidPrograms = new List<Residency_Program__c>();
			List<Institution__c> objNewInstitutions = new List<Institution__c>();

			if( objInstitutions == null )
				objInstitutions = objAPI.Query<Institution__c>(
				"select Id, Name, Address1__c, City__c, State__c from Institution__c order by Name, City__c" );
			else
				// sort to help with lookup (name, city, address)
				objInstitutions.Sort( ( p1, p2 ) => CompareInstitutions( p1, p2 ) );

			string strFileName = string.Concat( strAppPath, "CSV_ResidencyPrograms.csv" );
			// map columns to SF columns
			string strMapping = 
@"Program_ID__c,Name,Program_Institution__c,,Program_Type__c,,Program_Address_Line_1__c,Program_Address_Line_2__c"
+ @",,,Program_Zip_Code__c,Program_Speciality__c,Program_Director__c,Program_Contact_Phone__c,,Program_Contact_Email__c,,,,";

			// read Residency Programs from CSV file
			DataTable objDT = null;
			objDT = objDT.ReadFile( strFileName, strMapping, true );

			//"PROGRAM_ID","Primary_Name","Program_Name_ACGME2","PROGRAM_EXTRNAL_MAME","PROGRAM_TYPE","PROGRAM_SOURCE"
			// ,"Program_Address1","Program_Address2","Program_City","Program_State","Program_Zip","Program_Specialty"
			//,"Program_Director","Program_Phone","Program_Fax","Program_Email","Program_External_Address1"
			//,"Program_External_Address2","Program_External_City","Program_External_State"

			// attempt to link each program with an institution
			int iSkipped = 0;
			foreach( DataRow objDR in objDT.Rows )
			{
				Residency_Program__c objProgram = objDR.ConvertTo<Residency_Program__c>();

				string strName = Util.FirstNonNull( objProgram.Program_Institution__c
								, objDR[ "PROGRAM_EXTRNAL_MAME" ].ToString()
								, objProgram.Name.Replace( " Program", "" ) );

				string strAddress = objProgram.Program_Address_Line_1__c;
				string strCity = objDR[ "Program_City" ].ToString();
				string strState = objDR[ "Program_State" ].ToString();

				// convert institution code into lookup id
				Institution__c objInstitutionFound = Company2SFUtils.FindInstitution( objInstitutions, strName
													, objProgram.Name, strAddress, strCity, strState );

				string strId = "";
				if( objInstitutionFound == null )
				{
					// check whether the new institution is already collected to be inserted
					string strProgramName = objProgram.Name.Left( 80 );
					objInstitutionFound = objNewInstitutions.FirstOrDefault( i => i.Name.Equals( strProgramName ) );
					if( objInstitutionFound == null )
					{
						// create institution for inserting it later
						objInstitutionFound = new Institution__c();
						objInstitutionFound.Name = strProgramName;
						objInstitutionFound.Metaphone_Name__c = strProgramName.ToNormalizedMetaphone();
						objInstitutionFound.Code__c = objProgram.Name.Left( 100 );
						objInstitutionFound.Company_Agency_Match__c = objProgram.Name.Left( 50 );
						objInstitutionFound.Credential_Type__c = "Institution";
						objInstitutionFound.Address1__c = strAddress;
						objInstitutionFound.Metaphone_Address__c = strAddress.ToNormalizedMetaphone();
						objInstitutionFound.Address2__c = objProgram.Program_Address_Line_2__c;
						objInstitutionFound.City__c = strCity;
						objInstitutionFound.Metaphone_City__c = strCity.ToNormalizedMetaphone();
						objInstitutionFound.State__c = strState;
						objInstitutionFound.Zip__c = objProgram.Program_Zip_Code__c;
						objInstitutionFound.Fax__c = objDR[ "Program_Fax" ].ToString();
						objInstitutionFound.Phone__c = objProgram.Program_Contact_Phone__c;
						objInstitutionFound.Contact__c = objProgram.Program_Director__c;

						objNewInstitutions.Add( objInstitutionFound );
					}

					// blank value will be replaced later
					strId = "";

					iSkipped++;
				}
				else
					strId = objInstitutionFound.Id;

				if( objProgram.Program_Speciality__c.Equals( "" ) )
					objProgram.Name = objProgram.Name.Left( 80 );
				else
					objProgram.Name = string.Concat( objProgram.Name, "-", objProgram.Program_Speciality__c ).Left( 80 );

				objProgram.Program_Institution__c = strId;

				//tbStatus.Text = string.Concat( tbStatus.Text, "\r\nMatched:  ", strName
				//    , ", institution: ", objInstitutionFound.Name
				//    , ", city: ", objInstitutionFound.City__c, ", program: ", objProgram.Name, ", program city: ", strCity );

				objResidPrograms.Add( objProgram );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " rows with new institutions (Institution mismatch)." );

			// create the institutions
			UpsertResult[] objInstitutionResults = null;
			if( !bDisplayOnly )
				objInstitutionResults = objAPI.Upsert( "Name", objNewInstitutions.ToArray<sObject>() );

			// set the Ids
			Company2SFUtils.SetIdsReportErrors( objNewInstitutions, objInstitutionResults, tbStatus );

			// associate programs with the id of the new institutions
			foreach( Residency_Program__c objProgram in objResidPrograms.Where( p => p.Program_Institution__c.Equals( "" ) ) )
			{
				Institution__c objInstit = objNewInstitutions.FirstOrDefault( i => i.Name.Equals( objProgram.Name.Left( 80 ) ) );
				if( objInstit != null )
					objProgram.Program_Institution__c = objInstit.Id;
			}

			// upsert
			UpsertResult[] objResidProgramResults = null;
			if( !bDisplayOnly )
				objResidProgramResults = objAPI.Upsert( "Program_ID__c", objResidPrograms.ToArray<sObject>() );

			// set the Ids
			Company2SFUtils.SetIdsReportErrors( objResidPrograms, objResidProgramResults, tbStatus );

			return objResidPrograms;
		}

		public List<Resident__c> RefreshResidents( List<Residency_Program__c> objResidPrograms = null, bool bDisplayOnly = true )
		{
			string strFileName = string.Concat( strAppPath, "CSV_Residents.csv" );
			// map columns to SF columns
			string strMapping =
@",Program_ID__c,Name,LastName,FirstName,Degree__c,,,Gender__c,,,Email,,,,Address1__c,Address2__c"
+ @",City__c,State__c,Zipcode__c,YearOfBirth,Phone,,Fax";

			// read Residency Programs from CSV file
			DataTable objDT = null;
			objDT = objDT.ReadFile( strFileName, strMapping, true );

			if( objResidPrograms == null )
				objResidPrograms = objAPI.Query<Residency_Program__c>(
				"select Id, Program_Id__c, Name from Table that stores resident programs ORDER BY Program_Id__c" );

			List<Resident__c> objResidents = new List<Resident__c>();
			List<Contact> objNewProviders = new List<Contact>();
			List<Contact> objContacts = Company2SFUtils.GetProvidersFromSF( objAPI, lblError, " Birthdate ", "", " Name " );

			System.Text.RegularExpressions.Regex objRegEx = new System.Text.RegularExpressions.Regex(
						@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" );

			int iSkipped = 0;
			foreach( DataRow objDR in objDT.Rows )
			{
				Resident__c objResident = objDR.ConvertTo<Resident__c>();

				// find the residency program in which this resident participates
				string strProgramId = objDR[ "Program_Id__c" ].ToString();
				Residency_Program__c objProgFound = 
					objResidPrograms.FirstOrDefault( p => p.Program_ID__c != null && p.Program_ID__c.Equals( strProgramId ) );
				if( objProgFound == null )
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find program ID ", strProgramId );
					iSkipped++;
					continue;
				}

				string strFirstName = objDR[ "FirstName" ].ToString();
				string strMiddleName = "";
				int iPos = strFirstName.IndexOfAny( " -".ToCharArray() );
				if( iPos > 0 )
				{
					strMiddleName = strFirstName.Substring( iPos + 1 );
					strFirstName = strFirstName.Substring( 0, iPos );
				}
				string strLastName = objDR[ "LastName" ].ToString();

				// give an distinct name to the record and store the provider name to lookup the id later
				objResident.Name = string.Concat( strLastName, ", ", strFirstName, " f/program ", objProgFound.Name ).Left( 80 );

				// check whether this is a duplicate, if it is then skip it
				Resident__c objResidentFound = objResidents.FirstOrDefault( r => r.Name.Equals( objResident.Name ) );
				if( objResidentFound != null )
					continue;

				int iYearOfBirth = 0;
				Int32.TryParse( objDR[ "YearOfBirth" ].ToString(), out iYearOfBirth );

				// find a contact to associate the resident record
				Contact objProviderFound = objContacts.FirstOrDefault(
										c => c.FirstName != null && c.LastName != null
										 && c.FirstName.Equals( strFirstName ) && c.LastName.Equals( strLastName )
										 && ( c.Middle_Name__c == null || c.Middle_Name__c.Equals( strMiddleName ) )
										 && ( iYearOfBirth == 0 || c.Birthdate == null || iYearOfBirth == Convert.ToDateTime( c.Birthdate ).Year ) );
				if( objProviderFound == null )
					objProviderFound = objContacts.FirstOrDefault(
										c => c.FirstName != null && c.LastName != null
										 && c.FirstName.Equals( strFirstName ) && c.LastName.Equals( strLastName )
										 && ( iYearOfBirth == 0 || c.Birthdate == null || iYearOfBirth == Convert.ToDateTime( c.Birthdate ).Year ) );

				// if no contact found, create one
				string strId = "";
				if( objProviderFound == null )
				{
					// check whether provider is already on the list to be created
					objProviderFound = objNewProviders.FirstOrDefault( c =>
									c.FirstName.Equals( strFirstName ) && c.LastName.Equals( strLastName ) );
					if( objProviderFound == null )
					{
						// only allow valid emails
						string strEmail = objDR[ "Email" ].ToString().Replace( ",", "." ).Replace( "2", "@" ).Replace( "5", "@" )
							.Replace( ";", "" ).Replace( "bhs1.org", "@bhs1.org" ).Replace( " ", "" ).Replace( "@@", "@" );
						if( !objRegEx.IsMatch( strEmail ) )
							strEmail = "";

						// create contact
						objProviderFound = new Contact();
						//objProviderFound.Name = string.Concat( strFirstName, " ", strLastName );
						objProviderFound.FirstName = strFirstName;
						objProviderFound.Middle_Name__c = strMiddleName;
						objProviderFound.LastName = strLastName;
						objProviderFound.Degree__c = objDR[ "Degree__c" ].ToString();
						objProviderFound.Specialty__c = objResident.Residency_Program_Specialty__c;
						objProviderFound.Gender__c = objDR[ "Gender__c" ].ToString();
						objProviderFound.Email = strEmail;
						objProviderFound.Address_Line_1__c = objDR[ "Address1__c" ].ToString();
						objProviderFound.Address_Line_2__c = objDR[ "Address2__c" ].ToString();
						objProviderFound.City__c = objDR[ "City__c" ].ToString();
						objProviderFound.State__c = objDR[ "State__c" ].ToString();
						objProviderFound.Zipcode__c = objDR[ "Zipcode__c" ].ToString();
						objProviderFound.Phone = objDR[ "Phone" ].ToString();
						objProviderFound.Fax = objDR[ "Fax" ].ToString();
						objProviderFound.Description = string.Concat(
							strFirstName, " ", strMiddleName, " ", strLastName, " was imported for Residency Program ", objProgFound.Name, "." );

						// make zip valid
						if( objProviderFound.Zipcode__c.Length < 5 )
							objProviderFound.Zipcode__c = objProviderFound.Zipcode__c.PadLeft( 5, '0' );

						objNewProviders.Add( objProviderFound );
					}
				}
				else
				{
					strId = objProviderFound.Id;

					//// store middle name just in case it is one of those providers having records with and without middle name
					//objProviderFound.Middle_Name__c = strMiddleName;
				}

				objResident.Contact__c = strId;

				objResident.Residency_Program__c = objProgFound.Id;

				objResidents.Add( objResident );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " rows skipped due to missing residency program." );

			// create the contact records
			SaveResult[] objProvidersResults = null;
			if( !bDisplayOnly )
				objProvidersResults = objAPI.Insert( objNewProviders.ToArray<sObject>() );

			// report errors
			Company2SFUtils.ReportErrors( objProvidersResults, tbStatus );

			// associate residents with the id of the new contacts
			foreach( Resident__c objResident in objResidents.Where( r => r.Contact__c.Equals( "" ) ) )
			{
				// extract provider name from resident record name
				string strName = objResident.Name;
				int iPos = strName.IndexOf( " f/program " );
				if( iPos > 0 )
					strName = strName.Substring( 0, iPos );

				// try to find provider by first and last names
				string[] strNames = strName.Split( ",".ToCharArray() );
				Contact objContact = objNewProviders.FirstOrDefault(
						p => p.FirstName.Equals( strNames[ 1 ].Trim() ) && p.LastName.Equals( strNames[ 0 ].Trim() ) );
				if( objContact != null )
					objResident.Contact__c = objContact.Id;
			}

			// upsert
			UpsertResult[] objResidentsResults = null;
			if( !bDisplayOnly )
				objResidentsResults = objAPI.Upsert( "Name", objResidents.ToArray<sObject>() );

			// set the Ids
			Company2SFUtils.SetIdsReportErrors( objResidents, objResidentsResults, tbStatus );

			return objResidents;
		}

		public List<Reference__c> RefreshReferences( List<Contact> objProviders = null, bool bImportAllRecords = false, bool bDisplayOnly = true )
		{
			DataTable objDT = null;
			if( objProviders == null )
				objProviders = Company2SFUtils.GetProvidersFromSF( objAPI, lblError );

			if( bImportAllRecords )
				objDT = objDB.GetDataTableFromSQLFile( "SQLAllReferences.txt" );
			else
				objDT = objDB.GetDataTableFromSQLFile( "SQLReferences.txt", Company2SFUtils.CreateProviderScript( objProviders ) );

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDT.Rows.Count, " reference rows retrieved.\r\n" );
			if( !objDB.ErrorMessage.Equals( "" ) )
			{
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", objDB.ErrorMessage );
				return null;
			}

			SortedSet<string> strNames = new SortedSet<string>();

			List<Reference__c> objReferences = new List<Reference__c>();
			int iSkipped = 0;
			foreach( DataRow objDR in objDT.Rows )
			{
				Reference__c objNewReference = objDR.ConvertTo<Reference__c>();

				// link reference to Contact/Provider using PhysicianNumber/RecruitingId
				double dblPhysicianNumber = Convert.ToDouble( objNewReference.Contact__c );
				Contact objProvider = objProviders.FirstOrDefault( i =>
					i.PhysicianNumber__c == dblPhysicianNumber || i.RecruitingID__c == dblPhysicianNumber );

				// skip if provider was not found
				if( objProvider == null )
				{
					tbStatus.Text = string.Concat( tbStatus.Text, "\r\nCould not find physician nbr "
						, objNewReference.Contact__c, " for ", objNewReference.Name );
					iSkipped++;
					continue;
				}

				objNewReference.Contact__c = objProvider.Id;

				// attempt to keep distinct names by appending city to name
				if( strNames.Contains( objNewReference.Name ) )
					objNewReference.Name = string.Concat( objNewReference.Name, " at ", objNewReference.City__c );
				strNames.Add( objNewReference.Name );

				objReferences.Add( objNewReference );
			}

			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", iSkipped, " rows skipped (provider not found)." );

			UpsertResult[] objResults = null;
			if( !bDisplayOnly )
				objResults = objAPI.Upsert( "Name", objReferences.ToArray<sObject>() );

			// create CSV file / set the Ids in the list of candidates
			Company2SFUtils.SetIdsReportErrors( objReferences, objResults, tbStatus );

			return objReferences;
		}

	}
}
