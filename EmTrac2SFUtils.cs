using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EmTrac2SF.EMSC2SF;
using GenericLibrary;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Data;
using EmTrac2SF.Salesforce;
using System.Text;
using System.IO;

namespace EmTrac2SF
{
	public static class EmTrac2SFUtils
	{
		public static string strPath = HttpContext.Current.Request.PhysicalApplicationPath;
		public static string strHistoryFileName = string.Concat( EmTrac2SFUtils.strPath, "ActivityHistory.txt" );

		public static bool bMassDataLoad = false;

		// 450 bytes per row will be preallocated (estimate taken from Results_Contact.csv:  25 Mb for 59k rows)
		public static int iTypicalRowSize = 450;

		public static HtmlGenericControl OpenCSVList;

		public static bool ReportStatus( params string[] strStatus )
		{
			// store status in ActivityHistory.txt			
			if( !File.Exists( strHistoryFileName ) )
				// create the file with the status
				File.WriteAllText( strHistoryFileName, string.Concat( "Activity History created on ", DateTime.Now.ToString()
					, "\r\n\r\n", string.Concat( strStatus ) ) );
			else
				AppendTextToHistoryFile( string.Concat( "\r\n", string.Concat( strStatus ) ) );

			//StringBuilder strbStatusList = new StringBuilder( "", 5700 );
			//if( Session[ "Status" ] != null )
			//    strbStatusList.Append( Session[ "Status" ].ToString() );

			//strbStatusList.Append( "<br />" );
			//strbStatusList.Append( strStatus );
			//// truncate to the last 4000 characters when the list reaches 5500 (to avoid doing it everytime)
			//if( strbStatusList.Length > 5500 )
			//    strbStatusList.Remove( 0, strbStatusList.Length - 4000 );
			//Session[ "Status" ] = strbStatusList.ToString();

			//// publish status at certain periods and enclose it in HTML
			//if( strbStatusList.Length > 4100 || strbStatusList.Length < 100 )
			//{
			//    strbStatusList.Insert( 0
			//        , "<html><head><title>EmTrac2SF STATUS</title></head><body style='font-size: 8pt; font-family:Verdana;'>" );
			//    strbStatusList.Append( "</body></html>" );
			//    strFileName = string.Concat( EmTrac2SFUtils.strPath, "Status.html" );
			//    File.WriteAllText( strFileName, strbStatusList.ToString() );
			//}

			return true;
		}

		private static void AppendTextToHistoryFile( string strText )
		{
			int iRetries = 0;
			while( iRetries < 3 )
			{
				try
				{
					// append error list to the file
					File.AppendAllText( strHistoryFileName, strText );

					// if operation was completed then leave loop
					break;
				}
				catch( IOException )
				{
					// count and try again
					iRetries++;
				}
			}
		}

		public static string ValidatePostalCode( string strZipCode )
		{
			// validate zip
			if( strZipCode != null )
			{
				strZipCode = strZipCode.Replace( "/", "-" ).Replace( ",", "" );

				// if zip code is like 99999-, remove dash
				if( strZipCode.Contains( "-" ) && strZipCode.Length < 10 && strZipCode.Length > 5 )
					strZipCode = strZipCode.Substring( 0, 5 );
				// if zip code is like 999999, truncate it to 1st 5 digits
				if( strZipCode.Length > 5 && strZipCode.Length < 10 )
					strZipCode = strZipCode.Substring( 0, 5 );

				// to do:  handle Canadian zip codes (alphanum 3+space+3 characters)
				if( strZipCode.HasLetters() )
					strZipCode = "";
				if( strZipCode.Length < 5 )
					strZipCode = "";
			}

			return strZipCode;
		}

		public static string ValidPhone( System.Text.RegularExpressions.Regex objRegExPhone, string strPhone )
		{
			if( strPhone == null ) return "";

			System.Text.RegularExpressions.MatchCollection objMatch = objRegExPhone.Matches( strPhone );

			if( objMatch.Count > 0 )
				return objMatch[ 0 ].Value;

			return "";
		}

		public static List<EMSC2SF.User> GetSFUserList( ApiService objAPI, Label lblError, bool bForUpdate = false )
		{
			string strSOQL	= "select ID, Name, Username, LastName, FirstName, Email, IsActive, EmTrac_Username__c, EmTrac_Usercode__c, Sub_Region__c from User order by EmTrac_Username__c";
			if( bForUpdate )	// SOQL for updates cannot have Name in the select list
				strSOQL = "select ID, Username, Email, IsActive from User order by EmTrac_Username__c";

			List<EMSC2SF.User> objExistingUsers = objAPI.Query<EMSC2SF.User>( strSOQL );
			if( !objAPI.ErrorMessage.Equals( "" ) )
			{
				lblError.Text = objAPI.ErrorMessage;
				return null;
			}

			return objExistingUsers;
		}

		public static string ConvertToUserId( TextBox tbStatus, List<EMSC2SF.User> objExistingUsers, string strUser )
		{
			if( strUser.IsNullOrBlank() )
				return "";

			string strUserName = strUser.ToLower();
			User objFound = objExistingUsers.FirstOrDefault(
							i => ( i.EmTrac_Username__c != null ) ? i.EmTrac_Username__c.ToLower().Equals( strUserName ) : false );
			strUserName = objFound != null ? objFound.Id : "";

			if( strUserName.Equals( "" ) )
				tbStatus.Text = string.Concat( tbStatus.Text, "\r\n", strUser, " could not be found in SalesForce." );

			return strUserName;
		}

		public static void FlagDupesByAddressCityName( DataTable objDT )
		{
			// create columns to link duplicates to a single record
			DataColumn objDC = null;

			string strExceptions = "American Board of Surgery";

			string[] strCols = { "Address1__c", "Address2__c", "City__c", "OriginalName" };

			foreach( string strColumn in strCols )
			{
				string strTargetCol = string.Concat( "Match", strColumn );
				if( !objDT.Columns.Contains( strTargetCol ) )
				{
					objDC = new DataColumn( strTargetCol, typeof( string ) );
					objDT.Columns.Add( objDC );
				}
			}

			if( !objDT.Columns.Contains( "DuplicateOfRowNumber" ) )
			{
				objDC = new DataColumn( "DuplicateOfRowNumber", typeof( DataRow ) );
				objDT.Columns.Add( objDC );
			}

			if( !objDT.Columns.Contains( "DuplicateOfCode" ) )
			{
				objDC = new DataColumn( "DuplicateOfCode", typeof( string ) );
				objDT.Columns.Add( objDC );
			}

			if( !objDT.Columns.Contains( "DuplicateOfName" ) )
			{
				objDC = new DataColumn( "DuplicateOfName", typeof( string ) );
				objDT.Columns.Add( objDC );
			}

			if( !objDT.Columns.Contains( "ModifiedName" ) )
			{
				objDC = new DataColumn( "ModifiedName", typeof( string ) );
				objDT.Columns.Add( objDC );
			}

			System.Text.RegularExpressions.Regex objRegex = new System.Text.RegularExpressions.Regex( @"\d+" );

			// remove/replace non-alpha characters/spaces and convert to lowercase to help match names and find duplicates
			foreach( DataRow objDR in objDT.Rows )
			{
				foreach( string strColumn in strCols )
				{
					// get only what is before the dash/comma
					string strValueToMatch = Util.TrimUpToSeparator( objDR[ strColumn ].ToString() );

					// extract numbers from the column to match, if needed
					string strNumbers = Util.GetNumbersInString( strValueToMatch );

					string strMetaphoneValue = string.Concat( strNumbers, Util.NormalizeForMatching( strValueToMatch ) );
					string strTargetCol = string.Concat( "Match", strColumn );

					objDR[ strTargetCol ] = strMetaphoneValue.Trim();
				}

				// TO DO:  add code to normalize addresses

				// TO DO:  add code to give the record a score according to completion, accuracy, age or other criteria
			}

			// flag duplicates - the 1st institution/agency in a group of duplicates will be the "1st match" 
			// to which all others will refer to

			string strColumnList = "Credential_Type__c,MatchCity__c,MatchAddress1__c,MatchAddress2__c,MatchOriginalName,State__c,Code__c";
			//if( objDT.Columns.Contains( "Credential_Type__c" ) )
			//    strColumnList = string.Concat( "Credential_Type__c,", strColumnList );
			//if( objDT.Columns.Contains( "State__c" ) )
			//    strColumnList = string.Concat( strColumnList, ",State__c" );
			//if( objDT.Columns.Contains( "Code__c" ) )
			//    strColumnList = string.Concat( strColumnList, ",Code__c" );

			// 1ST PASS - only look at address/city			
			DataView objDV = new DataView( objDT, "", strColumnList, DataViewRowState.CurrentRows );
			int i1stMatch = 0;
			string str1stMatchOriginalName = "", str1stMatchAddress1__c = "", str1stMatchCity__c = "";
			for( int iIndex = 0; iIndex < objDV.Count; iIndex++ )
			{
				string strName = objDV[ iIndex ][ "MatchOriginalName" ].ToString();
				string strAddress = objDV[ iIndex ][ "MatchAddress1__c" ].ToString();
				string strAddress2 = objDV[ iIndex ][ "MatchAddress2__c" ].ToString();
				string strCity = objDV[ iIndex ][ "MatchCity__c" ].ToString();

				// attempt to match address line #1 or #2
				bool bAddressesMatch = strAddress.Equals( str1stMatchAddress1__c )
									|| ( strAddress2.Equals( str1stMatchAddress1__c )
											&& !strAddress2.Equals( "" ) );
				//|| string.Concat( strAddress, strAddress2 ).Equals( str1stMatchAddress1__c ) );

				// if address+city match, consider it a duplicate (skip blank addresses/cities)
				if( !strExceptions.Equals( objDV[ iIndex ][ "OriginalName" ] )
					&& bAddressesMatch
					&& strCity.Equals( str1stMatchCity__c )
					&& !strAddress.Equals( "" ) && !strCity.Equals( "" ) )
				{
					// if first match has blank fields, copy available data from duplicates
					if( objDV[ i1stMatch ][ "Contact__c" ].IsNullOrBlank() )
						if( !objDV[ iIndex ][ "Contact__c" ].IsNullOrBlank() )
							objDV[ i1stMatch ][ "Contact__c" ] = objDV[ iIndex ][ "Contact__c" ];
					if( objDV[ i1stMatch ][ "Phone__c" ].IsNullOrBlank() )
						if( !objDV[ iIndex ][ "Phone__c" ].IsNullOrBlank() )
							objDV[ i1stMatch ][ "Phone__c" ] = objDV[ iIndex ][ "Phone__c" ];
					if( objDV[ i1stMatch ][ "Fax__c" ].IsNullOrBlank() )
						if( !objDV[ iIndex ][ "Fax__c" ].IsNullOrBlank() )
							objDV[ i1stMatch ][ "Fax__c" ] = objDV[ iIndex ][ "Fax__c" ];

					// set the DuplicateOfCode pointing to the first match
					// this will be used to convert any agencies/institutions in a group of duplicates 
					// to a single EmTrac_Agency_Match__c (in the 1st match)
					objDV[ iIndex ][ "DuplicateOfCode" ] = objDV[ i1stMatch ][ "EmTrac_Agency_Match__c" ];
					objDV[ iIndex ][ "DuplicateOfName" ] = objDV[ i1stMatch ][ "OriginalName" ];
					objDV[ iIndex ][ "DuplicateOfRowNumber" ] = objDV[ i1stMatch ].Row;

					continue;
				}
				else
				{
					// name is different so flag it as the next 1st match to be compared to others
					i1stMatch = iIndex;
					str1stMatchOriginalName = strName;
					str1stMatchAddress1__c = strAddress;
					str1stMatchCity__c = strCity;
				}
			}

			// flag duplicates (2ND PASS) - consider the name

			strColumnList = "Credential_Type__c,MatchOriginalName,MatchCity__c,MatchAddress1__c,State__c,Code__c";
			//if( objDT.Columns.Contains( "Credential_Type__c" ) )
			//    strColumnList = string.Concat( "", strColumnList );
			//if( objDT.Columns.Contains( "State__c" ) )
			//    strColumnList = string.Concat( strColumnList, ",State__c" );
			//if( objDT.Columns.Contains( "Code__c" ) )
			//    strColumnList = string.Concat( strColumnList, ",Code__c" );
			objDV = new DataView( objDT, "", strColumnList, DataViewRowState.CurrentRows );
			i1stMatch = 0;
			str1stMatchOriginalName = ""; str1stMatchAddress1__c = ""; str1stMatchCity__c = "";
			for( int iIndex = 0; iIndex < objDV.Count; iIndex++ )
			{
				string strName = objDV[ iIndex ][ "MatchOriginalName" ].ToString();
				string strAddress = objDV[ iIndex ][ "MatchAddress1__c" ].ToString();
				string strCity = objDV[ iIndex ][ "MatchCity__c" ].ToString();

				bool bPotentialNew1stMatch = false;

				// if name matches and city differs, then it could be a coincidence of metaphone (not a duplicate)
				// example:  CNA Insurance Company vs Keane Insurance Co. = same metaphone value KN
				// or it should be a branch agency in another city (not a duplicate)

				// thus we only consider a duplicate if name and city matches, that is, 
				// agencies with many addresses in a same city are considered as duplicates
				if( !strExceptions.Equals( objDV[ iIndex ][ "OriginalName" ] )
					&& strName.Equals( str1stMatchOriginalName )
					&& !strAddress.Equals( str1stMatchAddress1__c ) )
				{
					// this will only be used to differentiatte institutions with same name in different cities
					objDV[ iIndex ][ "ModifiedName" ] = "";

					// name is duplicate but if it is located in a different city, append city to name and consider it distinct
					string str1stMatchCity = objDV[ i1stMatch ][ "City__c" ].ToString();
					string strCurrentCity = objDV[ iIndex ][ "City__c" ].ToString();
					if( !strCity.Equals( str1stMatchCity__c ) ) //strCurrentCity.ToLower().Equals( str1stMatchCity.ToLower() ) )
					{
						// make current record name distinct by appending the city to the name, this will be the new 1st match
						str1stMatchOriginalName = string.Concat( objDV[ iIndex ][ "OriginalName" ].ToString(), " - ", strCurrentCity );
						objDV[ iIndex ][ "Name" ] = str1stMatchOriginalName;
						//objDV[ iIndex ][ "MatchOriginalName" ] = string.Concat( strName, " ", strCurrentCity );
						objDV[ iIndex ][ "ModifiedName" ] = str1stMatchOriginalName;

						// append the city to the name in the first matching record too
						objDV[ i1stMatch ][ "Name" ] = string.Concat( objDV[ i1stMatch ][ "OriginalName" ].ToString(), " - ", str1stMatchCity );
						//objDV[ i1stMatch ][ "MatchOriginalName" ] = string.Concat( strName, " ", str1stMatchCity );
						objDV[ i1stMatch ][ "ModifiedName" ] = objDV[ i1stMatch ][ "Name" ];

						// flag this record as possible new unique entry, to be checked if already flagged as 
						// duplicate in the first pass
						bPotentialNew1stMatch = true;
					}
					else
					{

						// if first match has blank address, copy data from duplicate if available
						if( objDV[ i1stMatch ][ "Address1__c" ].IsNullOrBlank() )
						{
							if( !objDV[ iIndex ][ "Address1__c" ].IsNullOrBlank() )
							{
								objDV[ i1stMatch ][ "Address1__c" ] = objDV[ iIndex ][ "Address1__c" ];
								objDV[ i1stMatch ][ "Address2__c" ] = objDV[ iIndex ][ "Address2__c" ];
								objDV[ i1stMatch ][ "City__c" ] = objDV[ iIndex ][ "City__c" ];
								objDV[ i1stMatch ][ "State__c" ] = objDV[ iIndex ][ "State__c" ];
								objDV[ i1stMatch ][ "Zip__c" ] = objDV[ iIndex ][ "Zip__c" ];

								// update new match values
								str1stMatchAddress1__c = objDV[ iIndex ][ "MatchAddress1__c" ].ToString();
								str1stMatchCity__c = objDV[ iIndex ][ "MatchCity__c" ].ToString();
							}
						}

						// if first match has blank fields, copy available data from duplicates
						if( objDV[ i1stMatch ][ "Contact__c" ].IsNullOrBlank() )
							if( !objDV[ iIndex ][ "Contact__c" ].IsNullOrBlank() )
								objDV[ i1stMatch ][ "Contact__c" ] = objDV[ iIndex ][ "Contact__c" ];
						if( objDV[ i1stMatch ][ "Phone__c" ].IsNullOrBlank() )
							if( !objDV[ iIndex ][ "Phone__c" ].IsNullOrBlank() )
								objDV[ i1stMatch ][ "Phone__c" ] = objDV[ iIndex ][ "Phone__c" ];
						if( objDV[ i1stMatch ][ "Fax__c" ].IsNullOrBlank() )
							if( !objDV[ iIndex ][ "Fax__c" ].IsNullOrBlank() )
								objDV[ i1stMatch ][ "Fax__c" ] = objDV[ iIndex ][ "Fax__c" ];

						// this record can't be unique anymore, so we check whether there 
						// are records pointing to this current one and redirect them
						string strFalseUnique = objDV[ iIndex ][ "EmTrac_Agency_Match__c" ].ToString();
						string strRealUnique = objDV[ i1stMatch ][ "EmTrac_Agency_Match__c" ].ToString();
						CorrectFalseUniques( objDV, strFalseUnique, strRealUnique );

						objDV[ iIndex ][ "DuplicateOfCode" ] = objDV[ i1stMatch ][ "EmTrac_Agency_Match__c" ];
						objDV[ iIndex ][ "DuplicateOfName" ] = objDV[ i1stMatch ][ "OriginalName" ];
						objDV[ iIndex ][ "DuplicateOfRowNumber" ] = objDV[ i1stMatch ].Row;
					}
				}
				else
					bPotentialNew1stMatch = true;

				// checked whether record was already flagged as duplicate in the first pass
				if( bPotentialNew1stMatch )
				{
					// if already flagged as duplicate in the 1st pass, it can't be a 1st match anymore
					// so we redirect all other records that were pointing to this one
					if( !objDV[ iIndex ][ "DuplicateOfCode" ].ToString().Equals( "" ) )
					{
						// redirect all duplicate records that are pointing to other duplicate records
						string strFalseUnique = objDV[ iIndex ][ "EmTrac_Agency_Match__c" ].ToString();

						// follow the linked chain of DataRows to find the ultimate unique row
						DataRow	drRealUnique = (DataRow) objDV[ iIndex ][ "DuplicateOfRowNumber" ];
						while( !Convert.IsDBNull( drRealUnique[ "DuplicateOfRowNumber" ] ) )
							drRealUnique = ( (DataRow) drRealUnique[ "DuplicateOfRowNumber" ] );

						// get the ultimate unique id
						string strRealUnique = drRealUnique[ "EmTrac_Agency_Match__c" ].ToString();

						if( !strRealUnique.Equals( "" ) )
							CorrectFalseUniques( objDV, strFalseUnique, strRealUnique );
					}
					else
					{
						// name is different so flag it as the next 1st match to be compared to subsequent records
						i1stMatch = iIndex;
						str1stMatchOriginalName = strName;
						str1stMatchAddress1__c = strAddress;
						str1stMatchCity__c = strCity;
					}
				}
			}

			objDV.Dispose();

			return;
		}

		public static void CorrectFalseUniques( DataView objDV, string strFalseUnique, string strRealUnique )
		{
			for( int iSubIndex = 0; iSubIndex < objDV.Count; iSubIndex++ )
				if( objDV[ iSubIndex ][ "DuplicateOfCode" ].ToString().Equals( strFalseUnique ) )
					objDV[ iSubIndex ][ "DuplicateOfCode" ] = strRealUnique;
		}

		/// <summary>
		/// Retrieves Contact Id, Name, Phys Nbr, Recr ID, First/Last/Middle Name and additional columns if specified
		/// </summary>
		/// <param name="objAPI"></param>
		/// <param name="lblError"></param>
		/// <param name="strAdditionalColumns">A single column name or a comma separated list of columns</param>
		/// <param name="strWhere">A condition for the WHERE clause</param>
		/// <param name="strOrder">A single column name or a comma separated list of columns for the ORDER BY clause</param>
		/// <returns></returns>
		public static List<Contact> GetProvidersFromSF( ApiService objAPI, Label lblError, string strAdditionalColumns = null
			, string strWhere = null, string strOrder = null )
		{
			string strSQL = "SELECT Id, PhysicianNumber__c, RecruitingID__c, Name, FirstName, LastName, Middle_Name__c ";

			if( strWhere.IsNullOrBlank() )
				strWhere = "";
			else
				strWhere = string.Concat( " WHERE ", strWhere );

			if( strOrder == null ) strOrder = " PhysicianNumber__c ";

			if( strAdditionalColumns.IsNullOrBlank() )
				strSQL = string.Concat( strSQL, " FROM Contact ", strWhere, " ORDER BY ", strOrder );
			else
				strSQL = string.Concat( strSQL, ", ", strAdditionalColumns, " FROM Contact ", strWhere, " ORDER BY ", strOrder );

			List<Contact> objContacts = objAPI.Query<Contact>( strSQL );

			return objContacts;
		}

		public static Institution__c FindInstitution( List<Institution__c> objInstitutions, string strName, string strProgram
			, string strAddress, string strCity, string strState )
		{
			Institution__c obj2ndMatch = null, obj3rdMatch = null, obj4thMatch = null;

			// run a battery of tests with each institution trying to find a match
			foreach( Institution__c i in objInstitutions )
			{
				// if a perfect match was found, return it immediately
				if( i.Name.IsEqualOrPartiallyMatchedTo( strName ) && strCity.NullAwareEquals( i.City__c ) )
					return i;

				// save the comparison result for optimization
				bool bCityIsMetaphoneMatched = strCity.IsMetaphoneMatchedTo( i.City__c );

				// if no match found yet, attempt to find a 2nd next best match with metaphone
				if( obj2ndMatch == null && i.Name.IsMetaphoneMatchedTo( strName ) && bCityIsMetaphoneMatched )
					obj2ndMatch = i;

				// if no match found yet, attempt to find by address, city and state
				if( obj2ndMatch == null && obj3rdMatch == null
						&& strAddress.IsMetaphoneMatchedTo( i.Address1__c ) && bCityIsMetaphoneMatched
						&& strState.NullAwareEquals( i.State__c ) )
					obj3rdMatch = i;

				// if no match found yet, attempt to find by truncating the program name and broad matching by state instead of city
				if( obj2ndMatch == null && obj3rdMatch == null && obj4thMatch == null )
				{
					string strProgramName = strProgram.Replace( " Program", "" );
					// if institution name is the same as the program name, no need to attempt again
					if( !strProgramName.Equals( strName ) )
						if( i.Name.IsMetaphoneMatchedTo( strProgramName ) && bCityIsMetaphoneMatched )
							obj4thMatch = i;
				}
			}

			// return next best match or null
			if( obj2ndMatch == null )
				if( obj3rdMatch == null )
					if( obj4thMatch == null )
						return null;
					else
						return obj4thMatch;
				else
					return obj3rdMatch;
			else
				return obj2ndMatch;
		}

		public static void SaveStatusCSV( string strStatus, TextBox tbStatus, bool bOnlyError = false )
		{
			// create file name
			string strFileName = "";

			// save string in a file
			if( !bOnlyError )
			{
				strFileName = string.Concat( strPath, "MassDataLoadResults.csv" );
				System.IO.File.WriteAllText( strFileName, strStatus );
			}

			// create error file
			strFileName = string.Concat( strPath, "MassDataLoadErrors.csv" );

			// save string in a file
			System.IO.File.WriteAllText( strFileName, tbStatus.Text );
		}

		public static void UpdateStatus<T>( ref string strStatus, ref int iCount, ref int iTotalCount
			, List<T> objProviders, string strObjectName, TextBox tbStatus )
		{
			iCount = objProviders == null ? 0 : objProviders.Count;
			iTotalCount += iCount;
			strStatus = string.Concat( strStatus, "\r\n ", iCount, "\t ", strObjectName, "\t retrieved: \t", DateTime.Now.ToString() );

			SaveStatusCSV( strStatus, tbStatus );
		}

		public static DateTime? FixTimeZoneBug( DateTime? ndtValue )
		{
			if( ndtValue == null ) return null;

			DateTime dtValue = (DateTime) ndtValue;
			TimeSpan ts5Hours = new TimeSpan( 0, 5, 0, 0, 0 );
			return dtValue.Add( ts5Hours );
		}

		public static string FilterValidEmail( string strEmail )
		{
			int iDotPos = strEmail.LastIndexOf( "." );
			int iAtPos = strEmail.LastIndexOf( "@" );
			return ( iDotPos > 0 && iAtPos > 0 && iDotPos > iAtPos ) ? strEmail : "";

		}

		public static int DeleteTable( ApiService objAPI, string strTable, TextBox tbStatus )
		{
			string strSOQL = string.Concat( "SELECT Id FROM ", strTable );
			DeleteResult[] objResult = null;
			switch( strTable )
			{
				case "Contact":
					objResult = objAPI.QueryAndDelete<Contact>( strSOQL ); break;
				case "Facility__c":
					objResult = objAPI.QueryAndDelete<Facility__c>( strSOQL ); break;
				case "Credential__c":
					objResult = objAPI.QueryAndDelete<Credential__c>( strSOQL ); break;
				case "Education_or_Experience__c":
					objResult = objAPI.QueryAndDelete<Education_or_Experience__c>( strSOQL ); break;
				case "Credential_Agency__c":
					objResult = objAPI.QueryAndDelete<Credential_Agency__c>( strSOQL ); break;
				case "Institution__c":
					objResult = objAPI.QueryAndDelete<Institution__c>( strSOQL ); break;
				case "Credential_Subtype__c":
					objResult = objAPI.QueryAndDelete<Credential_Subtype__c>( strSOQL ); break;
				case "Candidate__c":
					objResult = objAPI.QueryAndDelete<Candidate__c>( strSOQL ); break;
				case "Residency_Program__c":
					objResult = objAPI.QueryAndDelete<Residency_Program__c>( strSOQL ); break;
				case "Resident__c":
					objResult = objAPI.QueryAndDelete<Resident__c>( strSOQL ); break;
				case "Reference__c":
					objResult = objAPI.QueryAndDelete<Reference__c>( strSOQL ); break;
				case "Job_Application__c":
					objResult = objAPI.QueryAndDelete<Job_Application__c>( strSOQL ); break;
				case "ContactFeed":
					objResult = objAPI.QueryAndDelete<ContactFeed>( strSOQL ); break;
				case "Facility__Feed":
					objResult = objAPI.QueryAndDelete<Facility__Feed>( strSOQL ); break;
				case "Candidate_Stage_Tracking__c":
					objResult = objAPI.QueryAndDelete<Candidate_Stage_Tracking__c>( strSOQL ); break;
				case "Provider_Contract__Feed":
					objResult = objAPI.QueryAndDelete<Provider_Contract__Feed>( strSOQL ); break;
				case "Provider_Contract__History":
					objResult = objAPI.QueryAndDelete<Provider_Contract__History>( strSOQL ); break;
				case "Candidate__Feed":
					objResult = objAPI.QueryAndDelete<Candidate__Feed>( strSOQL ); break;
				case "Reference__History":
					objResult = objAPI.QueryAndDelete<Reference__History>( strSOQL ); break;
			}

			// report results
			EmTrac2SFUtils.ReportDeleteErrors( objResult, tbStatus );

			return objResult.Count();
		}

		public static string ConvertToStandardValue( List<string> objStandardValues
			, string[] strWordsToReplace, string[] strSubstitutes, string strValueToConvert )
		{
			if( strValueToConvert == null ) return "";

			string strConvertedValue = "";

			// normalize value by replacing words in the value with one of the substitutes
			string strSource1 = string.Concat( " ", strValueToConvert.ToLower(), " " )
				.ReplaceKeywords( strWordsToReplace, strSubstitutes ).Trim();
			if( !strSource1.IsNullOrBlank() )
			{
				// compare to the standard values to select one
				foreach( string strStandardSpecialty in objStandardValues )
					if( strSource1.IsEqualOrPartiallyMatchedTo( strStandardSpecialty.ToLower() ) )
					{
						strConvertedValue = strStandardSpecialty;
						break;
					}
			}

			if( strConvertedValue.Equals( "" ) )
				strConvertedValue = "Other";

			return strConvertedValue;
		}

		public static void ReportErrorsToHistoryFile<T>( UpsertResult[] objResults, List<T> listObject = null ) where T : sObject
		{
			if( objResults != null )
			{
				StringBuilder strbText = GetErrorsStringBuilder( objResults, listObject );

				AppendTextToHistoryFile( strbText.ToString() );
			}
		}

		private static StringBuilder GetErrorsStringBuilder<T>( UpsertResult[] objResults, List<T> listObject = null ) where T:sObject
		{
			StringBuilder strbText = new StringBuilder( "\r\n", iTypicalRowSize * objResults.Count() );

			if( listObject != null )
			{
				System.Reflection.PropertyInfo objP = listObject[ 0 ].GetType().GetProperty( "Name" );
				if( objP == null )
					objP = listObject[ 0 ].GetType().GetProperty( "Title" );
				if( objP == null )
				{
					ConcatenateErrors( objResults, strbText );
					return strbText;
				}

				object objValue = null;
				string strValue = "                                                                                ";// 82 spaces
				for( int iIndex = 0; iIndex < objResults.Count(); iIndex++ )
				{
					UpsertResult objItem = objResults[ iIndex ];
					if( objItem.errors != null )
					{
						strbText.Append( "\r\n** Error:\t\t" );
						if( listObject.Count > iIndex )
						{
							strbText.Append( listObject[ iIndex ].Id );
							strbText.Append( "\t" );
							objValue = objP.GetValue( listObject[ iIndex ], null );
						}
						else
						{
							strbText.Append( "\t" );
							objValue = "";
						}
						strValue = ( objValue != null ) ? objValue.ToString() : "";
						strbText.Append( strValue );
						strbText.Append( "\t" );
						strbText.Append( objItem.errors[ 0 ].message );
						strbText.Append( " (" );
						strbText.Append( objItem.errors[ 0 ].statusCode );
						strbText.AppendLine( ")" );
					}
				}
			}
			else
				ConcatenateErrors( objResults, strbText );

			return strbText;
		}

		private static void ConcatenateErrors( UpsertResult[] objResults, StringBuilder strbText )
		{
			foreach( UpsertResult objItem in objResults )
			{
				if( objItem.errors != null )
				{
					strbText.Append( "\r\n** Error:\t\t" );
					strbText.Append( objItem.id );
					strbText.Append( "\t" );
					strbText.Append( objItem.errors[ 0 ].message );
					strbText.Append( " (" );
					strbText.Append( objItem.errors[ 0 ].statusCode );
					strbText.AppendLine( ")" );
				}
			}
		}

		public static void ReportErrors<T>( UpsertResult[] objResults, TextBox tbStatus, List<T> listObject = null ) where T : sObject
		{
			if( objResults != null )
			{
				StringBuilder strbText = GetErrorsStringBuilder( objResults, listObject );

				tbStatus.Text = string.Concat( tbStatus.Text, strbText.ToString() );
			}
		}

		public static void ReportErrors( SaveResult[] objResults, TextBox tbStatus )
		{
			if( objResults != null )
			{
				StringBuilder strbText = new StringBuilder( "\r\n", iTypicalRowSize * objResults.Count() );
				foreach( SaveResult objItem in objResults )
					if( objItem.errors != null )
					{
						strbText.Append( "\r\n" );
						strbText.Append( objItem.id );
						strbText.Append( "\t" );
						strbText.Append( objItem.errors[ 0 ].message );
						strbText.Append( " (" );
						strbText.Append( objItem.errors[ 0 ].statusCode );
						strbText.AppendLine( ")" );
					}

				tbStatus.Text = string.Concat( tbStatus.Text, strbText.ToString() );
			}

			//if( objResults != null )
			//{
			//    string strText = "\r\n";
			//    foreach( SaveResult objItem in objResults )
			//        if( objItem.errors != null )
			//            strText = string.Concat( strText, objItem.id, "\t"
			//                , objItem.errors[ 0 ].message, " (", objItem.errors[ 0 ].statusCode, ")\r\n" );

			//    tbStatus.Text = string.Concat( tbStatus.Text, strText );
			//}
		}

		public static void SetIdsReportErrors<T>( List<T> listObject, UpsertResult[] objResults, TextBox tbStatus ) where T : sObject
		{
			if( bMassDataLoad )
			{
				ReportErrors( objResults, tbStatus, listObject );
				return;
			}

			// preallocate 450 bytes per row (estimate taken from Results_Contact.csv:  25 Mb for 59k rows)
			StringBuilder strbText = new StringBuilder( "\r\n", iTypicalRowSize * listObject.Count );

			// estimate that there will be 17% of rows with errors
			StringBuilder strbErrors = new StringBuilder( "\r\n", ( iTypicalRowSize * listObject.Count ) / 6 );

			// create the header with column names
			if( listObject.Count > 0 )
			{
				strbText.Append( listObject[ 0 ].FieldNamesToTabString() );
				strbText.Append( "\tID/Error" );
			}

			if( objResults != null )
			{
				int iIndex = 0;
				foreach( T objItem in listObject )
				{
					// set the id in the original object for it may be used for lookups
					objItem.Id = objResults[ iIndex ].id;

					// add a line with the column values separated by tabs and append any error messages if existing
					strbText.Append( "\r\n" );
					string strValues = objItem.ToTabString();
					strbText.Append( strValues );

					// if there was an error in this row, append it to the list
					if( objResults[ iIndex ].errors != null )
					{
						strbErrors.Append( "\r\n" );
						strbErrors.Append( strValues );
						strbErrors.Append( "\t Error: " );
						strbErrors.Append( objResults[ iIndex ].errors[ 0 ].message );
						strbErrors.Append( " (" );
						strbErrors.Append( objResults[ iIndex ].errors[ 0 ].statusCode );
						strbErrors.Append( ")" );

						strbText.Append( "\t Error: " );
						strbText.Append( objResults[ iIndex ].errors[ 0 ].message );
						strbText.Append( " (" );
						strbText.Append( objResults[ iIndex ].errors[ 0 ].statusCode );
						strbText.Append( ")" );
					}

					iIndex++;
				}
			}
			else
				foreach( T objItem in listObject )
				{
					// add a line with the column values separated by tabs
					strbText.Append( "\r\n" );
					strbText.Append( objItem.ToTabString() );
				}

			// finish CSV content
			strbText.Append( "\r\n" );

			// create file name
			string strFileName = "Empty";
			if( listObject.Count > 0 )
				strFileName = listObject[ 0 ].GetType().ToString().Replace( "EmTrac2SF.EMSC2SF.", "" ).Replace( "__c", "" );
			strFileName = string.Concat( strPath, "Results_", strFileName, ".csv" );

			// display the tab separated list
			tbStatus.Text = string.Concat( tbStatus.Text, strbErrors.ToString() );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\nGenerated data results file:  ", strFileName );

			// save tab separated list in a file
			System.IO.File.WriteAllText( strFileName, strbText.ToString() );

			// create a link for the user to open the file in excel optionally
			if( OpenCSVList != null )
				OpenCSVList.InnerHtml = string.Concat( OpenCSVList.InnerHtml, "<br /><a href=\"DisplayCSV.aspx?file="
							, strFileName, "\" target=\"newwindow\">", strFileName, "</a>" );

			// hopefully this will release memory
			objResults = null;

			return;
		}

		public static void ReportDeleteErrors( DeleteResult[] objResults, TextBox tbStatus, bool bOnlyErrors = true )
		{
			string strText = "\r\n";
			int iIndex = 0;
			foreach( DeleteResult objItem in objResults )
			{
				if( bOnlyErrors && objItem.errors == null )
					continue;

				string strIDError = "OK";
				if( objItem.errors != null )
					strIDError = string.Concat( "\t Error: "
						, objItem.errors[ 0 ].message, " (", objItem.errors[ 0 ].statusCode, ")" );

				// add a line with the column values separated by tabs and append any error messages if existing
				strText = string.Concat( strText, "\r\n", objItem.id, "", strIDError );

				iIndex++;
			}

			// finish report and display the tab separated list
			tbStatus.Text = string.Concat( tbStatus.Text, strText, "\r\n" );

			return;
		}

		public static string CreateProviderScript( List<Contact> objList, bool bIncludeName = false )
		{
			// preallocate estimated size of script
			StringBuilder strbScript = new StringBuilder( "", 205 + 85 * objList.Count );

			// add CREATE temp table to the script
			strbScript.Append( "if OBJECT_ID( 'tempdb..#tmpProviders' ) > 0 drop table #tmpProviders;"
				+ "create table #tmpProviders (PhysicianNumber varchar(10), RecruitingID varchar(10)" );
			if( bIncludeName )
				strbScript.Append( ", [First Name] varchar(50), [Last Name] varchar(50)" );
			strbScript.Append( ");" );

			// avoid putting an if inside the loop
			if( bIncludeName )
				foreach( Contact objItem in objList )
				{
					strbScript.Append( "insert into #tmpProviders values (" );
					strbScript.Append( NullToZero( objItem.PhysicianNumber__c ) );
					strbScript.Append( "," );
					strbScript.Append( NullToZero( objItem.RecruitingID__c ) );
					strbScript.Append( ",'" );
					strbScript.Append( objItem.FirstName );
					strbScript.Append( "','" );
					strbScript.Append( objItem.LastName );
					strbScript.Append( "');" );
				}
			else
				foreach( Contact objItem in objList )
				{
					strbScript.Append( "insert into #tmpProviders values (" );
					strbScript.Append( NullToZero( objItem.PhysicianNumber__c ) );
					strbScript.Append( "," );
					strbScript.Append( NullToZero( objItem.RecruitingID__c ) );
					strbScript.Append( ");" );
				}

			return strbScript.ToString();
		}

		public static string NullToZero( double? dblValue )
		{
			if( dblValue == null || dblValue == 0 ) return "0";
			return dblValue.ToString();
		}

		public static string NullToZero( string strValue )
		{
			if( strValue.IsNullOrBlank() ) return "0";
			return strValue;
		}

		public static string ConvertToAgencyId( List<Credential_Agency__c> objAgencies, string strAgencyCode )
		{
			string strAgencyId = "";
			Credential_Agency__c objFoundAgency = objAgencies.FirstOrDefault( 
								i => i.EmTrac_Agency_Match__c.Equals( strAgencyCode ) );
			if( objFoundAgency != null )
				strAgencyId = objFoundAgency.Id;
			return strAgencyId;
		}

	}
}