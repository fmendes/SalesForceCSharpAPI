using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using System.Data;
using System.IO;
using System.Text;
using Metaphone;
using EmTrac2SF.EMSC2SF;

namespace GenericLibrary
{
	public static class Util
	{
		public static string FirstNonNull( params string[] strValues )
		{
			foreach( string strValue in strValues )
				if( strValue != null && ! strValue.Equals( "" ) )
					return strValue;

			return "";
		}

		public static string ToNormalizedMetaphone( this string strMain )
		{
			string strMainNbrs = Util.GetNumbersInString( strMain );
			string strMPMain = string.Concat( strMainNbrs, Util.NormalizeForMatching( strMain ) );

			return strMPMain;
		}

		public static string ToMetaphone( this string strMain )
		{
			string strMainNbrs = Util.GetNumbersInString( strMain );
			string strMPMain = string.Concat( strMainNbrs, Util.Metaphone( strMain ) );

			return strMPMain;
		}

		public static bool HasNumbers( this string strValue )
		{
			System.Text.RegularExpressions.Regex objRegex = new System.Text.RegularExpressions.Regex( @"\d+" );

			// detect numbers from the string to match
			System.Text.RegularExpressions.Match objMatch = objRegex.Match( strValue );
			return objMatch.Success;
		}

		public static bool HasLetters( this string strValue )
		{
			System.Text.RegularExpressions.Regex objRegex = new System.Text.RegularExpressions.Regex( @"[a-zA-Z]+" );

			// detect letters from the string to match
			System.Text.RegularExpressions.Match objMatch = objRegex.Match( strValue );
			return objMatch.Success;
		}

		public static string GetNumbersInString( string strValue )
		{
			if( strValue == null )
				return "";

			char[] buffer = new char[ strValue.Length ];
			int idx = 0;

			foreach( char c in strValue )
			{
				// this only collects digits and spaces
				if( c == ' ' || ( c >= '0' && c <= '9' ) )
				{
					buffer[ idx ] = c;
					idx++;
				}
			}

			return new string( buffer, 0, idx );

			//System.Text.RegularExpressions.Regex objRegex = new System.Text.RegularExpressions.Regex( @"\d+" );

			//// extract numbers from the column to match, if needed
			//System.Text.RegularExpressions.Match objMatch = objRegex.Match( strValue );
			//if( objMatch.Success )
			//    return objMatch.Value;

			//return "";
		}

		public static string TrimUpToSeparator( string strValue )
		{
			// get only what is before a dash/comma in a name
			int iPos = strValue.IndexOfAny( ",-/".ToCharArray() );
			if( iPos > 0 )
				return strValue.Substring( 0, iPos );

			return strValue;
		}

		public static string Metaphone( string strName )
		{
			// convert to sorted metaphone
			MultiWordMetaphone objMWM	= new MultiWordMetaphone();
			objMWM.Name = strName;

			return objMWM.MetaphoneKey;
		}

		//public static string[] strNonAlpha = { "-", ".", ",", "@", "$", "#", "(", ")", "/", "'" };
		public static char[] strNonAlpha = { '-', '.', ',', '@', '$', '#', '(', ')', '/', '\'' };

		// always use LOWERCASE
		public static string[] strSearchFor = { " & ", " + ", " saint ", " west ", " east ", " north ", " south ", " avenue "
								, " road ", " street ", " lane ", " po box ", " p o box ", " drive "
								, " parkway ", " boulevard ", " insurance ", " university ", " program ", " college "
								, " fort ", " school ", " medicine ", " som ", " center ", " cente ", " cnt "
								, " air force base ", " air force ", " hlth ", " sci ", " hospital ", " institute " };

		// always use LOWERCASE
		public static string[] strSubstitutions = { " and ", " and ", " st ", " w ", " e ", " n ", " s ", " ave " 
								, " rd ", " st ", " ln ", " pobox ", " pobox ", " dr "
								, " pkwy ", " blvd ", " ins ", " univ ", " prog ", " coll "
								, " ft ", " sch ", " med ", " school of medicine ", " ctr ", " ctr ", " ctr "
								, " usaf ", " usaf ", " health ", " sciences ", " hosp ", " inst" };

		// always use LOWERCASE
		public static string[] strWordsToRemove = { " company ", " co ", " corp ", " corporation ", " inc ", " incorporated "
					, " ltd ", " limited ", " plc ", " llc ", " corporation of america ", " the "
					, " dept of ", " department of ", " suite ", " group ", " grp ", " assoc ", " asso ", " association " };

		// words to keep separate (example:  match AtlantiCare with Atlanti-Care)
		public static string[] strWordsToSeparate = { "care", "health" };

		public static string NormalizeForMatching( string strName )
		{
			// prepare string for search/replace/removal
			string strKey = string.Concat( " ", strName.ToLower(), " " );

			// remove special characters
			strKey = strKey.ReplaceSpecialCharacters( "" );

			// remove non-alpha characters:  dashes, dots, commas
			strKey = strKey.ReplaceCharacters( strNonAlpha, ' ' );
			//strKey = strKey.ReplaceKeywords( strNonAlpha, "" );

			// replace & or + with "and"
			strKey = strKey.ReplaceKeywords( strSearchFor, strSubstitutions );

			// remove prefixes and suffixes
			strKey = strKey.ReplaceKeywords( strWordsToRemove, "" );

			// separate words
			strKey = strKey.SeparateKeywords( strWordsToSeparate );

			// remove unneeded spaces
			strKey = strKey.Trim().Replace( "  ", " " ).Replace( "  ", " " );

			// convert to sorted metaphone
			strKey = Metaphone( strKey );

			return strKey;
		}

		public static string[] strAMAInstitSearch = { " Sch Of Med ", " Coll Of Med ", " Inst "
													, " Med ", " Sch ", " Coll ", " Sci " }; //, " Univ "
		public static string[] strAMAInstitReplace = { " School of Medicine ", " College of Medicine ", " Institute "
													, " Medical ", " School ", " College ", " Sciences " }; //, " University "

		public static string NormalizeAMAInstitution( string strName )
		{
			strName = string.Concat( " ", strName, " " );
			strName = strName.ReplaceKeywords( strAMAInstitSearch, strAMAInstitReplace ).Trim();

			return strName;
		}

		public static string Capitalize( string strValue )
		{
			if( strValue == null ) return null;
			return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase( strValue.ToLower() );
		}

		public static string CapitalizeWithStateCode( string strValue )
		{
			if( strValue == null ) return null;

			strValue = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase( strValue.ToLower() );
			string strStatePattern = "[ ][A-Z][a-z]$";

			strValue = System.Text.RegularExpressions.Regex.Replace( strValue, strStatePattern
					, delegate( System.Text.RegularExpressions.Match match )
					{
						string strMatch = match.ToString();
						return strMatch.ToUpper();
					} );

			return strValue;
		}
	}

	public static class FileToObject
	{
		public static string ToSalesForceCSVString<T>( this T objT, string strColsToInclude = "" ) where T : sObject
		{
			// get individual column names from the list
			string[] strColumns = strColsToInclude.Split( ",".ToCharArray() );

			// store column info in the order received in the list
			List<PropertyInfo> objList = new List<PropertyInfo>( strColumns.Count() );
			foreach( string strCol in strColumns )
			{
				PropertyInfo objP = objT.GetType().GetProperty( strCol );
				if( objP != null )
					objList.Add( objP );
			}

			// concatenate values separated by comma
			StringBuilder strbResult = new StringBuilder( strColumns.Count() * 15 );			
			foreach( PropertyInfo objP in objList )
			{
				object objValue = objP.GetValue( objT, null );

				if( objValue != null )
				{
					string strColType = objP.PropertyType.ToString();
					switch( strColType )
					{
						case "System.String":
							// if value has commas, enclose it in quotes
							strbResult.Append( FixCSV( objValue.ToString() ) );
							break;
						case "System.DateTime": case "System.Nullable`1[System.DateTime]":
							strbResult.Append( FixCSV( (DateTime) objValue ) );
							break;
						default:
							strbResult.Append( objValue.ToString() );
							break;
					}
				}

				strbResult.Append( "," );
			}
			
			//objSW.AppendLine( strbLine.ToString().Substring( 1 ) );

			// remove last comma
			strbResult.Remove( strbResult.Length - 1, 1 );
			strbResult.AppendLine();

			return strbResult.ToString();
		}

		public static string ToSalesForceCSVString( this DataTable objDT )
		{
			StringBuilder objSW = new StringBuilder();

			// write column headers
			StringBuilder strbLine = new StringBuilder();
			foreach( DataColumn objDC in objDT.Columns )
			{
				strbLine.Append( "," );
				strbLine.Append( objDC.ColumnName );
			}
			objSW.AppendLine( strbLine.ToString().Substring( 1 ) );

			// write column values for each row
			foreach( DataRow objDR in objDT.Rows )
			{
				// write line with column values
				strbLine = new StringBuilder();
				for( int iColIndex = 0; iColIndex < objDT.Columns.Count; iColIndex++ )
				{
					strbLine.Append( "," );
					if( !Convert.IsDBNull( objDR[ iColIndex ] ) )
					{
						string strColType = objDR[ iColIndex ].GetType().ToString();
						switch( strColType )
						{
							case "System.String":
								strbLine.Append( FixCSV( objDR[ iColIndex ].ToString() ) );
								break;
							case "System.DateTime":
								strbLine.Append( FixCSV( (DateTime) objDR[ iColIndex ] ) );
								break;
							default:
								strbLine.Append( objDR[ iColIndex ].ToString() );
								break;
						}
					}
				}
				objSW.AppendLine( strbLine.ToString().Substring( 1 ) );
			}

			return objSW.ToString();
		}

		public static string SaveAsSalesForceCSV( this DataTable objDT, string strFileName )
		{
			StreamWriter objSW = new StreamWriter( strFileName, false );

			// write column headers
			StringBuilder strbLine = new StringBuilder(1100000); // max = 1.1 Kb per 2500 rows
			foreach( DataColumn objDC in objDT.Columns )
			{
				strbLine.Append( "," );
				strbLine.Append( objDC.ColumnName );
			}
			objSW.WriteLine( strbLine.ToString().Substring( 1 ) );

			// write column values for each row
			foreach( DataRow objDR in objDT.Rows )
			{
				// write line with column values
				strbLine = new StringBuilder();
				for( int iColIndex = 0; iColIndex < objDT.Columns.Count; iColIndex ++ )
				{
					strbLine.Append( "," );
					if( !Convert.IsDBNull( objDR[ iColIndex ] ) )
					{
						string strColType = objDR[ iColIndex ].GetType().ToString();
						switch( strColType )
						{
							case "System.String":
								strbLine.Append( FixCSV( objDR[ iColIndex ].ToString() ) );
								break;
							case "System.DateTime":
								strbLine.Append( FixCSV( (DateTime) objDR[ iColIndex ] ) );
								break;
							default:
								strbLine.Append( objDR[ iColIndex ].ToString() );
								break;
						}
					}
				}
				objSW.WriteLine( strbLine.ToString().Substring( 1 ) );
			}
			objSW.Close();

			return "";
		}

		public static string FixCSV( DateTime dtValue )
		{
			if( dtValue.Year > 2035 )
				return "2005-" + dtValue.Month.ToString("d2") + "-" + dtValue.Day.ToString("d2");

			return dtValue.ToString( "yyyy-MM-dd" );
		}

		public static string FixCSV( string strValue )
		{
			return "\"" + strValue.Replace("\"", "\"\"") + "\"";
		}

		/// <summary>
		/// Reads CSV file into a list of objects by placing each value in the respective object's members
		/// </summary>
		public static List<T> ReadFile<T>( this List<T> arr, string strFileName, string strColumnNames, bool bSkip1stRow = false )
		{
			List<T> obj = new List<T>();
			StreamReader objSR = new StreamReader( strFileName );

			string strLine;
			if( bSkip1stRow )
				strLine = objSR.ReadLine();

			while( ( strLine = objSR.ReadLine() ) != null )
				arr.Add( strLine.ConvertTo<T>( strColumnNames ) );

			objSR.Dispose();

			return arr;
		}

		/// <summary>
		/// Reads CSV file into a DataTable by placing each value in a column
		/// </summary>
		public static DataTable ReadFile( this DataTable objDT, string strFileName, string strColumnNames, bool b1stRowIsHeader = false )
		{
			StreamReader objSR = new StreamReader( strFileName );

			if( objDT == null ) objDT = new DataTable();

			string[] strSpecifiedCols = strColumnNames.Split( ',' );

			string strLine;
			if( b1stRowIsHeader )
			{
				// read 1st row/header to decide what names to give the columns
				strLine = objSR.ReadLine();
				string[] strCols = strLine.Split( ',' );
				int iIndex = 0;
				foreach( string strColumnName in strCols )
				{
					// only use column name from header if the column was not specified in the list
					string strName = strSpecifiedCols[ iIndex ];
					if( strName.Equals( "" ) )
					{
						// remove quotes since this is coming from a CSV file
						strName = strColumnName.Replace( "\"", "" );
						strSpecifiedCols[ iIndex ] = strName;
					}

					// create column with the name given, otherwise, the name from the header
					DataColumn objDC = new DataColumn( strName, typeof( string ) );
					objDT.Columns.Add( objDC );
					iIndex++;
				}
			}

			while( ( strLine = objSR.ReadLine() ) != null )
			{
				// place each value in the respective column
				string[] strValues = strLine.Split( ',' );
				int iIndex = 0;
				DataRow objDR = objDT.NewRow();
				foreach( string strColumnName in strSpecifiedCols )
				{
					string strValue = strValues[ iIndex ];

					// if the value starts with a quote but doesn't have an ending quote,
					// then the value may have been split by a comma
					if( strValue.StartsWith( "\"" ) && !strValue.EndsWith( "\"" ) )
						// keep appending to the value until finding an ending quote
						while( iIndex < strValues.Count() && !strValues[ iIndex ].EndsWith( "\"" ) )
						{
							iIndex++;
							strValue = string.Concat( strValue, ",", strValues[ iIndex ] );
						}

					//if( strColumnNames.Length <= iColIndex )
					//    break;
					//string strColName = strColumnNames[ iColIndex ];
					//if( strColName.Equals( "" ) )
					//{
					//    iColIndex++;
					//    continue;
					//}

					if( strValue.StartsWith( "\"" ) )
						strValue = strValue.Remove( 0, 1 );
					if( strValue.EndsWith( "\"" ) )
						strValue = strValue.Remove( strValue.Length - 1, 1 );

					// store column value without non-printable characters
					objDR[ strColumnName ] = strValue.ReplaceSpecialCharacters( "" );
					iIndex++;
				}

				// add new row to the table
				objDT.Rows.Add( objDR );
			}

			objSR.Dispose();

			return objDT;
		}
	}

	public static class AnonymousTypeConversion
	{
		/// <summary>
		/// Converts a single DataRow object into something else.
		/// The destination type must have a default constructor.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="objDR"></param>
		/// <returns></returns>
		public static T ConvertTo<T>( this DataRow objDR )
		{
			T item = Activator.CreateInstance<T>();
			for( int f = 0; f < objDR.Table.Columns.Count; f++ )
			{
				string strColName = objDR.Table.Columns[ f ].ColumnName;
				PropertyInfo p = item.GetType().GetProperty( strColName );

				if( p == null )
					continue;

				object objValue = objDR.Field<object>( strColName );

				// can't convert string to nullable datetime here
				// so check the target type p.PropertyType
				if( p.PropertyType == typeof( DateTime? ) )// && objType == typeof(string))
				{
					DateTime? dtValue = null;
					if( objValue != null && !objValue.ToString().Equals( "" ) )
						dtValue = Convert.ToDateTime( objValue );
					p.SetValue( item, dtValue, null );
				}
				else if( p.PropertyType == typeof( Boolean? ) )
				{
					Boolean? bValue = null;
					if( objValue != null && !objValue.ToString().Equals( "" ) )
					{
						if( objValue.ToString().Equals( "-1" ) )
							bValue = true;
						else if( objValue.ToString().Equals( "0" ) )
							bValue = false;
						else
							bValue = Convert.ToBoolean( objValue );
					}
					p.SetValue( item, bValue, null );
				}
				else if( p.PropertyType == typeof( Decimal? ) )
				{
					Decimal? decValue = null;
					if( objValue != null && !objValue.ToString().Equals( "" ) )
						decValue = Convert.ToDecimal( objValue );
					p.SetValue( item, decValue, null );
				}
				else if( p.PropertyType == typeof( Int32? ) || p.PropertyType == typeof( Int32 ) )
				{
					Int32? iValue = null;
					if( objValue != null && !objValue.ToString().Equals( "" ) )
						iValue = Convert.ToInt32( objValue );
					p.SetValue( item, iValue, null );
				}
				else if( p.PropertyType == typeof( Double? ) )
				{
					Double? dblValue = null;
					if( objValue != null && !objValue.ToString().Equals( "" ) )
						dblValue = Convert.ToDouble( objValue );
					p.SetValue( item, dblValue, null );
				}
				else if( p.PropertyType == typeof( string ) )
				{
					string strValue = null;
					if( objValue != null )
						strValue = objValue.ToString();
					p.SetValue( item, strValue, null );
				}
				else
					p.SetValue( item, objValue, null );

			}

			return item;
		}

		/// <summary>
		/// Converts a single DataRow object into something else.
		/// The destination type must have a default constructor.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="objDR"></param>
		/// <returns></returns>
		public static T ConvertTo<T>(this DataRow objDR, bool bFuzzyColumnMatch = false )
		{
			T item = Activator.CreateInstance<T>();
			for (int f = 0; f < objDR.Table.Columns.Count; f++)
			{
				string strColName = objDR.Table.Columns[f].ColumnName;
				Type objType = objDR.Table.Columns[f].DataType;
				PropertyInfo p = item.GetType().GetProperty(strColName);

				if (bFuzzyColumnMatch)
				{
					// if property name doesn't match, try removing spaces
					if (p == null)
					{
						string strNameUnderscore = strColName.Replace(" ", "");
						p = item.GetType().GetProperty(strNameUnderscore);
					}

					// if property name doesn't match, try replacing spaces with underscore
					if (p == null)
					{
						string strNameUnderscore = strColName.Replace(' ', '_');
						p = item.GetType().GetProperty(strNameUnderscore);
					}

					// if property name doesn't match, try lowercase minus initial
					if (p == null)
					{
						string strCapitalizedInitial = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(strColName);
						p = item.GetType().GetProperty(strCapitalizedInitial);
					}

					// if property name still doesn't match, try appending __c
					if (p == null)
					{
						string strCustomName = string.Concat(strColName, "__c");
						p = item.GetType().GetProperty(strCustomName);
					}

					// if property name doesn't match, try lowercase minus initial and appending __c
					if (p == null)
					{
						string strCustomName = string.Concat(strColName, "__c");
						string strCapitalizedInitial = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(strCustomName);
						p = item.GetType().GetProperty(strCapitalizedInitial);
					}

					if (p == null)
						if (strColName.Contains("Name"))
							p = item.GetType().GetProperty("Name");
				}

				if (p == null)
					continue;
				
				object objValue = objDR.Field<object>(strColName);

				// can't convert string to nullable datetime here
				// so check the target type p.PropertyType
				if (p.PropertyType == typeof(DateTime?) )// && objType == typeof(string))
				{
					DateTime? dtValue = null;
					if (objValue != null && !objValue.ToString().Equals(""))
						dtValue= Convert.ToDateTime(objValue);
					p.SetValue(item, dtValue, null);
				}
				else if( p.PropertyType == typeof(Boolean?) )
				{
					Boolean? bValue = null;
					if (objValue != null && !objValue.ToString().Equals(""))
					{
						if (objValue.ToString().Equals("-1"))
							bValue = true;
						else if (objValue.ToString().Equals("0"))
								bValue = false;
							else
								bValue = Convert.ToBoolean(objValue);
					}
					p.SetValue(item, bValue, null);
				}
				else if (p.PropertyType == typeof(Decimal?))
				{
					Decimal? decValue = null;
					if (objValue != null && !objValue.ToString().Equals(""))
						decValue = Convert.ToDecimal(objValue);
					p.SetValue(item, decValue, null);
				}
				else if (p.PropertyType == typeof(Int32?) || p.PropertyType == typeof(Int32))
				{
					Int32? iValue = null;
					if (objValue != null && !objValue.ToString().Equals(""))
						iValue = Convert.ToInt32(objValue);
					p.SetValue(item, iValue, null);
				}
				else if (p.PropertyType == typeof(Double?))
				{
					Double? dblValue = null;
					if (objValue != null && !objValue.ToString().Equals(""))
						dblValue = Convert.ToDouble(objValue);
					p.SetValue(item, dblValue, null);
				}
				else if (p.PropertyType == typeof(string))
				{
					string strValue = null;
					if (objValue != null)
						strValue = objValue.ToString();
					p.SetValue(item, strValue, null);
				}
				else
					p.SetValue(item, objValue, null);
				
			}

			return item;
		}

		/// <summary>
		/// Converts a string of CSV into an object placing each value into the object's members
		/// </summary>
		public static T ConvertTo<T>(this string strCSVRow, string strNames )
		{
			string[] strValues = strCSVRow.Split(',');
			string[] strColumnNames = strNames.Split(',');

			T item = Activator.CreateInstance<T>();
			int iColIndex = 0;
			for (int f = 0; f < strValues.Count(); f++)
			{
				// check whether the value is enclosed in quotes
				string strValue = strValues[ f ];

				// if the value starts with a quote but doesn't have an ending quote,
				// then the value may have been split by a comma
				if( strValue.StartsWith( "\"" ) && ! strValue.EndsWith( "\"" ) )
					// keep appending to the value until finding an ending quote
					while( f < strValues.Count() && ! strValues[ f ].EndsWith( "\"" ) )
					{
						f ++;
						strValue = string.Concat( strValue, ",", strValues[ f ] );
					}

				if( strColumnNames.Length <= iColIndex )
					break;
				string strColName = strColumnNames[ iColIndex ];
				if( strColName.Equals( "" ) )
				{
					iColIndex++;
					continue;
				}

				if( strValue.StartsWith( "\"" ) )
					strValue = strValue.Remove( 0, 1 );
				if( strValue.EndsWith( "\"" ) )
					strValue = strValue.Remove( strValue.Length - 1, 1 );

				object objValue = strValue;

				PropertyInfo p = item.GetType().GetProperty(strColName);

				// if property name doesn't match, try removing spaces
				if (p == null)
				{
					string strNameUnderscore = strColName.Replace(" ", "");
					p = item.GetType().GetProperty(strNameUnderscore);
				}

				// if property name doesn't match, try replacing spaces with underscore
				if (p == null)
				{
					string strNameUnderscore = strColName.Replace(' ', '_');
					p = item.GetType().GetProperty(strNameUnderscore);
				}

				// if property name doesn't match, try lowercase minus initial
				if (p == null)
				{
					string strCapitalizedInitial = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(strColName);
					p = item.GetType().GetProperty(strCapitalizedInitial);
				}

				// if property name still doesn't match, try appending __c
				if (p == null)
				{
					string strCustomName = string.Concat(strColName, "__c");
					p = item.GetType().GetProperty(strCustomName);
				}

				// if property name doesn't match, try lowercase minus initial and appending __c
				if (p == null)
				{
					string strCustomName = string.Concat(strColName, "__c");
					string strCapitalizedInitial = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(strCustomName);
					p = item.GetType().GetProperty(strCapitalizedInitial);
				}

				if (p == null)
					if (strColName.Contains("Name"))
						p = item.GetType().GetProperty("Name");

				if( p == null )
				{
					iColIndex++;
					continue;
				}

				// can't convert string to nullable datetime here
				// so check the target type p.PropertyType
				if (p.PropertyType == typeof(DateTime?))// && objType == typeof(string))
				{
					DateTime? dtValue = null;
					if (objValue.ToString() != "")
						dtValue = Convert.ToDateTime(objValue);
					p.SetValue(item, dtValue, null);
				}
				else if (p.PropertyType == typeof(Boolean?))
				{
					Boolean? bValue = null;
					if (objValue.ToString() != "")
						bValue = Convert.ToBoolean(objValue);
					p.SetValue(item, bValue, null);
				}
				else if (p.PropertyType == typeof(Decimal?))
				{
					Decimal? decValue = null;
					if (objValue.ToString() != "")
						decValue = Convert.ToDecimal(objValue);
					p.SetValue(item, decValue, null);
				}
				else if (p.PropertyType == typeof(Int32?) || p.PropertyType == typeof(Int32))
				{
					Int32? iValue = null;
					if (objValue.ToString() != "")
						iValue = Convert.ToInt32(objValue);
					p.SetValue(item, iValue, null);
				}
				else if (p.PropertyType == typeof(Double?))
				{
					Double? dblValue = null;
					if (objValue.ToString() != "")
						dblValue = Convert.ToDouble(objValue);
					p.SetValue(item, dblValue, null);
				}
				else
					p.SetValue(item, objValue.ToString(), null);

				iColIndex++;
			}

			return item;
		}

		/// <summary>
		/// Converts a list of DataRow to a list of something else.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="list"></param>
		/// <returns></returns>
		public static List<T> ConvertTo<T>(this List<DataRow> list)
		{
			List<T> result = (List<T>)Activator.CreateInstance<List<T>>();

			list.ForEach(rec =>
			{
				result.Add(rec.ConvertTo<T>());
			});

			return result;
		}
	}

	public static class StringExtensions
	{
		/// <summary>
		/// Returns TRUE if both are equal or if one contains the other, returns FALSE for nulls/blanks
		/// </summary>
		public static bool IsEqualOrPartiallyMatchedTo( this string strMain, string strCompareTo )
		{
			if( strCompareTo == null ) return false;
			if( strMain == null ) return false;

			string strMainLower = strMain.ToLower();
			string strCompareToLower = strCompareTo.ToLower();

			if( strMainLower.Equals( "" ) || strCompareToLower.Equals( "" ) )
				return false;

			if( strMainLower.Equals( strCompareToLower )
										|| strMainLower.Contains( strCompareToLower )
										|| strCompareToLower.Contains( strMainLower ) )
				return true;

			return false;
		}

		public static string Left( this string strMain, int iLimit )
		{
			if( strMain.Length > iLimit )
				return strMain.Substring( 0, iLimit );
			return strMain;
		}

		public static bool IsMetaphoneMatchedTo( this string strMain, string strCompareTo )
		{
			if( strCompareTo == null ) return false;
			if( strMain == null ) return false;

			string strMainNbrs = Util.GetNumbersInString( strMain );
			string strCompareToNbrs = Util.GetNumbersInString( strCompareTo );

			//string strMPMain = string.Concat( strMainNbrs, Util.Metaphone( strMain ) );
			//string strMPCompareTo = string.Concat( strCompareToNbrs, Util.Metaphone( strCompareTo ) );

			//// attempt to match metaphone
			//if( strMPMain.Equals( strMPCompareTo ) )
			//    return true;
			// if they didn't match, try again after more treatment

			string strCompareToKey = string.Concat( strMainNbrs, Util.NormalizeForMatching( strCompareTo ) );
			string strMainKey = string.Concat( strCompareToNbrs, Util.NormalizeForMatching( strMain ) );
			if( strMainKey.Equals( strCompareToKey ) )
				return true;

			return false;
		}

		/// <summary>
		/// Returns Equals if both are not null/blank, if either or both are null returns TRUE. It is a "relaxed" Equals
		/// </summary>
		public static bool NotNullBlankAndEquals( this string strMain, string strCompareTo )
		{
			if( strMain == null || strCompareTo == null ) return false;
			if( strMain.Trim().Equals( "" ) || strCompareTo.Trim().Equals( "" ) ) return false;
			return strMain.Equals( strCompareTo );
		}

		/// <summary>
		/// Returns Equals if both datetimes are not null, if either or both are null returns TRUE. It is a "relaxed" Equals
		/// </summary>
		public static bool NotNullAndEquals( this DateTime? dtMain, DateTime? dtCompareTo )
		{
			if( dtMain == null && dtCompareTo == null ) return false;
			if( dtMain != null && dtCompareTo == null ) return false;
			if( dtMain == null && dtCompareTo != null ) return false;
			DateTime dt1 = (DateTime) dtMain, dt2 = (DateTime) dtCompareTo;
			return dt1.CompareTo( dt2 ) == 0;
		}

		/// <summary>
		/// Returns Equals if both are not null, if either or both are null returns TRUE. It is a "relaxed" Equals
		/// </summary>
		public static bool NotNullAndEquals<T>( this T strMain, T strCompareTo )
		{
			if( strMain == null && strCompareTo == null ) return false;
			if( strMain != null && strCompareTo == null ) return false;
			if( strMain == null && strCompareTo != null ) return false;
			return strMain.Equals( strCompareTo );
		}

		/// <summary>
		/// Returns Equals if both are not null, returns TRUE if both are null, otherwise returns FALSE
		/// </summary>
		public static bool NullAwareEquals( this string strMain, string strCompareTo )
		{
			//if( strMain.Trim().Equals( "" ) || strCompareTo.Trim().Equals( "" ) ) return true;
			if( strMain == null && strCompareTo == null ) return true;
			if( strMain != null && strCompareTo == null ) return false;
			if( strMain == null && strCompareTo != null ) return false;
			return strMain.Equals( strCompareTo );
		}

		/// <summary>
		/// Returns Equals if both are not null, returns TRUE if both are null, otherwise returns FALSE
		/// </summary>
		public static bool NullAwareEquals<T>( this T strMain, T strCompareTo )
		{
			if( strMain == null && strCompareTo == null ) return true;
			if( strMain != null && strCompareTo == null ) return false;
			if( strMain == null && strCompareTo != null ) return false;
			return strMain.Equals( strCompareTo );
		}

		public static bool ContainsAnyPartOf( this string strList2Exclude, string strValue )
		{
			string[] strList = strList2Exclude != null ? strList2Exclude.Split( new char[] { '|' } ) : new string[] {};
			foreach( string strExclude in strList )
			{
				if( strValue.Contains( strExclude ) )
					return true;
			}

			return false;
		}

		public static string ParseFromTo( this string strValue, string strFrom, string strTo )
		{
			int iPos = strValue.IndexOf( strFrom ) + strFrom.Length;
			int iPosNext = strValue.IndexOf( strTo );

			string strResult = strValue.Substring( iPos, iPosNext - iPos );

			return strResult;
		}

	}

	public static class DataExtensions
	{
		public static string COLS_TO_EXCLUDE = "Account|Attachments|CreatedBy|CreatedDate|FeedSubscriptionsForEntity|IsDeleted|IsDeletedSpecified|LastModifiedBy|LastModifiedDate|LastModifiedDateSpecified|Notes|NotesAndAttachments|ProcessInstances|ProcessSteps|SystemModstamp|SystemModstampSpecified|fieldsToNull|__cSpecified";
		/// <summary>
		/// Checks if object is a blank or null string (more intelligible/faster this way)
		/// </summary>
		public static bool IsNullOrBlank( this string obj )
		{
			return ( obj == null || obj.Trim().Length == 0 );
		}

		/// <summary>
		/// Checks if object is a blank or null string (more intelligible/faster this way)
		/// </summary>
		public static bool IsNullOrBlank( this object obj )
		{
			return ( obj == null || obj.ToString().Length == 0 );
		}

		/// <summary>
		/// Checks if object is a blank or null string (more intelligible/faster this way)
		/// </summary>
		public static bool IsNullOrBlank<T>( this T obj )
		{
			return ( obj == null || obj.ToString().Length == 0 );
		}

		/// <summary>
		/// Creates a tab delimited string with the values of the DataRow
		/// </summary>
		public static string ToTabString( this DataRow objDR ) 
		{
			StringBuilder strbResult = new StringBuilder( objDR.Table.Columns.Count * 15 );
			foreach( DataColumn objDC in objDR.Table.Columns )
			{
				strbResult.Append( "\t" );
				strbResult.Append( ( objDR[ objDC ] != null ) ? objDR[ objDC ].ToString() : "" );
			}

			// remove 1st tab
			strbResult.Remove( 0, 1 );

			return strbResult.ToString();
		}

		/// <summary>
		/// Creates a tab delimited string with a list of the DataTable column names
		/// </summary>
		public static string ColumnsToTabString(this DataTable objDT)
		{
			StringBuilder strbResult = new StringBuilder( objDT.Columns.Count * 15 ); 
			
			foreach( DataColumn objDC in objDT.Columns )
			{
				strbResult.Append( "\t" );
				strbResult.Append( objDC.ColumnName );
			}

			// remove 1st tab
			strbResult.Remove( 0, 1 );

			return strbResult.ToString();
		}

		/// <summary>
		/// Creates a JSON string with a list of the object's members values
		/// </summary>
		public static string ToJSON<T>( this T objT, bool bIncludeAllColumns = false )
		{
			PropertyInfo[] p = objT.GetType().GetProperties();

			StringBuilder strbResult = new StringBuilder( p.Count() * 15 );

			// exclude columns if needed
			List<PropertyInfo> objList = p.ToList();
			if( !bIncludeAllColumns )
				objList = objList.Where( pi => !COLS_TO_EXCLUDE.ContainsAnyPartOf( pi.Name ) ).ToList();

			strbResult.Append( "{\"" );

			foreach( PropertyInfo objP in objList )
			{
				object objValue = objP.GetValue( objT, null );

				strbResult.Append( objP.Name );

				// if value has commas, enclose it in quotes
				if( objValue == null )
				{
					strbResult.Append( "\":null,\"" );
				}
				else
				{
					strbResult.Append( "\":\"" );
					strbResult.Append( objValue.ToString() );
					strbResult.Append( "\",\"" );
				}

			}

			// remove last comma and double-quote at the end
			strbResult.Remove( strbResult.Length - 2, 2 );
			strbResult.Append( "}" );

			return strbResult.ToString();
		}

		/// <summary>
		/// Creates a JSON string with a list of the object's members values
		/// </summary>
		public static string ToJSON<T>( this T objT, string strColumnList = null )
		{
			PropertyInfo[] p = objT.GetType().GetProperties();

			StringBuilder strbResult = new StringBuilder( p.Count() * 15 );

			// only include specified columns
			List<PropertyInfo> objList = p.ToList();
			if( strColumnList != null )
			{
				strColumnList = "," + strColumnList.ToLower() + ",";
				objList = objList.Where( pi => strColumnList.Contains( "," + pi.Name.ToLower() + "," ) ).ToList();
			}

			strbResult.Append( "{\"" );

			foreach( PropertyInfo objP in objList )
			{
				object objValue = objP.GetValue( objT, null );

				strbResult.Append( objP.Name );

				// if value has commas, enclose it in quotes
				if( objValue == null )
				{
					strbResult.Append( "\":null,\"" );
				}
				else
				{
					strbResult.Append( "\":\"" );
					strbResult.Append( objValue.ToString() );
					strbResult.Append( "\",\"" );
				}

			}

			// remove last comma and double-quote at the end
			strbResult.Remove( strbResult.Length - 2, 2 );
			strbResult.Append( "}" );

			return strbResult.ToString();
		}

		/// <summary>
		/// Creates a tab delimited string with a list of the object's members values
		/// </summary>
		public static string ToTabString<T>(this T objT, bool bIncludeAllColumns = false )
		{
			PropertyInfo[] p = objT.GetType().GetProperties();

			StringBuilder strbResult = new StringBuilder( p.Count() * 15 );

			// exclude columns if needed
			List<PropertyInfo> objList = p.ToList();
			if( ! bIncludeAllColumns )
				objList = objList.Where( pi => !COLS_TO_EXCLUDE.ContainsAnyPartOf( pi.Name ) ).ToList();

			// reorder the list
			objList.Sort( ( p1, p2 ) => CompareColumnNames( p1, p2 ) );

			foreach( PropertyInfo objP in objList )
			{
				object objValue = objP.GetValue(objT, null);
				string strValue = ( objValue != null ) ? objValue.ToString() : "";

				// remove interfering characters
				//strValue = strValue.Replace( "\n", "|" ).Replace( "\r", "|" ).Replace( "\t", " " )
				//                    .Replace( "\"", "'" );
				strValue = strValue.ReplaceSpecialCharacters( ' ' );

				strbResult.Append( "\t" );

				// if value has commas, enclose it in quotes
				if( strValue.Equals( "" ) || !strValue.Contains( "," ) )
					strbResult.Append( strValue );
				else
				{
					strbResult.Append( "'" );
					strbResult.Append( strValue );
					strbResult.Append( "'" );
				}
			}

			// remove first tab and return
			strbResult.Remove( 0, 1 );
			return strbResult.ToString();
		}

		public static string RemoveSpecialCharacters( this string str, bool bReplaceWithSpace = false
			, bool bSkipNLCR = false )
		{
			char[] buffer = new char[ str.Length ];
			int idx = 0;

			foreach( char c in str )
			{
				//if( ( c >= '0' && c <= '9' ) || ( c >= 'A' && c <= 'Z' )
				//	|| ( c >= 'a' && c <= 'z' ) || ( c == '.' ) || ( c == '_' ) || ( c == '@' ) )
				//if( ( c >= ' ' && c <= 'Z' ) || ( c >= 'a' && c <= 'z' ) )
				char cNew = c;
				if( ( cNew >= ' ' && cNew <= '~' ) )
				{
					// if valid character, copy it, otherwise skip it
					buffer[ idx ] = cNew;
					idx++;
				}
				else
					if( !bSkipNLCR && ( cNew == 10 || cNew == 13 ) )
					{
						// copy New Line and/or Carriage Return
						buffer[ idx ] = cNew;
						idx++;
					}
					else
					if( bReplaceWithSpace )
					{
						// add a space instead of the special character
						buffer[ idx ] = ' ';
						idx++;
					}
			}

			return new string( buffer, 0, idx );
		}

		public static string ReplaceKeywords( this string str, string[] strList, string strSubstitute )
		{
			foreach( string strKey in strList )
				str = str.Replace( strKey, strSubstitute );

			return str;
		}

		public static string ReplaceCharacters( this string str, char[] strList, char cSubstitute)
		{			
			char[] buffer = new char[ str.Length ];
			int idx = 0;

			foreach( char c in str )
			{
				if( strList.Contains( c ) )
					// if valid character, copy it, otherwise skip it
					buffer[ idx ] = cSubstitute;
				else
					buffer[ idx ] = c;
				idx++;
			}

			return new string( buffer, 0, idx );
		}

		public static string SeparateKeywords( this string str, string[] strList )
		{
			// surround each keyword with spaces, then remove duplicate spaces
			foreach( string strKey in strList )
				str = str.Replace( strKey, string.Concat( " ", strKey, " " ) ).Replace( "  ", " " );

			return str;
		}

		public static string ReplaceKeywords( this string str, string[] strList, string[] strSubstitute )
		{
			int iIndex = 0;
			foreach( string strKey in strList )
			{
				str = str.Replace( strKey, strSubstitute[ iIndex ] );
				iIndex ++;
			}

			return str;
		}

		public static string ReplaceSpecialCharacters( this string str, string strSubstitute )
		{
			if( strSubstitute.Length != 0 )
				return str.ReplaceSpecialCharacters( strSubstitute[0] );

			char[] buffer = new char[ str.Length ];
			int idx = 0;

			foreach( char c in str )
			{
				// this only excludes special characters (non-printable)
				if( ( c >= ' ' && c <= '~' ) )
				{
					buffer[ idx ] = c;
					idx++;
				}
			}

			return new string( buffer, 0, idx );
		}

		public static string ReplaceSpecialCharacters( this string str, char cSubstitute )
		{
			char[] buffer = new char[ str.Length ];
			int idx = 0;

			foreach( char c in str )
			{
				// this only excludes special characters (non-printable)
				if( ( c >= ' ' && c <= '~' ) )
					buffer[ idx ] = c;
				else
					buffer[ idx ] = cSubstitute;

				idx++;
			}

			return new string( buffer, 0, idx );
		}

		/// <summary>
		/// Comparison function to place column names with __c, then Id, Name, Error last in the sort order
		/// </summary>
		private static int CompareColumnNames( PropertyInfo p1, PropertyInfo p2 )
		{
			if( p1.Name == "Name" || p1.Name == "Id" || p1.Name == "ID/Error" )
				if( p2.Name == "Name" || p2.Name == "Id" || p2.Name == "ID/Error" )
					return -p1.Name.CompareTo( p2.Name );
				else
					return -1;	// Id/Name/Error come first
			if( p2.Name == "Name" || p2.Name == "Id" || p2.Name == "ID/Error" )
				return 1;	// Id/Name/Error are always later

			if( p1.Name.Contains( "__c" ) )
				if( !p2.Name.Contains( "__c" ) )
					return -1;	// fields with __c come first

			if( p2.Name.Contains( "__c" ) )
				if( !p1.Name.Contains( "__c" ) )
					return 1;	// fields with __c, are always later

			return - p1.Name.CompareTo( p2.Name );	// invert order to make it consistent
		}

		/// <summary>
		/// Creates tab delimited string with a list of object's members names
		/// </summary>
		public static string FieldNamesToTabString<T>( this T objT, bool bIncludeAllColumns = false )
		{
			PropertyInfo[] p = objT.GetType().GetProperties();
			StringBuilder strbResult = new StringBuilder( p.Count() * 15 ); 

			List<PropertyInfo> objList = p.ToList();
			if( ! bIncludeAllColumns )
				objList = objList.Where( pi => !COLS_TO_EXCLUDE.ContainsAnyPartOf( pi.Name ) ).ToList();

			// reorder the list
			objList.Sort( ( p1, p2 ) => CompareColumnNames( p1, p2 ) );

			foreach( PropertyInfo objP in objList )
			{
				strbResult.Append( "\t" );
				strbResult.Append( objP.Name );
			}

			// remove 1st tab
			strbResult.Remove( 0, 1 );

			return strbResult.ToString();
		}
	}

	public static class ArrayExtensions
	{
		/// <summary>
		/// Extracts a subarray of a given number of items (count) from an array starting from a given position (startIndex)
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="arr">Array from which to extract a subarray</param>
		/// <param name="startIndex">Position from where to obtain the array items</param>
		/// <param name="count">Number of items</param>
		/// <returns></returns>
		public static T[] SubArray<T>(this T[] arr, int startIndex, int count)
		{
			var sub = new T[arr.Length];
			int iActualCount = count;
			if( arr.Count() - startIndex < count )
				iActualCount = arr.Count() - startIndex;
			Array.Copy(arr, startIndex, sub, 0, iActualCount);
			return sub;
		}
	}
}