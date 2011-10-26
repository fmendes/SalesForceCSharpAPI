using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using EmTrac2SF.EMSC2SF;
using GenericLibrary;
using System.Net;
using System.Web.Script.Serialization;
using System.IO;
using System.Xml;

namespace EmTrac2SF.Salesforce
{
	public class ApiService : IDisposable
	{
		public string ConnectionString = "SalesforceLogin";
		public static Dictionary<Guid, List<sObject>> asyncResults;

		private SforceService salesforceService;
		const int defaultTimeout = 30000;
		public int Retries = 1;

		public string SessionID = "";
		public string ServerURL = "";

		public bool ConnectionFailed = false;
		public string ErrorMessage = "";
		public delegate bool OnErrorHandler( string strErrorMessage, string strCommand );
		public OnErrorHandler OnError;
		public delegate bool OnReportStatusHandler( params string[] strStatus );
		public OnReportStatusHandler OnReportStatus;

		public ApiService()
		{
			salesforceService = new SforceService();
			salesforceService.Timeout = defaultTimeout;
			asyncResults = new Dictionary<Guid, List<sObject>>();
		}

		public ApiService(int timeout) : this()
		{
			salesforceService.Timeout = timeout;
		}

		public bool ErrorHandler( string strCommand )
		{
			bool bCancel = false;
			
			if( OnError != null )
				bCancel = OnError( ErrorMessage, strCommand );

			ReportStatus( "Error executing ", strCommand, ": ", ErrorMessage );

			return bCancel;
		}

		public bool ReportStatus( params string[] strStatus )
		{
			bool bCancel = false;

			if( OnReportStatus != null )
				bCancel = OnReportStatus( strStatus );

			return bCancel;
		}

		public string RESTCreateJob( string strOperation, string strObjectName
			, string strExternalID, string strFileType )
		{
			string strBody = string.Concat( "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
,"<jobInfo xmlns=\"http://www.force.com/2009/06/asyncapi/dataload\"><operation>", strOperation
, "</operation><object>", strObjectName, "</object>" );

			if( strOperation.Equals( "upsert" ) )
				strBody = string.Concat( strBody, "<externalIdFieldName>", strExternalID, "</externalIdFieldName>" );

			strBody = string.Concat( strBody, "<contentType>", strFileType, "</contentType></jobInfo>" );

			string strError;
			string strResponse = GetRESTOperationResult( "/services/async/22.0/job", strBody, out strError
								, false, null, "application/xml" );

			if( strResponse.Equals( "" ) )
				return strError;

			string strJobId = "";
			// get Id of the new job and return it
			XmlReader objReadXML = XmlReader.Create( new StringReader( strResponse ) );
			if( strError.Equals( "" ) )
				strJobId = objReadXML.ReadToFollowing( "id" ) ? objReadXML.ReadElementContentAsString() : "";
			else
				strJobId = "ERROR:  " + ( objReadXML.ReadToFollowing( "exceptionMessage" ) ? 
					objReadXML.ReadElementContentAsString() : "" );

			return strJobId;
		}

		public string RESTSetJobState( string strJobId, string strState = "Closed" )
		{
			string strURL = string.Concat( "/services/async/22.0/job/", strJobId );

			// XML request to close job
			string strBody = string.Concat( "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
, "<jobInfo xmlns=\"http://www.force.com/2009/06/asyncapi/dataload\"><state>", strState, "</state></jobInfo>" );

			string strError;
			string strResponse = GetRESTOperationResult( strURL, strBody, out strError
								, false, null, "application/xml" );

			if( strResponse.Equals( "" ) )
				return strError;

			// get Id of the new batch and return it
			XmlReader objReadXML = XmlReader.Create( new StringReader( strResponse ) );
			
			string strJobState = objReadXML.ReadToFollowing( "state" ) ? objReadXML.ReadElementContentAsString() : "";

			return strJobState;
		}

		public string RESTCreateBatch( string strJobId, string strCSVContent )
		{
			string strURL = string.Concat( "/services/async/22.0/job/", strJobId, "/batch" );

			string strError;
			string strResponse = GetRESTOperationResult( strURL, strCSVContent, out strError
								, false, null, "text/csv" );

			if( strResponse.Equals( "" ) )
				return strError;

			// get Id of the new batch and return it
			XmlReader objReadXML = XmlReader.Create( new StringReader( strResponse ) );
			
			string strBatchId = objReadXML.ReadToFollowing( "id" ) ? objReadXML.ReadElementContentAsString() : "";

			return strBatchId;
		}

		public string RESTCheckBatch( string strJobId, string strBatchId )
		{
			string strURL = string.Concat( "/services/async/22.0/job/", strJobId, "/batch/", strBatchId );

			string strError;
			string strResponse = GetRESTOperationResult( strURL, "", out strError
								, false, null, "application/xml" );

			if( strResponse.Equals( "" ) )
				return strError;

			// get Id of the new batch and return it
			XmlReader objReadXML = XmlReader.Create( new StringReader( strResponse ) );
			string strState = objReadXML.ReadToFollowing( "state" ) ? 
								objReadXML.ReadElementContentAsString() : "";

			string strStateMessage = objReadXML.ReadToFollowing( "stateMessage" ) ? 
									objReadXML.ReadElementContentAsString() : "";
			
			string strNbrRecords = objReadXML.ReadToFollowing( "numberRecordsProcessed" ) ? 
									objReadXML.ReadElementContentAsString() : "";

			return strState + "-" + strStateMessage + "-" + strNbrRecords;
		}

		public string RESTGetBatchResult( string strJobId, string strBatchId )
		{
			string strURL = string.Concat( "/services/async/22.0/job/", strJobId, "/batch/", strBatchId, "/result" );

			string strError;
			string strResponse = GetRESTOperationResult( strURL, "", out strError
								, false, null, "text/csv" );

			if( strResponse.Equals( "" ) )
				return strError;

			return strResponse;
		}

		public SaveResult RESTCreateRecord( sObject obj, string strMemberList = null )
		{
			// transform sObject into JSON string
			string strBody = "";
			if( strMemberList == null )
				// by default, exclude columns from the default exclusion list
				strBody = obj.ToJSON( false );	 // was "Name,Id" 
			else
				strBody = obj.ToJSON( strMemberList );

			string strError;
			string strResponse = 
					GetRESTOperationResult( "/services/data/v20.0/sobjects/", strBody, out strError, false, obj );

			SaveResult objSResult = ConvertRESTResponseToSaveResult( strError, strResponse );

			return objSResult;
		}

		public static SaveResult ConvertRESTResponseToSaveResult( string strError, string strResponse )
		{
			// transform string into object
			JavaScriptSerializer objSerializer = new JavaScriptSerializer();
			SaveResult objSResult = null;
			if( ! strError.Equals( "" ) )
			{
				objSResult = new SaveResult();
				if( strResponse.Equals( "" ) )
				{
					objSResult.errors = ( new List<Error>( 1 ) ).ToArray();
					objSResult.errors[ 0 ].message = strError;
				}
				else
				{
					List<Error> objErrors = objSerializer.Deserialize<List<Error>>( strResponse );
					objSResult.errors = objErrors.ToArray();
				}
			}
			else
				objSResult = objSerializer.Deserialize<SaveResult>( strResponse );
			return objSResult;
		}

		public static string GetRESTResponse( WebRequest objWR, out string strError )
		{
			string strResponse = "";
			strError = "";
			WebResponse objWResp = null;
			try
			{
				objWResp = objWR.GetResponse();
			}
			catch( WebException excpt )
			{
				strError = excpt.Message;
				objWResp = excpt.Response;
			}

			if( objWResp == null )			
				return "";

			StreamReader objSR = new StreamReader( objWResp.GetResponseStream() );
			strResponse = objSR.ReadToEnd();
			objSR.Close();

			objWResp.Close();

			return strResponse;
		}

		public static void SendRESTRequest( WebRequest objWR, string strBody )
		{
			byte[] objRequestBody = System.Text.Encoding.ASCII.GetBytes( strBody );
			objWR.ContentLength = objRequestBody.Length;

			var objOutput = objWR.GetRequestStream();
			objOutput.Write( objRequestBody, 0, objRequestBody.Length );
			objOutput.Close();
		}

		public WebRequest CreateRESTRequest( string strServiceURLPart, sObject obj = null
					, string strContentType = "application/json" )
		{
			// find where the path starts in the url to truncate the path
			int iPos = ServerURL.IndexOf( "/service" );
			string strURL = ServerURL.Substring( 0, iPos );

			// append object name if specified
			string strObjectName = "";
			if( obj != null )
				strObjectName = obj.GetType().Name + "/";

			// concatenate the REST url and create the request
			strURL = string.Concat( strURL, strServiceURLPart, strObjectName );
			WebRequest objWR = WebRequest.Create( strURL );

			// configure headers
			objWR.Headers.Add( "Authorization: OAuth " + SessionID );
			objWR.ContentType = strContentType;
			objWR.Method = "POST";

			return objWR;
		}

		public string GetRESTOperationResult( string strServiceURLPart, string strBody, out string strError
						, bool bOAuth = true
						, sObject obj = null, string strContentType = "application/json" )
		{
			strError = "";
			if( !SetupService() ) return "";

			// find where the path starts in the url to truncate the path
			int iPos = ServerURL.IndexOf( "/service" );
			string strURL = ServerURL.Substring( 0, iPos );

			// append object name if specified
			string strObjectName = "";
			if( obj != null )
				strObjectName = obj.GetType().Name + "/";

			// concatenate the REST url and create the request
			strURL = string.Concat( strURL, strServiceURLPart, strObjectName );
			HttpWebRequest objWR = (HttpWebRequest) WebRequest.Create( strURL );

			// configure web request
			if( bOAuth )
				objWR.Headers.Add( "Authorization: OAuth " + SessionID );
			else
				objWR.Headers.Add( "X-SFDC-Session: " + SessionID );

			if( ! strContentType.Equals( "" ) )
				objWR.ContentType = strContentType;

			objWR.Method = strBody.Equals( "" ) ? "GET" : "POST";
			objWR.Timeout = 10000;
			objWR.ReadWriteTimeout = 20000;
			//objWR.KeepAlive = false;
			objWR.ProtocolVersion = HttpVersion.Version10;

			if( ! strBody.Equals( "" ) )
			{
				byte[] objRequestBody = System.Text.Encoding.ASCII.GetBytes( strBody );
				objWR.ContentLength = objRequestBody.Length;

				Stream objOutput = objWR.GetRequestStream();
				objOutput.Write( objRequestBody, 0, objRequestBody.Length );
				objOutput.Close();
			}

			string strResponse = "";
			WebResponse objWResp = null;
			try
			{
				objWResp = objWR.GetResponse();
			}
			catch( WebException excpt )
			{
				strError = excpt.Message;
				objWResp = excpt.Response;
			}

			if( objWResp == null )
				return "";

			StreamReader objSR = new StreamReader( objWResp.GetResponseStream() );
			strResponse = objSR.ReadToEnd();
			objSR.Close();

			objWResp.Close();

			return strResponse;
		}

		public string[] QueryIDs<T>( string soql ) where T : sObject, new()
		{
			List<T> returnList = new List<T>();

			if( !SetupService() ) { return null; }

			int iIndex = 0;
			QueryResult results = salesforceService.query( soql );
			while( iIndex < results.size )
			{
				for( int i = 0; i < results.records.Count(); i++ )
				{
					T item = results.records[ i ] as T;

					if( item != null )
						returnList.Add( item );
					else
						iIndex++; // this is to avoid infinite loop
				}

				iIndex += results.records.Count();

				// read more records if needed
				if( iIndex < results.size && !results.done )
					results = salesforceService.queryMore( results.queryLocator );

			}	// end while( iIndex < results.size )

			return returnList.ConvertAll( t => t.Id ).ToArray();
		}

		public List<T> Query<T>(string soql) where T : sObject, new()
		{
			List<T> returnList = new List<T>();

			T obj = new T();
			string strObjectName = obj.GetType().Name;

			if( !SetupService() ) { return null; }

			int iIndex = 0;
			QueryResult results = null;

			// attempt a certain number of times
			int iFailCount = 0;
			while( iFailCount < Retries )
			{
				ErrorMessage = "";

				try
				{
					results = salesforceService.query( soql );
					iFailCount = Retries;

					ReportStatus( "Querying ", strObjectName, ":  ", soql, " / ", iIndex.ToString(), " rows retrieved." );
				}
				catch( System.Web.Services.Protocols.SoapException excpt )
				{
					// catch error without interruption
					ErrorMessage = excpt.Message;
					iFailCount++;
					if( ErrorHandler( soql ) )
						return returnList;
				}
			}

			while( results != null && iIndex < results.size )
			{
				for (int i = 0; i < results.records.Count(); i++)
				{
					T item = results.records[i] as T;

					if (item != null)
						returnList.Add(item);
					else
						iIndex++; // this is to avoid infinite loop
				}

				iIndex += results.records.Count();

				// read more records if needed
				if( iIndex < results.size && !results.done )
				{
					// attempt a certain number of times after each failure
					iFailCount = 0;
					while( iFailCount < Retries )
					{
						try
						{
							results = salesforceService.queryMore( results.queryLocator );
							iFailCount = Retries;

							ReportStatus( "Querying more ", strObjectName, ":  ", iIndex.ToString(), " rows retrieved." );
						}
						catch( System.Web.Services.Protocols.SoapException excpt )
						{
							// catch error without interruption
							ErrorMessage = excpt.Message;
							iFailCount++;
							if( ErrorHandler( soql ) )
								return returnList;
						}
					}
				}
			}	// end while( iIndex < results.size )

			ReportStatus( "Finished querying ", soql, " / ", returnList.Count().ToString(), " rows" );

			return returnList;
		}

		public T QuerySingle<T>(string soql) where T : sObject, new()
		{
			T returnValue = null;

			if( ! SetupService() ) { return null; }

			QueryResult results = null;

			// attempt a certain number of times
			int iFailCount = 0;
			while( iFailCount < Retries )
			{
				try
				{
					results = salesforceService.query( soql );
					iFailCount = Retries;
				}
				catch( System.Web.Services.Protocols.SoapException excpt )
				{
					// catch error without interruption
					ErrorMessage = excpt.Message;
					iFailCount++;
					if( ErrorHandler( soql ) )
						return returnValue;
				}
			}

			if( results.size == 1 )
				returnValue = results.records[ 0 ] as T;
			else
				if( results.size > 1 )
					ErrorMessage = string.Concat( "WARNING:  query returned more than 1 record - query:  ", soql );
				else
					ErrorMessage = string.Concat( "WARNING:  query returned NO records - query:  ", soql );

			return returnValue;
		}

		public Guid QueryAsync(string soql)
		{
			if( ! SetupService() ) { return new Guid(); }

			salesforceService.queryCompleted += salesforceService_queryCompleted;
			
			Guid id = Guid.NewGuid();

			salesforceService.queryAsync(soql, id);

			return id;
		}

		void salesforceService_queryCompleted(object sender, queryCompletedEventArgs e)
		{
			Guid id = (Guid)e.UserState;
			List<sObject> results = e.Result.records.ToList();

			if (asyncResults.ContainsKey(id))
				asyncResults[id].AddRange(results);
			else
				asyncResults.Add((Guid)e.UserState, results);
		}

		public SaveResult[] Update(sObject[] items)
		{
			if( !SetupService() ) { return null; }

			ErrorMessage = "";

			// send out batches of 200 updates
			int iIndex = 0;
			List<SaveResult> objResults = new List<SaveResult>( items.Count() );
			while( iIndex < items.Count() )
			{
				sObject[] obj200Items = items.SubArray<sObject>( iIndex, 200 );
				SaveResult[] objPartialResult = null;

				// attempt a certain number of times
				int iFailCount = 0;
				while( iFailCount < Retries )
				{
					try
					{
						objPartialResult = salesforceService.update( obj200Items );
						iFailCount = Retries;
					}
					catch( System.Web.Services.Protocols.SoapException excpt )
					{
						// catch error without interruption
						ErrorMessage = excpt.Message;
						iFailCount++;
						if( ErrorHandler( string.Concat( "Update ", items.ToString(), ": ", ErrorMessage ) ) )
							return objResults.ToArray();
					}
				}

				objResults.AddRange( objPartialResult );

				iIndex += 200;
			}

			return objResults.ToArray();

			//return salesforceService.update(items);
		}

		public UpsertResult[] Upsert(string externalID, sObject[] items)
		{
			if( ! SetupService() ) { return null; }

			string strObjectName = "";
			if( items.Count() > 0 )
				strObjectName = items[ 0 ].GetType().Name;

			ReportStatus( "Initiating Upsert ", strObjectName );

			// send out batches of 200 upserts
			int iIndex = 0;
			List<UpsertResult> objResults = new List<UpsertResult>(items.Count());
			while (iIndex < items.Count())
			{
				sObject[] obj200Items = items.SubArray<sObject>(iIndex, 200);

				UpsertResult[] objPartialResult = null;

				// attempt a certain number of times
				int iFailCount = 0;
				while( iFailCount < Retries )
				{
					ErrorMessage = "";

					try
					{
						objPartialResult = salesforceService.upsert( externalID, obj200Items );
						iFailCount = Retries;

						if( items.Count() > 0 )
							ReportStatus( "Upserted ", strObjectName, " 200 rows from ", iIndex.ToString(), " (out of ", items.Count().ToString(), ")" );
					}
					catch( InvalidOperationException excpt )
					{
						ErrorMessage = excpt.Message;
						iFailCount++;
						if( ErrorHandler( string.Concat( "Error: Upsert ", strObjectName, ": ", ErrorMessage, " - ", iIndex.ToString() ) ) )
							return objResults.ToArray();
					}
					catch( System.Web.Services.Protocols.SoapException excpt )
					{
						// catch error without interruption
						ErrorMessage = excpt.Message;
						iFailCount++;
						if( ErrorHandler( string.Concat( "Error: Upsert ", strObjectName, ": ", ErrorMessage, " - ", iIndex.ToString() ) ) )
							return objResults.ToArray();
					}
				}

				// if successful, update the ids and store results 
				if( objPartialResult != null )
				{
					// update the ids in the source array
					for( int iPartialIndex = 0; iPartialIndex < objPartialResult.Count(); iPartialIndex++ )
						obj200Items[ iPartialIndex ].Id = objPartialResult[ iPartialIndex ].id;

					objResults.AddRange( objPartialResult );
				}
				else // if errored out, copy the error message to each of the results
					foreach( sObject objItem in obj200Items )
					{
						UpsertResult objUpsertResult = new UpsertResult();
						objUpsertResult.success = false;
						objUpsertResult.errors = new Error[ 1 ];
						objUpsertResult.errors[ 0 ] = new Error();
						objUpsertResult.errors[ 0 ].message = ErrorMessage;
						objResults.Add( objUpsertResult );
					}

				iIndex += 200;
			}

			ReportStatus( "Finished Upserting ", strObjectName, " ", items.Count().ToString(), " rows" );

			return objResults.ToArray();
		}

		public SaveResult[] Insert(sObject[] items)
		{
			if( !SetupService() ) { return null; }

			// send out batches of 200 inserts
			int iIndex = 0;
			List<SaveResult> objResults = new List<SaveResult>( items.Count() );
			while( iIndex < items.Count() )
			{
				sObject[] obj200Items = items.SubArray<sObject>( iIndex, 200 );

				SaveResult[] objPartialResult = null;

				int iFailCount = 0;
				while( iFailCount < Retries )
				{
					try
					{
						objPartialResult = salesforceService.create( obj200Items );
						iFailCount = Retries;
					}
					catch( System.Web.Services.Protocols.SoapException excpt )
					{
						ErrorMessage = excpt.Message;
						iFailCount++;
						if( ErrorHandler( string.Concat( "Insert ", items.ToString(), ": ", ErrorMessage ) ) )
							return objResults.ToArray();
					}
				}

				if( objPartialResult != null )
				{
					// update the ids in the source array
					for( int iPartialIndex = 0; iPartialIndex < objPartialResult.Count(); iPartialIndex++ )
						obj200Items[ iPartialIndex ].Id = objPartialResult[ iPartialIndex ].id;

					objResults.AddRange( objPartialResult );
				}
				else
					foreach( sObject objItem in obj200Items )
					{
						SaveResult objSaveResult = new SaveResult();
						objSaveResult.success = false;
						objSaveResult.errors = new Error[ 1 ];
						objSaveResult.errors[ 0 ] = new Error();
						objSaveResult.errors[ 0 ].message = ErrorMessage;
						objResults.Add( objSaveResult );
					}

				iIndex += 200;
			}

			return objResults.ToArray();

			//return salesforceService.create(items);
		}

		public DeleteResult[] Delete(string[] ids)
		{
			if( !SetupService() ) { return null; }

			// send out batches of 200 updates
			int iIndex = 0;
			List<DeleteResult> objResults = new List<DeleteResult>( ids.Count() );
			while( iIndex < ids.Count() )
			{
				string[] obj200Items = ids.SubArray( iIndex, 200 );

				objResults.AddRange( salesforceService.delete( obj200Items ) );

				iIndex += 200;

				ReportStatus( "Deleted:  ", iIndex.ToString(), " rows." );
			}

			return objResults.ToArray();

			//return salesforceService.delete(ids);
		}

		public DeleteResult[] QueryAndDelete<T>( string soql ) where T : sObject, new()
		{
			if( !SetupService() ) { return null; }

			ReportStatus( "Initiating Query And Delete" );
			
			List<T> returnList = new List<T>();

			T obj = new T();
			string strObjectName = obj.GetType().Name;

			int iIndex = 0;
			QueryResult results = salesforceService.query( soql );
			if( results != null )
				ReportStatus( "Querying for deletion of ", strObjectName, ":  ", soql
							, " / ", results.size.ToString(), " rows retrieved." );

			while( iIndex < results.size )
			{
				for( int i = 0; i < results.records.Count(); i++ )
				{
					T item = results.records[ i ] as T;

					if( item != null )
						returnList.Add( item );
					else
						iIndex++; // this is to avoid infinite loop
				}

				iIndex += results.records.Count();

				// read more records if needed
				if( iIndex < results.size && !results.done )
				{
					results = salesforceService.queryMore( results.queryLocator );
					if( results != null )
						ReportStatus( "Querying more for deletion of ", strObjectName, ":  "
									, iIndex.ToString(), " rows retrieved." );
				}

			}	// end while( iIndex < results.size )

			string[] ids = returnList.ConvertAll( t => t.Id ).ToArray();

			// send out batches of 200 updates
			iIndex = 0;
			List<DeleteResult> objResults = new List<DeleteResult>( ids.Count() );
			while( iIndex < ids.Count() )
			{
				string[] obj200Items = ids.SubArray( iIndex, 200 );

				DeleteResult[] objPartialResult = null;

				int iFailCount = 0;
				while( iFailCount < Retries )
				{
					try
					{
						objPartialResult = salesforceService.delete( obj200Items );
						iFailCount = Retries;

						ReportStatus( "Deleted ", strObjectName, " 200 rows from ", iIndex.ToString()
									, " (out of ", ids.Count().ToString(), ")" );
					}
					catch( System.Web.Services.Protocols.SoapException excpt )
					{
						ErrorMessage = excpt.Message;
						iFailCount++;
						if( ErrorHandler( string.Concat( "QueryAndDelete ", strObjectName, ": ", ErrorMessage ) ) )
							return objResults.ToArray();
					}
				}

				objResults.AddRange( objPartialResult );

				iIndex += 200;
			}

			ReportStatus( "Finished Deleting ", strObjectName, " ", ids.Count().ToString(), " rows" );

			return objResults.ToArray();
		}

		public UndeleteResult[] Undelete(string[] ids)
		{
			if( !SetupService() ) { return null; }

			// send out batches of 200 updates
			int iIndex = 0;
			List<UndeleteResult> objResults = new List<UndeleteResult>( ids.Count() );
			while( iIndex < ids.Count() )
			{
				string[] obj200Items = ids.SubArray( iIndex, 200 );

				objResults.AddRange( salesforceService.undelete( obj200Items ) );

				iIndex += 200;
			}

			return objResults.ToArray();

			//return salesforceService.undelete(ids);
		}

		private bool SetupService( bool bRefresh = false )
		{
			ErrorMessage = "";

			// if a refresh request was not specified, only login if a new session id is needed
			if( ! bRefresh )
				// if we already got a server url and session, nothing else needed
				if( ! ServerURL.Equals( "" ) && ! SessionID.Equals( "" ) )
					return true;

			ForceConnection connection = new ForceConnection(ConnectionString);
			salesforceService.SessionHeaderValue =
				new SessionHeader() { sessionId = connection.SessionID };

			if( connection.ConnectionFailed )
			{
				ConnectionFailed = connection.ConnectionFailed;
				ErrorMessage = string.Concat( connection.ErrorMessage, " when using URL ", salesforceService.Url );

				return false;
			}

			salesforceService.Url = connection.ServerUrl;
			ServerURL = connection.ServerUrl;
			SessionID = connection.SessionID;

			// disable Chatter tracking of new records
			DisableFeedTrackingHeader objSoapHeader = new DisableFeedTrackingHeader();
			objSoapHeader.disableFeedTracking = false;
			salesforceService.DisableFeedTrackingHeaderValue = objSoapHeader;

			return true;
		}

		public void  Dispose()
		{
			Dispose( true );
			GC.SuppressFinalize( this );

		}
		protected virtual void Dispose( bool bCleanManaged = true )
		{
			OnError = null;
			OnReportStatus = null;
			salesforceService.Dispose();
		}
	}
}