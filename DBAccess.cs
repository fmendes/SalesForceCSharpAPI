using System;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Configuration;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Web.Caching;

namespace EmTrac2SF
{
	/// <summary>
	/// Class with static methods to handle all database interaction
	/// </summary>
	public class DBAccess : IDisposable
	{
		public string MainDB = "EmTrac";
		public const string strLogFile = "dberrlog.txt";
		public const int iLogLimit = 5000;

		public string strPath = HttpContext.Current.Request.PhysicalApplicationPath;
		public static object objSync = new object();
		public string strLogFlag = System.Configuration.ConfigurationManager.AppSettings["LogDBErrors"];
		public SqlConnection objConn, objCodsScrubConn;
		public IsolationLevel objIsolationLevel = IsolationLevel.Unspecified;

		public string ErrorMessage = "";

		public Label ErrorLabel;
		public Cache Cache;

		public static int iConnectionRetries = Convert.ToInt32( System.Configuration.ConfigurationManager.AppSettings["DBRetries"] );

		public void Dispose()
		{
			Dispose( true );

			if( objConn != null )
				objConn.Dispose();
			if( objCodsScrubConn != null )
				objCodsScrubConn.Dispose();

			GC.SuppressFinalize( this );
		}

		protected virtual void Dispose( bool bNative = true )
		{
			
		}

		public string GetSQLFromFile( string strFile )
		{
			return System.IO.File.ReadAllText(string.Concat(strPath, "", strFile));
		}

		public DataTable GetDataTableFromSQL( string strSQL )
		{
			DataSet objDS = GetDataSetFromSQL( strSQL );

			if( !ErrorMessage.Equals( "" ) || objDS.Tables.Count == 0 )
				objDS.Tables.Add( "Dummy" );

			return objDS.Tables[ 0 ];
		}

		public DataTable GetDataTableFromSQLFile(string strFileName, string strPrependScript = null, string strCondition = null )
		{
			string strSQL = GetSQLFromFile(strFileName);
			
			if( strPrependScript != null )
				strSQL = string.Concat(strPrependScript, strSQL);

			if( strCondition != null )
				strSQL = strSQL.Replace( "-- condition placeholder", strCondition );

			DataSet objDS = GetDataSetFromSQL(strSQL);

			if (!ErrorMessage.Equals("") || objDS.Tables.Count == 0)
				objDS.Tables.Add("Dummy");

			return objDS.Tables[0];
		}

		#region Misc methods

		public void ReportError(string strErrorMsg)
		{
			if (strErrorMsg.IndexOf("Timeout expired.") >= 0)
				strErrorMsg = "Timeout expired:  The server took too long to complete the operation. Please try again.";

			ErrorMessage = strErrorMsg;

			string strUser = "";
			string strPage = "";
			if (ErrorLabel != null)
			{
				ErrorLabel.Text = strErrorMsg;

				if (strLogFlag == null)
					return;
				if (strLogFlag.Equals("0"))
					return;

				strUser = ErrorLabel.Page.User.Identity.Name;
				strPage = ErrorLabel.Page.AppRelativeVirtualPath;
			}

			StreamReader objFS = null;
			StreamWriter objSW = null;
			try
			{
				string strFile = String.Format( "{0}{1}", strPath, strLogFile );
				string strLog = "";

				FileInfo objFI = new FileInfo( strFile );

				if( objFI.Exists )
				{
					lock( objSync )
					{
						objFS = new StreamReader( strFile );
						strLog = objFS.ReadToEnd();
						objFS.Close();
					}
				}

				strLog = String.Format( "{0}\t{1}\t{2}\t{3}\n{4}"
								, DateTime.Now.ToString(), strUser
								, strPage, strErrorMsg, strLog );

				if( strLog.Length > iLogLimit )
					strLog = strLog.Remove( iLogLimit );

				lock( objSync )
				{
					objSW = new StreamWriter( strFile );
					objSW.Write( strLog );
					objSW.Close();
				}
			}
			catch( Exception )
			{
				// ignore errors while writing to the log file, we don't want to bother
			}
			finally
			{
				if( objFS != null )
					objFS.Dispose();
				if( objSW != null )
					objSW.Dispose();
			}
		}

		/// <summary>
		/// Does the same thing as 'string.Format' but duplicating single quotes so it will be accepted
		/// in the SQL statement context. Example:  O'Hare will become O''Hare
		/// </summary>
		/// <param name="strFormat">String containing zero or more format items.</param>
		/// <param name="args">A list of strings to fill in the format string.</param>
		/// <returns>A formatted string.</returns>
		public static string StringFormat(string strFormat, params string[] args)
		{
			for (int iIdx = 0; iIdx < args.Length; iIdx++)
			{
				// replace single quote with double quotes
				args[iIdx] = args[iIdx].Replace("'", "''");
			}

			return string.Format(strFormat, args);
		}

		public static string ReplaceApostrophes(string strValue)
		{
			// try to replace single apostrophes with double apostrophes
			// first, remove double apostrophes just in case
			strValue = strValue.Replace("''", "'");
			strValue = strValue.Replace("'", "''");

			return strValue;
		}

		#endregion

		#region Connection methods
		public static string GetConnectionString(string strName)
		{
			string strConnString = System.Configuration.ConfigurationManager.ConnectionStrings[strName].ToString();

			return strConnString;
		}
		public void SetConnectionRef(string strName, SqlConnection sc)
		{
			if (strName.Equals(MainDB))
				objConn = sc;
			//switch (strName)
			//{
			//    case MainDB:
			//        objConn = sc;
			//        break;
			//    //case CADDATA:
			//    //	objCADConn = sc;
			//    //	break;
			//    //case ENTERPRISE:
			//    //	objEnterpriseConn = sc;
			//    //	break;
			//    //case EPCR:
			//    //	objEPCRConn = sc;
			//    //	break;

			//    default:
			//        return;
			//}
			return;
		}
		public SqlConnection GetConnectionRef(string strName)
		{
			SqlConnection sc = null;

			if (strName.Equals(MainDB))
				sc = objConn;

			//switch (strName)
			//{
			//    case MainDB:
			//        sc = objConn;
			//        break;
			//    //case CADDATA:
			//    //	sc = objCADConn;
			//    //	break;
			//    //case ENTERPRISE:
			//    //	sc = objEnterpriseConn;
			//    //	break;
			//    //case EPCR:
			//    //	sc = objEPCRConn;
			//    //	break;

			//    default:
			//        return null;
			//}
			return sc;
		}
		/// <summary>
		/// Creates a SqlConnection using the given Connection String
		/// </summary>
		/// <param name="strName">A connection string</param>
		public SqlConnection GetConnection(string strName)
		{
			SqlConnection sc = new SqlConnection(GetConnectionString(strName));
			SetConnectionRef(strName, sc);
			return sc;
		}
		/// <summary>
		/// Opens the connection to the database using the given Connection String
		/// </summary>
		/// <param name="strName">A Connection String</param>
		public bool TryConnect(string strName)
		{
			SqlConnection sc = GetConnectionRef(strName);

			if (sc == null)
				return false;

			// due to intermittent connection errors on test,
			// this is gonna try several times then give it up
			for (int iCount = 0; iCount < iConnectionRetries; iCount++)
			{
				ErrorMessage = "";
				try
				{
					sc.Open();
				}
				catch (Exception excpt)
				{
					ReportError(excpt.Message);
				}

				if (sc.State.Equals(ConnectionState.Open))
					return true;
			}

			return sc.State.Equals(ConnectionState.Open);
		}

		public string GetConnectionString()
		{
			string strConnString = System.Configuration.ConfigurationManager.ConnectionStrings[MainDB].ToString();

			return strConnString;
		}
		/// <summary>
		/// Creates a SqlConnection using the Default Connection String
		/// </summary>
		public SqlConnection GetConnection()
		{
			objConn = new SqlConnection(GetConnectionString());

			return objConn;
		}
		/// <summary>
		/// Opens the connection to the database
		/// </summary>
		public bool TryConnect()
		{
			if (objConn == null)
				return false;

			// due to intermittent connection errors on test,
			// this is gonna try several times then give it up
			for (int iCount = 0; iCount < iConnectionRetries; iCount++)
			{
				ErrorMessage = "";
				try
				{
					objConn.Open();
				}
				catch (Exception excpt)
				{
					ReportError(excpt.Message);
				}

				if (objConn.State.Equals(ConnectionState.Open))
					return true;
			}

			return objConn.State.Equals(ConnectionState.Open);
		}

		private void CommitAndDisconnect(ref SqlTransaction objTrans)
		{
			// commit transaction if it exists, before disconnecting
			if (objTrans != null)
			{
				objTrans.Commit();
				objTrans = null;
			}

			objConn.Close();
			objConn.Dispose();
		}
		private void RollbackAndReportError(ref SqlTransaction objTrans, Exception excpt)
		{

			// rollback transaction if it exists
			if (objTrans != null)
			{
				objTrans.Rollback();
				objTrans = null;
			}

			ReportError(excpt.Message);
		}

		#endregion

		#region Query/DML methods
		public DataSet GetDataSetFromSQL(string strSQL)
		{
			DataSet ds = new DataSet();

			// add code to retry if dataset is empty
			int iRetries = 0;
			while (iRetries < iConnectionRetries)
			{
				iRetries++;

				ErrorMessage = "";

				GetConnection();

				SqlTransaction objTrans = null;
				SqlDataAdapter objAdapter = null;
				try
				{
					TryConnect();

					objAdapter = new SqlDataAdapter(strSQL, objConn);

					// if isolation level was set, open a transaction with the isolation level
					if (!objIsolationLevel.Equals(IsolationLevel.Unspecified))
					{
						objTrans = objConn.BeginTransaction(objIsolationLevel);
						objAdapter.SelectCommand.Transaction = objTrans;
					}

					objAdapter.Fill(ds);
				}
				catch (Exception excpt)
				{
					RollbackAndReportError(ref objTrans, excpt);
				}
				finally
				{
					CommitAndDisconnect(ref objTrans);
					objAdapter.Dispose();
				}

				// if successful, return to caller
				if (ds.Tables.Count > 0)
					return ds;
			}

			return ds;
		}

		public SqlDataReader GetSqlDataReaderFromSQL( string strSQL )
		{
			GetConnection();
			SqlCommand objComm = new SqlCommand(strSQL, objConn);
			SqlDataReader objReader = null;

			// add code to retry if dataset is empty
			int iRetries = 0;
			while (iRetries < iConnectionRetries)
			{
				iRetries++;

				ErrorMessage = "";

				SqlTransaction objTrans = null;

				try
				{
					TryConnect();

					// if isolation level was set, open a transaction with the isolation level
					if( !objIsolationLevel.Equals( IsolationLevel.Unspecified ) )
					{
						objTrans = objConn.BeginTransaction( objIsolationLevel );
						objComm.Transaction = objTrans;
					}

					objReader = objComm.ExecuteReader();
				}
				catch( Exception excpt )
				{
					RollbackAndReportError( ref objTrans, excpt );
				}
				finally
				{
					//objComm.Dispose();
				}

				// if successful, return to caller
				if (objReader != null)
					return objReader;
			}

			return objReader;
		}

		public string GetStringFromSQL( string strSQL )
		{
			// add code to retry if dataset is empty
			int iRetries = 0;
			while (iRetries < iConnectionRetries)
			{
				iRetries++;

				ErrorMessage = "";

				GetConnection();
				SqlCommand objComm = new SqlCommand(strSQL, objConn);
				object objResult = null;

				SqlTransaction objTrans = null;

				try
				{
					TryConnect();

					// if isolation level was set, open a transaction with the isolation level
					if (!objIsolationLevel.Equals(IsolationLevel.Unspecified))
					{
						objTrans = objConn.BeginTransaction(objIsolationLevel);
						objComm.Transaction = objTrans;
					}

					objResult = objComm.ExecuteScalar();
				}
				catch (Exception excpt)
				{
					RollbackAndReportError(ref objTrans, excpt);
				}
				finally
				{
					CommitAndDisconnect(ref objTrans);
					objComm.Dispose();
				}

				// if successful, return to caller
				if (ErrorMessage.Equals(""))
					if (objResult != null)
						if (!Convert.IsDBNull(objResult))
							return objResult.ToString();
						else
							return "";
			}

			return "";
		}

		public int GetIntegerFromSQL( string strSQL )
		{
			// add code to retry if dataset is empty
			int iRetries = 0;
			while (iRetries < iConnectionRetries)
			{
				iRetries++;

				ErrorMessage = "";

				GetConnection();
				SqlCommand objComm = new SqlCommand(strSQL, objConn);
				object objResult = null;

				SqlTransaction objTrans = null;

				try
				{
					TryConnect();

					// if isolation level was set, open a transaction with the isolation level
					if (!objIsolationLevel.Equals(IsolationLevel.Unspecified))
					{
						objTrans = objConn.BeginTransaction(objIsolationLevel);
						objComm.Transaction = objTrans;
					}

					objResult = objComm.ExecuteScalar();
				}
				catch (Exception excpt)
				{
					RollbackAndReportError(ref objTrans, excpt);
				}
				finally
				{
					CommitAndDisconnect(ref objTrans);
					objComm.Dispose();
				}

				// if successful, return to caller
				if (ErrorMessage.Equals(""))
					if (objResult != null)
						if (!Convert.IsDBNull(objResult))
							return Convert.ToInt32(objResult);
						else
							return 0;
			}

			return -1;
		}

		public int RunSQL( string strSQL )
		{
			ErrorMessage = "";

			GetConnection();
			SqlCommand objComm = new SqlCommand(strSQL, objConn);
			int iResult = 0;

			SqlTransaction objTrans = null;

			try
			{
				TryConnect();

				// if isolation level was set, open a transaction with the isolation level
				if (!objIsolationLevel.Equals(IsolationLevel.Unspecified))
				{
					objTrans = objConn.BeginTransaction(objIsolationLevel);
					objComm.Transaction = objTrans;
				}

				iResult = objComm.ExecuteNonQuery();
			}
			catch (Exception excpt)
			{
				RollbackAndReportError(ref objTrans, excpt);
			}
			finally
			{
				CommitAndDisconnect(ref objTrans);
			}

			return iResult;
		}

		public void UpdateAdapter( DataSet ds, SqlDataAdapter objSqlAdapter, SqlCommand objComm )
		{
			GetConnection();

			ErrorMessage = "";

			objSqlAdapter.InsertCommand = objComm;

			try
			{
				TryConnect();
				objComm.Connection = objConn;

				// copy all rows from the dataset to the DIV_HIER table
				objSqlAdapter.Update(ds.Tables[0]);
			}
			catch (Exception excpt)
			{
				ReportError(excpt.Message);
			}
			finally
			{
				objConn.Close();
				objConn.Dispose();
			}

			return;
		}

		public void UpdateAdapterRowCheck( DataSet ds, SqlDataAdapter objSqlAdapter, SqlCommand objComm )
		{
			GetConnection();

			ErrorMessage = "";

			objSqlAdapter.InsertCommand = objComm;

			try
			{
				TryConnect();
				objComm.Connection = objConn;

				objSqlAdapter.RowUpdating += new SqlRowUpdatingEventHandler(objSqlAdapter_RowUpdating);
				objSqlAdapter.RowUpdated += new SqlRowUpdatedEventHandler(objSqlAdapter_RowUpdated);

				// copy all rows from the dataset to the DIV_HIER table
				objSqlAdapter.Update(ds.Tables[0]);
			}
			catch (Exception excpt)
			{
				DataColumn[] objCols = objLastUpdateRow.GetColumnsInError();
				string strColumn = objCols[0].ToString();
				string strError = string.Concat(excpt.Message, "\n", strColumn, " = ", objLastUpdateRow[strColumn].ToString());
				ReportError(strError);
			}
			finally
			{
				objConn.Close();
				objConn.Dispose();
			}

			return;
		}
		public DataRow objLastUpdateRow = null;
		public void objSqlAdapter_RowUpdated( object sender, SqlRowUpdatedEventArgs e )
		{
		}
		public void objSqlAdapter_RowUpdating( object sender, SqlRowUpdatingEventArgs e )
		{
			objLastUpdateRow = e.Row;
		}



		#endregion
	}

}
