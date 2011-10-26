using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace EmTrac2SF
{
	public partial class DisplayCSV : System.Web.UI.Page
	{
		protected void Page_Load( object sender, EventArgs e )
		{
			// receive file name as parameter
			string strFile = Request.Params[ "file" ];

			if( strFile == null || strFile.Equals( "" ) )
			{
				Response.Write( "Missing parameter: CSV file name." );
				return;
			}

			if( !System.IO.File.Exists( strFile ) )
			{
				Response.Write( string.Concat( "File does not exist:  ", strFile ) );
				return;
			}
			// read the file into a string
			string strText = System.IO.File.ReadAllText( strFile );
			// test value "C:\\NewDev\\EmTrac2SF\\EmTrac2SF\\EmTrac2SF\\Results_Credential_Subtype.csv"

			strText = string.Concat( strText.Replace( "\t", "\",\"" ).Replace( "\r\n", "\"\r\n\"" ), "\r\n" );
			strFile = strFile.Replace( "C:\\NewDev\\EmTrac2SF\\EmTrac2SF\\EmTrac2SF\\", "" );

			Response.Clear();
			Response.AddHeader( "content-disposition", string.Format( "attachment; filename={0}.csv", strFile ) );
			Response.ContentType = "text/csv";
			Response.Write( strText );
			Response.End();

		}
	}
}