using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace EmTrac2SF
{
	public partial class Status : System.Web.UI.Page
	{
		protected void Page_Load( object sender, EventArgs e )
		{
			objDiv.InnerHtml = ( ( Session[ "Status" ] == null ) ? "" : Session[ "Status" ].ToString() ).Replace( "\r\n", "<br />" );
		}
	}
}