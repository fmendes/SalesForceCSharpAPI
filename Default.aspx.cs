using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using EmTrac2SF.Salesforce;
using EmTrac2SF.EMSC2SF;
using Metaphone;
using GenericLibrary;
using System.Text;
using System.IO;

namespace EmTrac2SF
{
	public partial class Default : System.Web.UI.Page
	{
		DataLoader objDL;
		string strInstance = "";
		DateTime dtBegin, dtEnd;

		protected void Page_Load(object sender, EventArgs e)
		{
			objDL = new DataLoader( Request.PhysicalApplicationPath, tbStatus, lblError, OpenCSVList );

			if (Session["ApiService"] != null)
				objDL.API = (ApiService) Session[ "ApiService" ];
			else
			{
				objDL.API = new ApiService();
				Session[ "ApiService" ] = objDL.API;
			}

			// auto configure API according to the setting in the web.config
			strInstance = System.Configuration.ConfigurationManager.AppSettings[ "Instance" ].ToLower();
			switch( strInstance )
			{
				case "dev1":
					Title = "EmTrac2SF - DEV1";
					break;
				case "test1":
					Title = "EmTrac2SF - TEST1";
					break;
				case "prod":
					if(Properties.Settings.Default.EmTrac2SF_EMSC_SF_SforceService.StartsWith( "https://login.salesforce.com" ))
						Title = "EmTrac2SF - PRODUCTION";
					else
						lblError.Text = "ERROR:  In order to connect to Production, please switch the Salesforce URL to https://login.salesforce.com in the web.config.";
					break;
			}
		}

		protected void btnRefreshProviders_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			objDL.RefreshProviders( bImportNotes: true, bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Provider load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshFacilities_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			objDL.RefreshFacilities( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Facility load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshCredentials_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			if (cbxAllRecords.Checked)
				objDL.RefreshCredentials( null, null, null, cbxDisplayOnly.Checked );
			else
			{
				List<Contact> objProviders = objDL.RefreshProviders( false );
				objDL.RefreshCredentials( objProviders, objDL.RefreshAgencies( objProviders ), objDL.RefreshSubtypes(), cbxDisplayOnly.Checked );
			}
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Credentials load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshEducationExperience_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			if (cbxAllRecords.Checked)
				objDL.RefreshEducationExperience( bDisplayOnly: cbxDisplayOnly.Checked );
			else
			{
				List<Contact> objProviders = EmTrac2SFUtils.GetProvidersFromSF( objDL.API, lblError );
				objDL.RefreshEducationExperience( objProviders
												, objDL.RefreshInstitutions( objProviders, bDisplayOnly: cbxDisplayOnly.Checked ), bDisplayOnly: cbxDisplayOnly.Checked );
			}
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Education/Experience load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnAddEmCareAccount_Click(object sender, EventArgs e)
		{
			objDL.InitializeAccountAndSettings();
		}

		protected void btnRefreshCandidate_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			if (cbxAllRecords.Checked)
				objDL.RefreshCandidates( bDisplayOnly: cbxDisplayOnly.Checked );
			else
			{
				List<Contact> objProviders = objDL.RefreshProviders( false, bDisplayOnly: cbxDisplayOnly.Checked );
				//List<Facility__c> objFacilities = objDL.RefreshFacilities();
				objDL.RefreshCandidates( objProviders, bDisplayOnly: cbxDisplayOnly.Checked );
			}
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Candidate load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshAgencies_Click(object sender, EventArgs e)
		{
			dtBegin = DateTime.Now;
			if (cbxAllRecords.Checked)
			{
				objDL.RefreshAgencies( bDisplayOnly:  cbxDisplayOnly.Checked );
				objDL.RefreshInstitutions( bDisplayOnly: cbxDisplayOnly.Checked );
			}
			else
			{
				List<Contact> objProviders = objDL.RefreshProviders( false, bDisplayOnly: cbxDisplayOnly.Checked );
				objDL.RefreshAgencies( objProviders, bDisplayOnly: cbxDisplayOnly.Checked );
				objDL.RefreshInstitutions( objProviders, bDisplayOnly: cbxDisplayOnly.Checked );
			}
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Agencies/Institutions load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshSubTypes_Click(object sender, EventArgs e)
		{
			objDL.RefreshSubtypes( bDisplayOnly: cbxDisplayOnly.Checked );
		}

		protected void btnRefreshInterviews_Click(object sender, EventArgs e)
		{
			//RefreshInterviews();
		}

		protected void btnRefreshUsers_Click(object sender, EventArgs e)
		{
			objDL.RefreshUsers( bDisplayOnly: cbxDisplayOnly.Checked );
		}

		protected void btnUpdateUsers_Click( object sender, EventArgs e )
		{
			objDL.UpdateUsersEmail();
		}

		protected void btnReportDuplicateAgencies_Click( object sender, EventArgs e )
		{
			objDL.ReportDuplicateAgencies();
		}

		protected void btnReportDuplicateInstitutions_Click( object sender, EventArgs e )
		{
			objDL.ReportDuplicateInstitutions();
		}

		protected void btnRefreshResidency_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.RefreshResidencyPrograms( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Residency Programs load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshResidents_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.RefreshResidents( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** Residents load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshReferences_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.RefreshReferences( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** References load completed in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnDelete_Click( object sender, EventArgs e )
		{
			if( ddlTable.SelectedValue == null || ddlTable.SelectedValue.Equals( "" ) )
				return;

			dtBegin = DateTime.Now;
			string strTable = ddlTable.SelectedValue;
			int iCount = EmTrac2SFUtils.DeleteTable( objDL.API, strTable, tbStatus );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** ", iCount, " ", strTable, " records deleted in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnMassDataLoad_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.MassDataLoad();
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** ", " All records updated in ", tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnRefreshNotes_Click( object sender, EventArgs e )
		{
			objDL.RefreshProviderNotes( bDisplayOnly: cbxDisplayOnly.Checked );
		}

		protected void btnRefreshProvidersOnly_Click( object sender, EventArgs e )
		{
			objDL.RefreshProviders( bImportNotes: false, bDisplayOnly: cbxDisplayOnly.Checked );
		}

		protected void btnRefreshAMAProviders_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.RefreshAMAProviders( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** AMA Providers load completed in "
				, tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );			
		}

		protected void btnRefreshAMACredentials_Click( object sender, EventArgs e )
		{
			dtBegin = DateTime.Now;
			objDL.RefreshAMAProviderCredentials( bDisplayOnly: cbxDisplayOnly.Checked );
			dtEnd = DateTime.Now;
			TimeSpan tsDuration = dtEnd.Subtract( dtBegin );
			tbStatus.Text = string.Concat( tbStatus.Text, "\r\n\r\n** AMA Credentials load completed in "
				, tsDuration.Hours, " hours "
				, tsDuration.Minutes, " minutes. **\r\n" );
		}

		protected void btnBulkLoad_Click( object sender, EventArgs e )
		{
			objDL.RemoveAMADuplicates();

			//objDL.LoadCredentialsReportedWithErrors();

			//objDL.LoadMissingCredentials();
			//objDL.CleanJobAppsByRecrID();

			//objDL.RefreshAgencies( null, false );
			//objDL.RemoveDuplicateNotes();
			//objDL.CleanCandidates();
			//objDL.RefreshCandidates( null, null, false );

			//objDL.BulkLoadMissingEmTracNotes();
			//objDL.BulkLoadJobAppFromEmTrac();
			//objDL.BulkLoadJobAppFromAMA();
			//objDL.BulkLoadProviderFromEmTrac();
			//objDL.BulkLoadProviderNotesFromEmTrac();
			//objDL.CleanJobAppsByMENbr();
		}

	}
}