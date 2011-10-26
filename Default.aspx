<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="EmTrac2SF.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>EmForce Data Load</title>
</head>
<body style="font-family: Arial, Helvetica, sans-serif; font-size: x-small">
	<form id="form1" runat="server">
	<div>
	
		Export from EmTrac to SalesForce:&nbsp;&nbsp;&nbsp;
		<table style="width:100%; padding: 1px 1px 1px 1px; ">
			<tr>
				<td>
		<asp:Button ID="btnAddEmCareAccount" runat="server" 
			onclick="btnAddEmCareAccount_Click" Text="Initialize Account & Settings" Width="172px" 
						style="font-size: small" />
				</td>
				<td>
					<asp:Button ID="btnRefreshFacilities" runat="server" 
			onclick="btnRefreshFacilities_Click" Text="Refresh Hierarchy/Facilities" Width="172px" />
				</td>
				<td>
		<asp:Button ID="btnRefreshProviders" runat="server" 
			onclick="btnRefreshProviders_Click" Text="Refresh Providers/Notes" Width="172px" /><br />			
		<asp:Button ID="btnRefreshProvidersOnly" runat="server" 
			onclick="btnRefreshProvidersOnly_Click" Text="Refresh Providers" Width="172px" />			
		<asp:Button ID="btnRefreshNotes" runat="server" 
			onclick="btnRefreshNotes_Click" Text="Refresh Notes" Width="172px" />
				</td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="btnRefreshSubTypes" runat="server" 
			onclick="btnRefreshSubTypes_Click" Text="Refresh SubTypes" Width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshAgencies" runat="server" 
			onclick="btnRefreshAgencies_Click" Text="Refresh Agencies" Width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshEducationExperience" runat="server" 
			onclick="btnRefreshEducationExperience_Click" Text="Refresh Degree/Experience" 
						Width="172px" />
				</td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="btnRefreshCredentials" runat="server" 
			onclick="btnRefreshCredentials_Click" Text="Refresh Credentials" Width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshCandidate" runat="server" 
			onclick="btnRefreshCandidate_Click" Text="Refresh Candidate" Width="172px" />
				</td>
				<td>
					<asp:Button ID="btnUpdateUsers" runat="server" 
			onclick="btnUpdateUsers_Click" Text="Update Users Emails" Width="172px" />
				</td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="btnRefreshResidency" runat="server" 
						onclick="btnRefreshResidency_Click" Text="Refresh Resid. Programs" width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshResidents" runat="server" 
						onclick="btnRefreshResidents_Click" Text="Refresh Residents" width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshReferences" runat="server" 
						onclick="btnRefreshReferences_Click" Text="Refresh References" width="172px" />
				</td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="btnReportDuplicateInstitutions" runat="server" 
						onclick="btnReportDuplicateInstitutions_Click" Text="Report Dupe Institutions" 
						width="172px" />
				</td>
				<td>
					<asp:Button ID="btnReportDuplicateAgencies" runat="server" 
						onclick="btnReportDuplicateAgencies_Click" Text="Report Dupe Agencies" width="172px" />
				</td>
				<td>
					<asp:Button ID="btnRefreshUsers" runat="server" 
			onclick="btnRefreshUsers_Click" Text="Refresh Users (optional)" Width="172px" />
				</td>
			</tr>
			<tr>
				<td>
		<asp:Button ID="btnRefreshAMAProviders" runat="server" 
			onclick="btnRefreshAMAProviders_Click" Text="Refresh AMA Providers" />
				</td>
				<td>
		<asp:Button ID="btnRefreshAMACredentials" runat="server" 
			onclick="btnRefreshAMACredentials_Click" Text="Refresh AMA Credentials" />		
				</td>
				<td>
					<asp:Button ID="btnBulkLoad" runat="server" onclick="btnBulkLoad_Click" 
						Text="Bulk Operation" />
				</td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="btnMassDataLoad" runat="server" 
						onclick="btnMassDataLoad_Click" Text="Mass Data Load" width="172px" />
				</td>
				<td>
					Table:
					<asp:DropDownList ID="ddlTable" runat="server">
						<asp:ListItem Text="- Not Selected -" Value="" />
						<asp:ListItem Text="Contacts" Value="Contact" />
						<asp:ListItem Text="Facilities" Value="Facility__c" />
						<asp:ListItem Text="Credentials" Value="Credential__c" />
						<asp:ListItem Text="Education/Experience" Value="Education_or_Experience__c" />
						<asp:ListItem Text="Agencies" Value="Credential_Agency__c" />
						<asp:ListItem Text="Institutions" Value="Institution__c" />
						<asp:ListItem Text="Subtypes" Value="Credential_Subtype__c" />
						<asp:ListItem Text="Candidates" Value="Candidate__c" />
						<asp:ListItem Text="Residency Programs" Value="Residency_Program__c" />
						<asp:ListItem Text="Residents" Value="Resident__c" />
						<asp:ListItem Text="References" Value="Reference__c" />
						<asp:ListItem Text="Job Applications" Value="Job_Application__c" />
						<asp:ListItem Text="ContactFeed" Value="ContactFeed" />
						<asp:ListItem Text="Facility__Feed" Value="Facility__Feed" />
						<asp:ListItem Text="Candidate Stage Tracking" Value="Candidate_Stage_Tracking__c" />
						<asp:ListItem Text="Provider_Contract__Feed" Value="Provider_Contract__Feed"></asp:ListItem>
						<asp:ListItem Text="Provider_Contract__History" Value="Provider_Contract__History"></asp:ListItem>
						<asp:ListItem Text="Candidate__Feed" Value="Candidate__Feed"></asp:ListItem>
						<asp:ListItem Text="Reference__History" Value="Reference__History" />
					</asp:DropDownList>
&nbsp;<asp:Button ID="btnDelete" runat="server" Text="Delete" onclick="btnDelete_Click" 
						width="79px" />
					<br />
					<asp:Button ID="btnRefreshInterviews" runat="server" 
			onclick="btnRefreshInterviews_Click" Text="Refresh Interviews (optional)" Width="172px" 
						Visible="False" />
				</td>
				<td>
					<asp:CheckBox ID="cbxAllRecords" runat="server" Text="All Records" />
				&nbsp;
					<asp:CheckBox ID="cbxDisplayOnly" runat="server" Text="Display Only" 
						Checked="True" />
				&nbsp;
					<asp:CheckBox ID="cbxVerbose" runat="server" Text="Verbose" Visible="False" />
					<br />
					<asp:HyperLink ID="hlnkStatus" runat="server" NavigateUrl="Status.html" 
						Target="Status">Right click open here to see status</asp:HyperLink>
					<br />
					<asp:HyperLink ID="hlnkActivityHistory" runat="server" NavigateUrl="ActivityHistory.txt" 
						Target="History">Right click open here to see Activity History</asp:HyperLink>
				</td>
			</tr>
		</table>
		<br />
		Results:
		<asp:Label ID="lblError" runat="server" Font-Bold="True" Font-Size="Small" 
			ForeColor="Red"></asp:Label>
		<br />
		<asp:TextBox ID="tbStatus" runat="server" Rows="10" TextMode="MultiLine" 
			Width="841px" Height="340px" Wrap="False"></asp:TextBox>
	
		<br />
		<div style="visibility: hidden;">Data:<br /></div>
		<asp:TextBox ID="tbData" runat="server" Rows="10" TextMode="MultiLine" 
			Width="841px" Wrap="False" Visible="False"></asp:TextBox>
	
	</div>
	<div id="OpenCSVList" name="OpenCSVList" runat="server" ></div>
	</form>
</body>
</html>
