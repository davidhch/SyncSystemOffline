<%
nCID = request.querystring("CID")
nCampID = request.querystring("CAMPID")

nExportRecordCount = request.querystring("RECS")
nExportRecordCount = cLng(nExportRecordCount)
strSQL = ""
nCurrentBlock = 0

Do While nCurrentBlock < nExportRecordCount AND MaxLoops < 100 ' we limit our export to 300,000

	strSQL = strSQL & " SELECT *, View_CUSTOMFIELDS.CustomFieldName AS 'Custom_Field_Name', View_CUSTOMFIELDS.CustomFieldData AS 'Custom_Field_Data' FROM (SELECT ROW_NUMBER() OVER (ORDER BY [View_ReportSummary].[EmailAddress]) AS Row, MIN(DISTINCT[View_ReportSummary].[ClickedDateTime]) AS [ClickedDateTime], [View_ReportSummary].[EmailID], [View_ReportSummary].[OpenedDateTime], [View_ReportSummary].[SuccessfullySent], [View_ReportSummary].[Bounced], [View_ReportSummary].[Opened], [View_ReportSummary].[EmailAddress], [View_ReportSummary].[CustomerID], [View_ReportSummary].[Clicked], [View_ReportSummary].[NoTimesOpened] , EmailAddresses.Title AS Subscriber_Title, EmailAddresses.FirstName AS Subscriber_First_Name, EmailAddresses.LastName AS Subscriber_Last_Name, EmailAddresses.EmailList, EmailAddresses.EmailType, EmailAddresses.MobilePhone, EmailAddresses.CompanyName, EmailAddresses.DateAdded, EmailAddresses.Source, EmailAddresses.Confirmed, EmailAddresses.ConfirmDate, EmailAddresses.[Use],EmailAddresses.Unsubscribed,EmailAddresses.UnsubscribeDate,Addresses.FirstName, Addresses.LastName, Addresses.Title, Addresses.JobTitle, Addresses.Company, Addresses.Address1, Addresses.Address2, Addresses.Address3, Addresses.City, Addresses.County, Addresses.Postcode, Addresses.Country, Addresses.TelNo, Addresses.FaxNo FROM [View_ReportSummary] LEFT OUTER JOIN EmailAddresses ON View_ReportSummary.EmailID = EmailAddresses.ID LEFT OUTER JOIN Addresses ON View_ReportSummary.EmailID = Addresses.EmailID WHERE [View_ReportSummary].[CampaignID]="&nCampID&"  AND [View_ReportSummary].[EmailAddress] LIKE '%'  GROUP BY Addresses.FirstName, Addresses.LastName, Addresses.Title, Addresses.JobTitle, Addresses.Company, Addresses.Address1, Addresses.Address2, Addresses.Address3, Addresses.City, Addresses.County, Addresses.Postcode, Addresses.Country, Addresses.TelNo, Addresses.FaxNo, EmailAddresses.Title, EmailAddresses.FirstName, EmailAddresses.LastName, EmailAddresses.EmailType, EmailAddresses.MobilePhone, EmailAddresses.CompanyName, EmailAddresses.DateAdded, EmailAddresses.Source, EmailAddresses.ConfirmDate, EmailAddresses.EmailList, EmailAddresses.Confirmed, EmailAddresses.[Use],EmailAddresses.Unsubscribed,EmailAddresses.UnsubscribeDate,[View_ReportSummary].[EmailID], [View_ReportSummary].[SuccessfullySent], [View_ReportSummary].[Bounced], [View_ReportSummary].[Opened], [View_ReportSummary].[EmailAddress], [View_ReportSummary].[CustomerID], [View_ReportSummary].[Clicked], [View_ReportSummary].[NoTimesOpened], [View_ReportSummary].[OpenedDateTime]) AS View_ReportSummary LEFT OUTER JOIN View_CUSTOMFIELDS ON View_ReportSummary.EmailID = View_CUSTOMFIELDS.EmailAddressID WHERE Row > "&nCurrentBlock&" AND Row <= "&(nCurrentBlock+3000)&" AND (View_CUSTOMFIELDS.CustomerID = "&nCID&" OR View_CUSTOMFIELDS.CustomerID IS NULL) ORDER BY ROW, View_CUSTOMFIELDS.CustomFieldName"
	
	nCurrentBlock = nCurrentBlock + 3000
	MaxLoops = MaxLoops + 1
			
Loop

strSQL = strSQL & " FOR XML AUTO, ELEMENTS"

Response.Write strSQL
%>