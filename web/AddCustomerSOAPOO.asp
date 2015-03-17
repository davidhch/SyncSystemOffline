<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibrary.asp"-->
<!--#include file="_private\CSyncOffline.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
	
	
	Response.Buffer = True 
	'bDebug = True
	
	strLine = ""
	strLine = strLine & "FILE STARTED AddCustomerSoapOO.asp"
	logSoap(strLine)
	
	'DB Connection
	strDB = strOEDB	
	
	Dim objSoapSync
	Set objSoapSync = New CSoap
	objSoapSync.Init	
	
	nCustomerID 				= objSoapSync.ReadSingleNode ("CustomerID")
	nNewZappCID 				= objSoapSync.ReadSingleNode ("NewZappCID")
	strCompanyName 				= objSoapSync.ReadSingleNode ("CompanyName")
	strBillingAddress 			= objSoapSync.ReadSingleNode ("BillingAddress")
	strCity 				= objSoapSync.ReadSingleNode ("City")
	strStateOrProvince 			= objSoapSync.ReadSingleNode ("StateOrProvince")
	strPostalCode 				= objSoapSync.ReadSingleNode ("PostalCode") 
	nCountryID 				= objSoapSync.ReadSingleNode ("CountryID") 
	strContactTitle 			= objSoapSync.ReadSingleNode ("ContactTitle")
	strPhoneNumber 				= objSoapSync.ReadSingleNode ("PhoneNumber")
	strFaxNumber 				= objSoapSync.ReadSingleNode ("FaxNumber")
	strEmailAddress 			= objSoapSync.ReadSingleNode ("EmailAddress")
	strWebAddress 				= objSoapSync.ReadSingleNode ("WebAddress")
	strNotes 				= objSoapSync.ReadSingleNode ("Notes")
	strWebSiteUserName 			= objSoapSync.ReadSingleNode ("WebSiteUserName")
	strWebSitePassword 			= objSoapSync.ReadSingleNode ("WebSitePassword")
	bLeadGenerator 				= objSoapSync.ReadSingleNode ("LeadGenerator")
	nLeadReferralID 			= objSoapSync.ReadSingleNode ("LeadReferralID")
	nVATCode 				= objSoapSync.ReadSingleNode ("VATCode")
	strVATNo 				= objSoapSync.ReadSingleNode ("VATNo")
	'bReseller 				= objSoapSync.ReadSingleNode ("Reseller")
	bReseller 				= 0
	bShowCustomer 				= objSoapSync.ReadSingleNode ("ShowCustomer")
	bCreditCardCustomer 			= objSoapSync.ReadSingleNode ("CreditCardCustomer")
	nCreditLimit 				= objSoapSync.ReadSingleNode ("CreditLimit")
	strCustomerFirstName			= objSoapSync.ReadSingleNode ("FirstName")
	strCustomerLastName			= objSoapSync.ReadSingleNode ("LastName")
	nOnlineID 				= objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID 				= objSoapSync.ReadSingleNode ("OfflineID")
	nTableID 				= objSoapSync.ReadSingleNode ("TableID")
	bLocked 				= objSoapSync.ReadSingleNode ("Locked")
	bExecuted 				= objSoapSync.ReadSingleNode ("Executed")
	
	strLine = ""
	strLine = strLine & "nCustomerID: " & nCustomerID
	logSoap(strLine)
	
	
	
	'This would run on an ASP page
	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline	 

	'...now do what you want to do!!!!
	
	objSyncOffline.addCustomer nCustomerID,nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit,strCustomerFirstName,strCustomerLastName,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted 
	
	Set objSyncOffline = Nothing
	
	Sub logSoap(strLine)
		dim strFilePath
		strFilePath = server.MapPath("logSoap.txt")
		
		Dim fso, f
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(strFilePath, 8, True)
		f.Write strLine & " - " & vbcrlf & now() & vbcrlf & "----------------" & vbcrlf & vbcrlf
	End Sub

%>