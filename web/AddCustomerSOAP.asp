<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->
<%
	Response.Buffer = True 
	
	bAction = False
	
	Sub ReturnSub( bReportError )
	
		returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		returnXML = returnXML & "<SOAP:Body>"
		returnXML = returnXML & "<MessageReturn>"
		returnXML = returnXML & "<Time>"&Now()&"</Time>"
		returnXML = returnXML & "<Message>"&bReportError&"</Message>"
		returnXML = returnXML & "<MessageResponse>"&strMessage&"</MessageResponse>"
		returnXML = returnXML & "<NewZappCustomerID>"&nNewZappCustomerID&"</NewZappCustomerID>"
		returnXML = returnXML & "<OrderEntryCustomerID>"&nOrderEntryCustomerID&"</OrderEntryCustomerID>"
		returnXML = returnXML & "</MessageReturn>"
		returnXML = returnXML & "</SOAP:Body>"
		returnXML = returnXML & "</SOAP:Envelope>"

		Response.Write (returnXML) 
		
	End Sub	
	
	On Error Resume Next
	
	strMessage = "********Getting to saving page!"
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load( Request )
	
	nCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nNewZappCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappCID").Text
	strCompanyName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CompanyName").Text
	strBillingAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/BillingAddress").Text
	strCity = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/City").Text
	strStateOrProvince = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/StateOrProvince").Text
	strPostalCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PostalCode").Text 
	nCountryID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CountryID").Text 
	strContactTitle = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ContactTitle").Text
	strPhoneNumber = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PhoneNumber").Text
	strFaxNumber = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FaxNumber").Text
	strEmailAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/EmailAddress").Text
	strWebAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/WebAddress").Text
	strNotes = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Notes").Text
	strWebSiteUserName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/WebSiteUserName").Text
	strWebSitePassword = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/WebSitePassword").Text
	bLeadGenerator = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/LeadGenerator").Text
	nLeadReferralID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/LeadReferralID").Text
	nVATCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/VATCode").Text
	strVATNo = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/VATNo").Text
	bReseller = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Reseller").Text
	bShowCustomer = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TimeStamp").Text
	bCreditCardCustomer = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CreditCardCustomer").Text
	nCreditLimit = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CreditLimit").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text
	
	strMessage = strMessage & "********nCustomerID: " & nCustomerID

	strMessage = strMessage & "********nNewZappCustomerID: " & nNewZappCID
	strMessage = strMessage & "********nOrderEntryCustomerID: " & nCustomerID
	
	'bDebug = True
	
	strDB = strOEDB
	
	If Len(nOnlineID) = 0 Then
		nOnlineID = nNewZappCID
	End If
	
	'This should only return 1 record
	Set objCheckForID = Nothing
	Set objCheckForID = New CDBInterface
	objCheckForID.Init arrCustDetails, nTotCustDetails
		
	objCheckForID.NewSQL
	objCheckForID.AddFieldToQuery SYNCLOOKUP_OFFLINEID
	objCheckForID.AddWhere SYNCLOOKUP_ONLINEID, "=", nNewZappCID, "number", "and"
	objCheckForID.AddWhere SYNCLOOKUP_TABLEID, "=", 7, "number", "and"
	objCheckForID.AddWhere SYNCLOOKUP_LOCKED, "=", 0, "number", "and"
	objCheckForID.AddWhere SYNCLOOKUP_EXECUTED, "=", 1, "number", "and"		
	objCheckForID.QueryDatabase
	
	nRecords = objCheckForID.GetNumberOfRecords()
		
	If nRecords > 0 Then
		bFoundIDRelationship = True
		nCustomerID = objCheckForID.GetValue(SYNCLOOKUP_OFFLINEID)
		nRecords = 0
	End If		
	
	strMessage = strMessage & "********POST SYNC LOOKUP CHECK"
	strMessage = strMessage & "********nCustomerID: " & nCustomerID
	
	Dim objAddCustomer
	Set objAddCustomer = Nothing
	Set objAddCustomer = New CDBInterface
	objAddCustomer.Init arrCustDetails, nTotCustDetails
	
	objAddCustomer.NewSQL
	objAddCustomer.AddFieldToQuery CUSTOMERS_ID
	objAddCustomer.AddWhere CUSTOMERS_ID, "=", nCustomerID, "number", "and"
	objAddCustomer.QueryDatabase
	
	nRecords = objAddCustomer.GetNumberOfRecords()
		
	strMessage = strMessage & "********nRecords:" & nRecords
	
	If nRecords = 0 Then
		strMessage = strMessage & "********NO CUSTOMER FOUND, NEW CUSTOMER"
		bAddNewCustomer = True
	Else
		strMessage = strMessage & "********CUSTOMER FOUND, CUSTOMER UPDATED"
		bUpdateCurrentCustomer = True
	End If
	
	If bAddNewCustomer Then
		
		bAction = True
				
		objAddCustomer.StoreValue CUSTOMERS_COMPANYNAME, strCompanyName
		objAddCustomer.StoreValue CUSTOMERS_BILLINGADDRESS, strBillingAddress
		objAddCustomer.StoreValue CUSTOMERS_CITY, strCity
		objAddCustomer.StoreValue CUSTOMERS_STATEORPROVINCE, strStateOrProvince
		objAddCustomer.StoreValue CUSTOMERS_POSTALCODE, strPostalCode
		objAddCustomer.StoreValue CUSTOMERS_COUNTRYID, nCountryID
		objAddCustomer.StoreValue CUSTOMERS_CONTACTTITLE, strContactTitle
		objAddCustomer.StoreValue CUSTOMERS_PHONENUMBER, strPhoneNumber
		objAddCustomer.StoreValue CUSTOMERS_FAXNUMBER, strFaxNumber
		objAddCustomer.StoreValue CUSTOMERS_EMAILADDRESS, strEmailAddress
		objAddCustomer.StoreValue CUSTOMERS_WEBADDRESS, strWebAddress
		objAddCustomer.StoreValue CUSTOMERS_NOTES, strNotes
		objAddCustomer.StoreValue CUSTOMERS_WEBSITEUSERNAME, strWebSiteUserName
		objAddCustomer.StoreValue CUSTOMERS_WEBSITEPASSWORD, strWebSitePassword
		If Len(bLeadGenerator) > 0 Then
			objAddCustomer.StoreValue CUSTOMERS_LEADGENERATOR, bLeadGenerator
		Else
			objAddCustomer.StoreValue CUSTOMERS_LEADGENERATOR, 0
		End If
		objAddCustomer.StoreValue CUSTOMERS_LEADREFERRALID, nLeadReferralID
		objAddCustomer.StoreValue CUSTOMERS_VATCODE, nVATCode
		objAddCustomer.StoreValue CUSTOMERS_VATNO, strVATNo


		' ************************************************
		' TCM - Commented Out Reseller 03/01/12
		' Flagging all online orders as a reseller customer.  Undesired effect for OE.
		' Please speak to TCM before commenting back in
		' ************************************************
		'If Len(bReseller) > 0 Then
		'	objAddCustomer.StoreValue CUSTOMERS_RESELLER, bReseller
		'Else
			objAddCustomer.StoreValue CUSTOMERS_RESELLER, 0
		'End If
		' ************************************************


		objAddCustomer.StoreValue CUSTOMERS_SHOWCUSTOMER, bShowCustomer
		objAddCustomer.StoreValue CUSTOMERS_CREDITCARDCUSTOMER, bCreditCardCustomer
		objAddCustomer.StoreValue CUSTOMERS_CREDITLIMIT, nCreditLimit
		objAddCustomer.StoreValue CUSTOMERS_NEWZAPPCID, nNewZappCID
		objAddCustomer.InsertUsingStoredValues "Customers"
		
		bNewCustomer = True
			Call GetNewCustomerCID()
	End If
	
	strMessage = strMessage & "********Before Update!!!!!"
	strMessage = strMessage & "********nCustomerID: " & nCustomerID
	strMessage = strMessage & "********nNewZappCID: " & nNewZappCID
	
	If bUpdateCurrentCustomer Then
	
		strMessage = strMessage & "********Inside Update!!!!!"
		strMessage = strMessage & "********nCustomerID: " & nCustomerID
		strMessage = strMessage & "********nNewZappCID: " & nNewZappCID
		
		If Len(strCompanyName) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_COMPANYNAME, strCompanyName
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strBillingAddress) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_BILLINGADDRESS, strBillingAddress
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strCity) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_CITY, strCity
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strStateOrProvince) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_STATEORPROVINCE, strStateOrProvince
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strPostalCode) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_POSTALCODE, strPostalCode
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(nCountryID) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_COUNTRYID, nCountryID
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strContactTitle) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_CONTACTTITLE, strContactTitle
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strPhoneNumber) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_PHONENUMBER, strPhoneNumber
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strFaxNumber) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_FAXNUMBER, strFaxNumber
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strEmailAddress) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_EMAILADDRESS, strEmailAddress
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strWebAddress) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_WEBADDRESS, strWebAddress
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strNotes) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_NOTES, strNotes
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strWebSiteUserName) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_WEBSITEUSERNAME, strWebSiteUserName
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strWebSitePassword) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_WEBSITEPASSWORD, strWebSitePassword
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(bLeadGenerator) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_LEADGENERATOR, bLeadGenerator
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(nLeadReferralID) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_LEADREFERRALID, nLeadReferralID
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(nVATCode) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_VATCODE, nVATCode
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(strVATNo) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_VATNO, strVATNo
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(bReseller) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_RESELLER, bReseller
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(bShowCustomer) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_SHOWCUSTOMER, bShowCustomer
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(bCreditCardCustomer) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_CREDITCARDCUSTOMER, bCreditCardCustomer
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(nCreditLimit) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_CREDITLIMIT, nCreditLimit
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If
		If Len(nNewZappCID) > 0 Then	
			bAction = True
			objAddCustomer.NewSQL
			objAddCustomer.StoreValue CUSTOMERS_NEWZAPPCID, nNewZappCID
			objAddCustomer.SelectRecordForUpdate CUSTOMERS_ID, nCustomerID
			objAddCustomer.UpdateUsingStoredValues "Customers"	
		End If

	End If
	
	Function GetNewCustomerCID()
	
		Dim objAddCustomer
		Set objAddCustomer = Nothing
		Set objAddCustomer = New CDBInterface
		objAddCustomer.Init arrCustDetails, nTotCustDetails
	
		objAddCustomer.NewSQL
		objAddCustomer.AddFieldToQuery CUSTOMERS_ID
		objAddCustomer.AddFieldToQuery CUSTOMERS_COMPANYNAME
		objAddCustomer.AddWhere CUSTOMERS_COMPANYNAME, "=", strCompanyName, "text", "and"
		objAddCustomer.QueryDatabase
	
		nRecords = objAddCustomer.GetNumberOfRecords()
		
		If nRecords > 0 Then
			nCustomerID = objAddCustomer.GetValue(CUSTOMERS_ID)
		End If
	
	End Function
	
	If Len(nOnlineID)>0 Then
		bLocked = 0
		bExecuted = 1
		nOfflineID = nCustomerID
		Call CreateOnlineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	End If
	
	Function CreateOnlineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Set objAddSyncLookup = Nothing
		Set objAddSyncLookup = New CDBInterface
		objAddSyncLookup.Init arrCustDetails, nTotCustDetails
		
		objAddSyncLookup.NewSQL
		objAddSyncLookup.AddFieldToQuery SYNCLOOKUP_ID
		objAddSyncLookup.AddWhere SYNCLOOKUP_ONLINEID, "=", nOnlineID, "number", "and"
		objAddSyncLookup.AddWhere SYNCLOOKUP_TABLEID, "=", nTableID, "number", "and"
		objAddSyncLookup.QueryDatabase
	
		nRecords = objAddSyncLookup.GetNumberOfRecords()
			
		If nRecords = 0 Then
			bAddNewSyncLookup = True
		Else
			bUpdateSyncLookup = True
			nSyncLookupID = objAddSyncLookup.GetValue(SYNCLOOKUP_ID)
		End If
	
		If bAddNewSyncLookup Then
			objAddSyncLookup.NewSQL         
			objAddSyncLookup.StoreValue SYNCLOOKUP_ONLINEID, nOnlineID
			objAddSyncLookup.StoreValue SYNCLOOKUP_OFFLINEID, nOfflineID
			objAddSyncLookup.StoreValue SYNCLOOKUP_TABLEID, nTableID
			objAddSyncLookup.StoreValue SYNCLOOKUP_LOCKED, bLocked
			objAddSyncLookup.StoreValue SYNCLOOKUP_EXECUTED, bExecuted
			objAddSyncLookup.InsertUsingStoredValues "SyncLookUp"
		End If
		
		If bUpdateSyncLookup Then
			If Len(nOnlineID) > 0 Then
				objAddSyncLookup.NewSQL
				objAddSyncLookup.StoreValue SYNCLOOKUP_ONLINEID, nOnlineID
				objAddSyncLookup.SelectRecordForUpdate SYNCLOOKUP_ID, nSyncLookupID
				objAddSyncLookup.UpdateUsingStoredValues "SyncLookUp"	
			End If
			If Len(nOfflineID) > 0 Then
				objAddSyncLookup.NewSQL
				objAddSyncLookup.StoreValue SYNCLOOKUP_OFFLINEID, nOfflineID
				objAddSyncLookup.SelectRecordForUpdate SYNCLOOKUP_ID, nSyncLookupID
				objAddSyncLookup.UpdateUsingStoredValues "SyncLookUp"	
			End If
			If Len(bLocked) > 0 Then	
				objAddSyncLookup.NewSQL
				objAddSyncLookup.StoreValue SYNCLOOKUP_LOCKED, bLocked
				objAddSyncLookup.SelectRecordForUpdate SYNCLOOKUP_ID, nSyncLookupID
				objAddSyncLookup.UpdateUsingStoredValues "SyncLookUp"	
			End If
			If Len(bExecuted) > 0 Then	
				objAddSyncLookup.NewSQL
				objAddSyncLookup.StoreValue SYNCLOOKUP_EXECUTED, bExecuted
				objAddSyncLookup.SelectRecordForUpdate SYNCLOOKUP_ID, nSyncLookupID
				objAddSyncLookup.UpdateUsingStoredValues "SyncLookUp"	
			End If
		End If
		
		Call CompleteSyncProcess()
	
	End Function
	
	Function CompleteSyncProcess()

		If nTableID = 7 Then
			nOnlineTableID = 6
		End If

		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)
	
		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<OnlineID>"&nOnlineID&"</OnlineID>"
		sendReq = sendReq & "<OfflineID>"&nOfflineID&"</OfflineID>"
		sendReq = sendReq & "<TableID>"&nOnlineTableID&"</TableID>"
		sendReq = sendReq & "<Locked>"&bLocked&"</Locked>"
		sendReq = sendReq & "<Executed>"&bExecuted&"</Executed>"
		sendReq = sendReq & "</Request>"
		sendReq = sendReq & "</SOAP:Body>"
		sendReq = sendReq & "</SOAP:Envelope>"	

		objxmlhttp.open "POST", SoapServerURL, False
		objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
		objxmlhttp.setRequestHeader "MessageType", "CALL"
		objxmlhttp.setRequestHeader "Content-Type", "text/xml"
		objxmlhttp.send(sendReq)

		Response.write objxmlhttp.responseText

		Set objxmlhttp = Nothing
		Set objxmldom = Nothing
	
	End Function
	
	If bAction Then 
		ReturnSub "True"
	Else 
		ReturnSub "False"
	End If
	
	Set objxmldom = Nothing
	Set objxmlhttp = Nothing
%>