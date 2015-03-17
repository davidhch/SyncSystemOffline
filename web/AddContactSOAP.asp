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
	
	Sub ReturnSub(bReportError)
	
		'strMessage = strMessage & "<br>Before Function" 	
			
		returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		returnXML = returnXML & "<SOAP:Body>"
		returnXML = returnXML & "<MessageReturn>"
		returnXML = returnXML & "<Time>"&Now()&"</Time>"
		returnXML = returnXML & "<Message>"&bAction&"</Message>"
		returnXML = returnXML & "<MessageResponse>"&strMessage&"</MessageResponse>"
		returnXML = returnXML & "<NewZappCustomerID>"&nNewZappCustomerID&"</NewZappCustomerID>"
		returnXML = returnXML & "<OrderEntryCustomerID>"&nOrderEntryCustomerID&"</OrderEntryCustomerID>"
		returnXML = returnXML & "</MessageReturn>"
		returnXML = returnXML & "</SOAP:Body>"
		returnXML = returnXML & "</SOAP:Envelope>"

		Response.Write (returnXML) 
		
	End Sub	
	
	On Error Resume Next

	Response.Write "<br>Test Saving"
		
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	'strMessage = strMessage & "****request: <br>"
	
	'strMessage = strMessage & "<br><br>tgtfyduituiguifgBefore variables<br><br>"
	
	'If(TypeName(objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FirstName")) = "Nothing") Then 
	'	strMessage = strMessage & "<br><br>****FALSE<br><br>"
	'Else
	'	strMessage = strMessage & "<br><br>&&&&TRue<br><br>"
	'End If
	
	nID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ID").Text
	nCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CID").Text
	nNewZappCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappCID").Text
	strEmailAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/EmailAddress").Text
	strFirstName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FirstName").Text
	strLastName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/LastName").Text
	strJobTitle = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/JobTitle").Text
	strMobileNo = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/MobileNo").Text 
	bMainContact = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/MainContact").Text 
	bAccountsContact = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AccountsContact").Text
	strExtension = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Extension").Text
	strUsername = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Username").Text
	strPassword = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Password").Text
	strWebSite = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/WebSite").Text
	bNewsletter = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Newsletter").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text
	
	'strMessage = strMessage & "<br>nCID: " & nCID & "<br>"
	'strMessage = strMessage & "nNewZappCID: " & nNewZappCID & "<br>"
	'strMessage = strMessage & "strEmailAddress: " & strEmailAddress & "<br>"
	'strMessage = strMessage & "strFirstName: " & strFirstName & "<br>"
	'strMessage = strMessage & "strLastName: " & strLastName & "<br>"
	'strMessage = strMessage & "strJobTitle: " & strJobTitle & "<br>"
	'strMessage = strMessage & "strMobileNo: " & strMobileNo & "<br>"
	'strMessage = strMessage & "bMainContact: " & bMainContact & "<br>"
	'strMessage = strMessage & "bAccountsContact: " & bAccountsContact & "<br>"
	'strMessage = strMessage & "strExtension: " & strExtension & "<br>"
	'strMessage = strMessage & "strUsername: " & strUsername & "<br>"
	'strMessage = strMessage & "strPassword: " & strPassword & "<br>"
	'strMessage = strMessage & "strWebSite: " & strWebSite & "<br>"
	'strMessage = strMessage & "bNewsletter: " & bNewsletter & "<br>"
	'strMessage = strMessage & "nOnlineID: " & nOnlineID & "<br>"
	'strMessage = strMessage & "nOfflineID: " & nOfflineID & "<br>"
	'strMessage = strMessage & "nTableID: " & nTableID & "<br>"
	'strMessage = strMessage & "bLocked: " & bLocked & "<br>"
	'strMessage = strMessage & "bExecuted: " & bExecuted & "<br>"

	nNewZappCustomerID = nNewZappCID
	nOrderEntryCustomerID = nCID
	
	bDebug = True
	
	strDB = strOEDB
	
	Set objAddContact = Nothing
	Set objAddContact = New CDBInterface
	objAddContact.Init arrCustDetails, nTotCustDetails
	
	objAddContact.NewSQL
	objAddContact.AddFieldToQuery CONTACTS_ID
	objAddContact.AddWhere CONTACTS_ID, "=", nID, "number", "and"
	objAddContact.QueryDatabase
	
	nRecords = objAddContact.GetNumberOfRecords()
	
	Response.Write "nRecords: " & nRecords 
	
	If nRecords = 0 Then
		bAddNewContact = True
	Else
		bUpdateCurrentContact = True
	End If
	
	If bAddNewContact Then
		'strMessage = strMessage & "Adding new contact" & "<br>"
		
		bAction = True
				
		objAddContact.NewSQL
		objAddContact.StoreValue CONTACTS_CUSTOMERID, nCID
		objAddContact.StoreValue CONTACTS_EMAILADDRESS, strEmailAddress
		objAddContact.StoreValue CONTACTS_FIRSTNAME, strFirstName
		objAddContact.StoreValue CONTACTS_LASTNAME, strLastName
		objAddContact.StoreValue CONTACTS_JOBTITLE, strJobTitle
		objAddContact.StoreValue CONTACTS_MOBILENO, strMobileNo
		objAddContact.StoreValue CONTACTS_MAINCONTACT, bMainContact
		objAddContact.StoreValue CONTACTS_ACCOUNTCONTACT, bAccountsContact
		objAddContact.StoreValue CONTACTS_EXTENSION, strExtension
		objAddContact.StoreValue CONTACTS_USERNAME, strUsername
		objAddContact.StoreValue CONTACTS_PASSWORD, strPassword
		objAddContact.StoreValue CONTACTS_WEBSITE, strWebSite
		objAddContact.StoreValue CONTACTS_NEWSLETTER, bNewsletter
		objAddContact.StoreValue CONTACTS_NEWZAPPCID, nNewZappCID
		objAddContact.InsertUsingStoredValues "Contacts"
		
	End If
	
	If bUpdateCurrentContact Then
		'strMessage = strMessage & "Updating contact" & "<br>"
		
		nID = objAddContact.GetValue(CONTACTS_ID) 
		
		bAction = True
		
		If Len(nCID) > 0 Then
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_CUSTOMERID, nCID
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(nNewZappCID) > 0 Then
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_NEWZAPPCID, nNewZappCID
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strEmailAddress) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_EMAILADDRESS, strEmailAddress
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strFirstName) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_FIRSTNAME, strFirstName
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strLastName) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_LASTNAME, strLastName
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strJobTitle) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_JOBTITLE, strJobTitle
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strMobileNo) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_MOBILENO, strMobileNo
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strExtension) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_EXTENSION, strExtension
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strUsername) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_USERNAME, strUsername
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strPassword) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_PASSWORD, strPassword
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(strWebSite) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_WEBSITE, strWebSite
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		If Len(bNewsletter) > 0 Then	
			objAddContact.NewSQL
			objAddContact.StoreValue CONTACTS_NEWLETTER, bNewsletter
			objAddContact.SelectRecordForUpdate CONTACTS_ID, nID
			objAddContact.UpdateUsingStoredValues "Contacts"	
		End If
		
	End If
	
	If Len(nOnlineID)>0 Then
		bLocked = 0
		bExecuted = 1
		Call CreateOfflineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	End If
	
	Function CreateOfflineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
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
		
		Call CompleteSyncProcess(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
	End Function
	
	Function CompleteSyncProcess(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)

		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)
	
		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<OnlineID>"&nOnlineID&"</OnlineID>"
		sendReq = sendReq & "<OfflineID>"&nOfflineID&"</OfflineID>"
		sendReq = sendReq & "<TableID>"&nTableID&"</TableID>"
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