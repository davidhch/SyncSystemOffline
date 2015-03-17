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
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	bDebug = True
	
	strDB = strOEDB
	
	nID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ID").Text
	nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
	nOrderEntryOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryOrderID").Text
	nProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductID").Text
	nOrderEntryProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryProductID").Text
	nQuantity = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Quantity").Text
	nUnitPrice = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UnitPrice").Text
	nDiscount = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Discount").Text
	nAccountType = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AccountType").Text
	nCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nNewZappCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappCID").Text
	dRenewalDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/RenewalDate").Text
	bStatistics = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Statistics").Text
	bFrontPage = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FrontPage").Text
	bServerSideIncludes = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ServerSideIncludes").Text
	bSSL = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SSL").Text
	bIndexServer = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/IndexServer").Text
	bCurrent = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Current").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text

	'Response.Write "<br>nOrderEntryProductID: " & nOrderEntryProductID

	Function createUniqueID()
		
		Set objTypeLib = Server.CreateObject("Scriptlet.TypeLib")
		strGUID = objTypeLib.Guid
		createUniqueID = Left( strGUID, Len( strGUID ) -2 )
		Set objTypeLib = Nothing
	
	End Function
	
	Function getOEID(strUniqueOEID)
	
		Set objGetOEID = Nothing
		Set objGetOEID = New CDBInterface
		objGetOEID.Init arrCustDetails, nTotCustDetails
	
		objGetOEID.NewSQL
		objGetOEID.AddFieldToQuery ORDERDETAILS_ID
		objGetOEID.AddWhere ORDERDETAILS_UNIQUEKEY, "LIKE", strUniqueOEID, "text", "and"
		objGetOEID.QueryDatabase
	
		nRecords = objGetOEID.GetNumberOfRecords()
		
		If nRecords > 0 Then
			getOEID = objGetOEID.GetValue(ORDERDETAILS_ID)
		End If
	
	End Function
	
	'bDebug = True
	
	Set objAddOrderDetails = Nothing
	Set objAddOrderDetails = New CDBInterface
	objAddOrderDetails.Init arrCustDetails, nTotCustDetails
	
	objAddOrderDetails.NewSQL
	objAddOrderDetails.AddFieldToQuery ORDERDETAILS_ID
	objAddOrderDetails.AddWhere ORDERDETAILS_ORDERID, "=", nOrderEntryOrderID, "number", "and"
	objAddOrderDetails.QueryDatabase
	
	nRecords = objAddOrderDetails.GetNumberOfRecords()
	nRecordID = objAddOrderDetails.GetValue(ORDERDETAILS_ID)
	
	If nRecords = 0 Then
		bAddNewOrderDetails = True
		'Response.Write "<br>bAddNewOrderDetails: " & bAddNewOrderDetails & "<br>"
	Else
		bUpdateOrderDetails = True
		'Response.Write "<br>bUpdateOrderDetails: " & bUpdateOrderDetails & "<br>"
	End If
	
	strUniqueOEID = createUniqueID()
	
	If bAddNewOrderDetails Then
	
		bAction = True
		
		'Response.Write "<br>Correct Place: " & " - " & nOrderEntryProductID & "<br>"
		
		objAddOrderDetails.NewSQL         
		objAddOrderDetails.StoreValue ORDERDETAILS_ORDERID, nOrderEntryOrderID
		objAddOrderDetails.StoreValue ORDERDETAILS_PRODUCTID, nOrderEntryProductID
		objAddOrderDetails.StoreValue ORDERDETAILS_QUANTITY, nQuantity
		objAddOrderDetails.StoreValue ORDERDETAILS_UNITPRICE, nUnitPrice
		objAddOrderDetails.StoreValue ORDERDETAILS_DISCOUNT, nDiscount
		objAddOrderDetails.StoreValue ORDERDETAILS_UNIQUEKEY, strUniqueOEID
		objAddOrderDetails.InsertUsingStoredValues "Order Details"
		
		nOEID = getOEID(strUniqueOEID)
		Response.Write "<br>Testing  nOEID: " & nOEID
		
		objAddOrderDetails.NewSQL         
		objAddOrderDetails.StoreValue WEBHOSTING_ACCOUNTTYPE, nAccountType
		objAddOrderDetails.StoreValue WEBHOSTING_CUSTOMERID, nCustomerID
		objAddOrderDetails.StoreValue WEBHOSTING_CID, nNewZappCID
		objAddOrderDetails.StoreValue WEBHOSTING_RENEWALDATE, dRenewalDate
		objAddOrderDetails.StoreValue WEBHOSTING_STATISTICS, bStatistics
		objAddOrderDetails.StoreValue WEBHOSTING_FRONTPAGE, bFrontPage
		objAddOrderDetails.StoreValue WEBHOSTING_SERVERSIDEINCLUDES, bServerSideIncludes
		objAddOrderDetails.StoreValue WEBHOSTING_SSL, bSSL
		objAddOrderDetails.StoreValue WEBHOSTING_INDEXSERVER, bIndexServer
		objAddOrderDetails.StoreValue WEBHOSTING_CURRENT, bCurrent
		objAddOrderDetails.InsertUsingStoredValues "Web Hosting"
	
	End If
	
	'Added to stop update of old records
	bUpdateOrderDetails = False
	
	If bUpdateOrderDetails Then
		
		bAction = True
			
		If Len(nOrderID) > 0 Then
			objAddOrderDetails.NewSQL
			objAddOrderDetails.StoreValue ORDERDETAILS_ORDERID, nOrderID
			objAddOrderDetails.SelectRecordForUpdate ORDERDETAILS_ID, nID
			objAddOrderDetails.UpdateUsingStoredValues "Order Details"	
		End If
		If Len(nProductID) > 0 Then
			objAddOrderDetails.NewSQL
			objAddOrderDetails.StoreValue ORDERDETAILS_PRODUCTID, nProductID
			objAddOrderDetails.SelectRecordForUpdate ORDERDETAILS_ID, nID
			objAddOrderDetails.UpdateUsingStoredValues "Order Details"	
		End If
		If Len(nQuantity) > 0 Then	
			objAddOrderDetails.NewSQL
			objAddOrderDetails.StoreValue ORDERDETAILS_QUANTITY, nQuantity
			objAddOrderDetails.SelectRecordForUpdate ORDERDETAILS_ID, nID
			objAddOrderDetails.UpdateUsingStoredValues "Order Details"	
		End If
		If Len(nUnitPrice) > 0 Then	
			objAddOrderDetails.NewSQL
			objAddOrderDetails.StoreValue ORDERDETAILS_UNITPRICE, nUnitPrice
			objAddOrderDetails.SelectRecordForUpdate ORDERDETAILS_ID, nID
			objAddOrderDetails.UpdateUsingStoredValues "Order Details"	
		End If
		If Len(nDiscount) > 0 Then	
			objAddOrderDetails.NewSQL
			objAddOrderDetails.StoreValue ORDERDETAILS_DISCOUNT, nDiscount
			objAddOrderDetails.SelectRecordForUpdate ORDERDETAILS_ID, nID
			objAddOrderDetails.UpdateUsingStoredValues "Order Details"	
		End If
		
	End If
	
	If Len(nOnlineID)>0 Then
		nOfflineID = nOEID

		bLocked = 0
		bExecuted = 1
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
