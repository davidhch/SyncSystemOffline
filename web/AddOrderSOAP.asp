<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->
<%
	Response.Buffer = True 
	
	Dim nOrderID
	nOrderID=0
	
	bAction = False
	
	Sub ReturnSub(bReportError)
	
		returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		returnXML = returnXML & "<SOAP:Body>"
		returnXML = returnXML & "<MessageReturn>"
		returnXML = returnXML & "<Version>DHH-16012007-1</Version>"
		returnXML = returnXML & "<Time>"&Now()&"</Time>"
		returnXML = returnXML & "<Message>"&bAction&"</Message>"
		returnXML = returnXML & "<MessageResponse>"&strMessage&"</MessageResponse>"
		returnXML = returnXML & "<InvoiceID>"&nOrderID&"</InvoiceID>"
		returnXML = returnXML & "<NewZappCustomerID>"&nNewZappCustomerID&"</NewZappCustomerID>"
		returnXML = returnXML & "<OrderEntryCustomerID>"&nOrderEntryCustomerID&"</OrderEntryCustomerID>"
		returnXML = returnXML & "</MessageReturn>"
		returnXML = returnXML & "</SOAP:Body>"
		returnXML = returnXML & "</SOAP:Envelope>"

		Response.Write (returnXML) 
		
	End Sub	
	
	Sub GetItemFromSyncLookup(nNonLocalID, nTable, nCurrentValue)	
		If Len(nCurrentValue)=0 Then
			Dim  objCheckForID
			Set objCheckForID = New CDBInterface
			objCheckForID.Init arrCustDetails, nTotCustDetails
				
			objCheckForID.NewSQL
			objCheckForID.AddFieldToQuery SYNCLOOKUP_OFFLINEID
			objCheckForID.AddWhere SYNCLOOKUP_ONLINEID, "=", nNonLocalID, "number", "and"
			objCheckForID.AddWhere SYNCLOOKUP_TABLEID, "=", nTable, "number", "and"
			objCheckForID.AddWhere SYNCLOOKUP_LOCKED, "=", 0, "number", "and"
			objCheckForID.AddWhere SYNCLOOKUP_EXECUTED, "=", 1, "number", "and"		
			objCheckForID.QueryDatabase	
			nRecords = objCheckForID.GetNumberOfRecords()
				
			If nRecords > 0 Then
				If objCheckForID.GetValue(SYNCLOOKUP_OFFLINEID)>0 Then
					nCurrentValue = objCheckForID.GetValue(SYNCLOOKUP_OFFLINEID)
				End If
			End If			
		End If	
	End Sub
		
	
	'On Error Resume Next
		
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	'nID = 0
	'nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
	nCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nOrderEntryCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryCID").Text
	nQuotationID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/QuotationID").Text
	nEmployeeID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/EmployeeID").Text
	dOrderDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderDate").Text
	strPurchaseOrderNumber = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PurchaseOrderNumber").Text
	strInvoiceDescription = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/InvoiceDescription").Text
	strShipName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipName").Text 
	strShipAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipAddress").Text 
	strShipCity = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipCity").Text
	strShipStateOrProvince = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipStateOrProvince").Text
	strShipPostalCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipPostalCode").Text
	strShipCountry = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipCountry").Text
	strShipPhoneNumber = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipPhoneNumber").Text
	dShipDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShipDate").Text
	nShippingMethodID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ShippingMethodID").Text
	nFreightCharge = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FreightCharge").Text
	nSalesTaxRate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SalesTaxRate").Text
	nLeadReferralID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/LeadReferralID").Text
	strProposalReference = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProposalReference").Text
	dEntryDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/EntryDate").Text
	bCommissionPaid = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CommissionPaid").Text
	nInvoiceID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/InvoiceID").Text
	'strUniqueKey = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UniqueKey").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text
	
	strDB = strOEDB
	
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
		objGetOEID.AddFieldToQuery ORDERS_ID
		objGetOEID.AddWhere ORDERS_UNIQUEKEY, "LIKE", strUniqueOEID, "text", "and"
		objGetOEID.QueryDatabase
	
		nRecords = objGetOEID.GetNumberOfRecords()
		
		If nRecords > 0 Then
			getOEID = objGetOEID.GetValue(ORDERS_ID)
		End If
	
	End Function
	
	'bDebug = True
	nRecords = 0
	
	Dim objAddOrder
	Set objAddOrder = New CDBInterface
	objAddOrder.Init arrCustDetails, nTotCustDetails
	
	If Len(nInvoiceID)=0 Then
		GetItemFromSyncLookup nOnlineID, nTableID, nInvoiceID
	End If
	
	If Len(nInvoiceID)>0 Then	
		objAddOrder.NewSQL
		objAddOrder.AddFieldToQuery ORDERS_ID
		objAddOrder.AddWhere ORDERS_ID, "=", nInvoiceID, "number", "and"
		objAddOrder.QueryDatabase
	
		nRecords = objAddOrder.GetNumberOfRecords()
	End If
	
	If nRecords = 0 Then
		bAddNewOrder = True
	Else
		bUpdateOrder = True
	End If
	
	strUniqueOEID = createUniqueID()
	
	If 	bAddNewOrder Then	
		
		bAction = True
			
		objAddOrder.NewSQL		
		objAddOrder.StoreValue ORDERS_CUSTOMERID, nOrderEntryCID
		objAddOrder.StoreValue ORDERS_QUOTATIONID, nQuotationID
		objAddOrder.StoreValue ORDERS_EMPLOYEEID, nEmployeeID
		If Len(dOrderDate) > 0 Then
			objAddOrder.StoreValue ORDERS_ORDERDATE, dOrderDate
		Else
			objAddOrder.StoreValue ORDERS_ORDERDATE, Date()
		End If
		objAddOrder.StoreValue ORDERS_PURCHASEORDERNUMBER, strPurchaseOrderNumber
		objAddOrder.StoreValue ORDERS_INVOICEDESCRIPTION, strInvoiceDescription
		objAddOrder.StoreValue ORDERS_SHIPNAME, strShipName
		objAddOrder.StoreValue ORDERS_SHIPADDRESS, strShipAddress
		objAddOrder.StoreValue ORDERS_SHIPCITY, strShipCity
		objAddOrder.StoreValue ORDERS_SHIPSTATEORPROVINCE, strShipStateOrProvince
		objAddOrder.StoreValue ORDERS_SHIPPOSTCODE, strShipPostalCode
		objAddOrder.StoreValue ORDERS_SHIPCOUNTRY, strShipCountry
		objAddOrder.StoreValue ORDERS_SHIPPHONENUMBER, strShipPhoneNumber
		objAddOrder.StoreValue ORDERS_SHIPDATE, dShipDate
		objAddOrder.StoreValue ORDERS_SHIPPINGMETHODID, nShippingMethodID 
		objAddOrder.StoreValue ORDERS_FREIGHTCHARGE, nFreightCharge
		objAddOrder.StoreValue ORDERS_SALESTAXRATE, nSalesTaxRate
		If Len(nLeadReferralID) > 0 Then
			objAddOrder.StoreValue ORDERS_LEADREFERRALID, nLeadReferralID
		End If
		objAddOrder.StoreValue ORDERS_PROPOSALREFERENCE, strProposalReference
		If Len(dEntryDate) > 0 Then
			objAddOrder.StoreValue ORDERS_ENTRYDATE, dEntryDate
		Else
			objAddOrder.StoreValue ORDERS_ENTRYDATE, Now()
		End If
		objAddOrder.StoreValue ORDERS_COMMISSIONPAID, bCommissionPaid
		objAddOrder.StoreValue ORDERS_INVOICEID, nInvoiceID
		objAddOrder.StoreValue ORDERS_UNIQUEKEY, strUniqueOEID
		objAddOrder.InsertUsingStoredValues "Orders"
		
		nOEID = getOEID(strUniqueOEID)
		
	End If
	
	If bUpdateOrder Then
	
		bAction = True
		
		If Len(nOrderEntryCID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_CUSTOMERID, nOrderEntryCID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nQuotationID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_QUOTATIONID, nQuotationID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nEmployeeID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_EMPLOYEEID, nEmployeeID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(dOrderDate) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_ORDERDATE, dOrderDate
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strPurchaseOrderNumber) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_PURCHASEORDERNUMBER, strPurchaseOrderNumber
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strInvoiceDescription) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_INVOICEDESCRIPTION, strInvoiceDescription
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipName) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPNAME, strShipName
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipAddress) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPADDRESS, strShipAddress
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipCity) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPCITY, strShipCity
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipStateOrProvince) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPSTATEORPROVINCE, strShipStateOrProvince
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipPostalCode) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPPOSTCODE, strShipPostalCode
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipCountry) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPCOUNTRY, strShipCountry
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strShipPhoneNumber) > 0 Then
			objAddOrderDetails.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPPHONENUMBER, strShipPhoneNumber
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(dShipDate) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPDATE, dShipDate
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nShippingMethodID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SHIPPINGMETHODID, nShippingMethodID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nFreightCharge) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_FREIGHTCHARGE, nFreightCharge
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nSalesTaxRate) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_SALESTAXRATE, nSalesTaxRate
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nLeadReferralID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_LEADREFERRALID, nLeadReferralID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(strProposalReference) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_PROPOSALREFERENCE, strProposalReference
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(dEntryDate) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_ENTRYDATE, dEntryDate
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(bCommissionPaid) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_COMMISSIONPAID, bCommissionPaid
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
		If Len(nInvoiceID) > 0 Then
			objAddOrder.NewSQL
			objAddOrder.StoreValue ORDERS_INVOICEID, nInvoiceID
			objAddOrder.SelectRecordForUpdate ORDERS_ID, nInvoiceID
			objAddOrder.UpdateUsingStoredValues "Orders"	
		End If
	
		nOEID = getOEID(strUniqueOEID)
		
	End If
	
	If Len(nOnlineID)>0 Then
		nOfflineID = nOEID
		nOrderID = nOfflineID
		
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