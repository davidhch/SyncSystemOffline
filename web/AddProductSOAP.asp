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

	nProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductID").Text
	nOnlineProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineProductID").Text
	strProductName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductName").Text
	nProductCategoryID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductCategoryID").Text
	nUnitPrice = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UnitPrice").Text
	nUnitCost = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UnitCost").Text
	strOneOffCost = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OneOffCost").Text
	nSupplierID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SupplierID").Text
	strDescription = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Description").Text 
	strSystemDescription = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SystemDescription").Text 
	nSetupPrice = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SetupPrice").Text
	nOrder = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Order").Text
	bActive = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Active").Text
	nOrderEntryProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryProductID").Text
	nNewZappCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/nNewZappCustomerID").Text
	nOrderEntryCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/nOrderEntryCustomerID").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text
	
	'nNewZappCustomerID = nNewZappCID
	'nOrderEntryCustomerID = nCID
	
	Response.Write "<br>nOrderEntryProductID: " & nOrderEntryProductID
	
	bDebug = True
	strDB = strOEDB	
	
	Set objAddProduct = Nothing
	Set objAddProduct = New CDBInterface
	objAddProduct.Init arrCustDetails, nTotCustDetails
	
	objAddProduct.NewSQL
	objAddProduct.AddFieldToQuery PRODUCTS_ID
	objAddProduct.AddWhere PRODUCTS_ID, "=", nOrderEntryProductID, "number", "and"
	objAddProduct.QueryDatabase
	
	nRecords = objAddProduct.GetNumberOfRecords()
	
	Response.Write "<br>*********************************nRecords: " & nRecords
	
	If nRecords = 0 Then
		bAddNewProduct = True
	Else
		bUpdateCurrentProduct = True
	End If

	If 	bAddNewProduct Then	
		
		bAction = True
		
		objAddProduct.NewSQL			
		objAddProduct.StoreValue PRODUCTS_PRODUCTNAME, strProductName
		objAddProduct.StoreValue PRODUCTS_PRODUCTCATEGORYID, nProductCategoryID
		objAddProduct.StoreValue PRODUCTS_UNITPRICE, nUnitPrice
		objAddProduct.StoreValue PRODUCTS_UNITCOST, nUnitCost
		objAddProduct.StoreValue PRODUCTS_ONEOFFCOST, strOneOffCost
		objAddProduct.StoreValue PRODUCTS_SUPPLIERID, nSupplierID
		objAddProduct.StoreValue PRODUCTS_DESCRIPTION, strDescription
		objAddProduct.StoreValue PRODUCTS_SYSTEMDESCRIPTION, strSystemDescription
		objAddProduct.StoreValue PRODUCTS_SETUPPRICE, nSetupPrice
		objAddProduct.StoreValue PRODUCTS_ORDER, nOrder
		objAddProduct.StoreValue PRODUCTS_ACTIVE, bActive
		objAddProduct.InsertUsingStoredValues "Products"
		
		bNewProduct = True
			Call GetNewProductID()
	
	End If
	
	bUpdateCurrentProduct = False
	
	If bUpdateCurrentProduct Then
		
		bAction = True
	
		If Len(strProductName) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_PRODUCTNAME, strProductName
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nProductCategoryID) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_PRODUCTCATEGORYID, nProductCategoryID
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nUnitPrice) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_UNITPRICE, nUnitPrice
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nUnitCost) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_UNITCOST, nUnitCost
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(strOneOffCost) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_ONEOFFCOST, strOneOffCost
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nSupplierID) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_SUPPLIERID, nSupplierID
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(strDescription) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_DESCRIPTION, strDescription
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(strSystemDescription) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_SYSTEMDESCRIPTION, strSystemDescription
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nSetupPrice) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_SETUPPRICE, nSetupPrice
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(nOrder) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_ORDER, nOrder
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
		If Len(bActive) > 0 Then
			objAddProduct.NewSQL
			objAddProduct.StoreValue PRODUCTS_ACTIVE, bActive
			objAddProduct.SelectRecordForUpdate PRODUCTS_ID, nProductID
			objAddProduct.UpdateUsingStoredValues "Products"	
		End If
	End If
	
	Function GetNewProductID()
		
		Dim objGetProductID
		Set objGetProductID = Nothing
		Set objGetProductID = New CDBInterface
		objGetProductID.Init arrCustDetails, nTotCustDetails
	
		objGetProductID.NewSQL
		objGetProductID.AddFieldToQuery PRODUCT_ID
		objGetProductID.AddFieldToQuery PRODUCT_NAME
		objGetProductID.AddWhere PRODUCT_NAME, "=", strProductName, "text", "and"
		objGetProductID.QueryDatabase
	
		nRecords = objGetProductID.GetNumberOfRecords()
		
		If nRecords > 0 Then
			nOfflineID = objGetProductID.GetValue(PRODUCT_ID)
		End If
		
	End Function

	If Len(nOnlineID)>0 Then
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
		'strMessage = "Woooo Hoooo Something Has Changed"
		ReturnSub "True"
	Else 
		ReturnSub "False"
	End If
	
	Set objxmldom = Nothing
	Set objxmlhttp = Nothing
%>