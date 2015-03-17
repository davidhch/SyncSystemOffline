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
	
	nPaymentsID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PaymentsID").Text
	'nOnlinePaymentID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlinePaymentID").Text
	nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
	nOrderEntryOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryOrderID").Text
	nPaymentAmount = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PaymentAmount").Text
	dPaymentDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PaymentDate").Text
	nCreditCardID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CreditCardID").Text
	nPaymentMethodID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PaymentMethodID").Text
	strTransactionCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TransactionCode").Text
	strProcessingValid = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProcessingValid").Text
	strProcessingCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProcessingCode").Text
	strProcessingCodeMessage = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProcessingCodeMessage").Text
	strAuthCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AuthCode").Text
	strAuthMessage = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AuthMessage").Text
	strResponseCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ResponseCode").Text
	strResponseCodeMessage = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ResponseCodeMessage").Text
	nAmountToCharge = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AmountToCharge").Text
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text

	'bDebug = True
	strDB = strOEDB
	
	Set objAddPayment = Nothing
	Set objAddPayment = New CDBInterface
	objAddPayment.Init arrCustDetails, nTotCustDetails
	
	objAddPayment.NewSQL
	objAddPayment.AddFieldToQuery PAYMENTS_ID
	objAddPayment.AddWhere PAYMENTS_ID, "=", nPaymentsID, "number", "and"
	objAddPayment.QueryDatabase
	
	nRecords = objAddPayment.GetNumberOfRecords()
	
	If nRecords = 0 Then
		bAddNewPayment = True
	Else
		bUpdateCurrentPayment = True
	End If

	If 	bAddNewPayment Then	
		
		bAction = True
		
		objAddPayment.NewSQL			
		objAddPayment.StoreValue PAYMENTS_ORDERID, nOrderEntryOrderID
		objAddPayment.StoreValue PAYMENTS_PAYMENTAMOUNT, nPaymentAmount
		objAddPayment.StoreValue PAYMENTS_PAYMENTDATE, dPaymentDate
		objAddPayment.StoreValue PAYMENTS_CREDITCARDID, nCreditCardID
		objAddPayment.StoreValue PAYMENTS_PAYMENTMETHODID, nPaymentMethodID
		objAddPayment.StoreValue PAYMENTS_TRANSACTIONCODE, strTransactionCode
		objAddPayment.StoreValue PAYMENTS_PROCESSINGVALID, strProcessingValid
		objAddPayment.StoreValue PAYMENTS_PROCESSINGCODE, strProcessingCode
		objAddPayment.StoreValue PAYMENTS_PROCESSINGCODEMESSAGE, strProcessingCodeMessage
		objAddPayment.StoreValue PAYMENTS_AUTHCODE, strAuthCode
		objAddPayment.StoreValue PAYMENTS_AUTHMESSAGE, strAuthMessage
		objAddPayment.StoreValue PAYMENTS_RESPONSECODE, strResponseCode
		objAddPayment.StoreValue PAYMENTS_RESPONSECODEMESSAGE, strResponseCodeMessage
		objAddPayment.StoreValue PAYMENTS_AMOUNTTOCHARGE, nAmountToCharge
		objAddPayment.InsertUsingStoredValues "Payments"
	
	End If
	
	'Added to stop update of old records
	bUpdateCurrentPayment = False
	
	If bUpdateCurrentPayment Then
		
		bAction = True
	
		If Len(nOrderEntryOrderID) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_ORDERID, nOrderEntryOrderID
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(nPaymentAmount) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PAYMENTAMOUNT, nPaymentAmount
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(dPaymentDate) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PAYMENTDATE, dPaymentDate
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(nCreditCardID) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_CREDITCARDID, nCreditCardID
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(nPaymentMethodID) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PAYMENTMETHODID, nPaymentMethodID
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strTransactionCode) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_TRANSACTIONCODE, strTransactionCode
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strProcessingValid) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PROCESSINGVALID, strProcessingValid
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strProcessingCode) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PROCESSINGCODE, strProcessingCode
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strProcessingCodeMessage) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_PROCESSINGCODEMESSAGE, strProcessingCodeMessage
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strAuthCode) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_AUTHCODE, strAuthCode
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strAuthMessage) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_AUTHMESSAGE, strAuthMessage
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strResponseCode) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_RESPONSECODE, strResponseCode
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(strResponseCodeMessage) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_RESPONSECODEMESSAGE, strResponseCodeMessage
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
		If Len(nAmountToCharge) > 0 Then
			objAddPayment.NewSQL
			objAddPayment.StoreValue PAYMENTS_AMOUNTTOCHARGE, nAmountToCharge
			objAddPayment.SelectRecordForUpdate PAYMENTS_ID, nPaymentsID
			objAddPayment.UpdateUsingStoredValues "Payments"	
		End If
	End If
	
	If Len(nOnlineID)>0 Then

		nOfflineID = GetNewPaymentID(nOrderEntryOrderID,strTransactionCode)
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
	
	Function GetNewPaymentID(nOrderEntryOrderID,strTransactionCode)
		
		Dim objGetPayment
		Set objGetPayment = new CDBInterface
		objGetPayment.Init arrCustDetails, nTotCustDetails
  
		objGetPayment.NewSQL
		objGetPayment.AddFieldToQuery PAYMENTS_ID
		objGetPayment.AddWhere PAYMENTS_ORDERID, "=", nOrderEntryOrderID, "number", "and"
		objGetPayment.AddWhere PAYMENTS_TRANSACTIONCODE, "LIKE", strTransactionCode, "text", "and"
		objGetPayment.QueryDatabase
  
		objGetPayment.GetFirst 
		
		GetNewPaymentID = objGetPayment.getValue(PAYMENTS_ID)
	
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