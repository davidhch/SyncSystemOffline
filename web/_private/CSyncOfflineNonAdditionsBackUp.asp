<%
Class CSyncOffline

	'************************************************************
	'** Copyright DestiNet Limited
	'** CSyncOnline	v0.0l
	'**
	'**	Dependencies:
	'**	Class CSyncLibrary
	'**
	'************************************************************
				
	' Declare Private Objects / Variables
	Private objSyncLib
		
	'************************************************************
	'**	Initialize and Terminate Sub Required by every class
	'************************************************************
	Private Sub Class_Initialize()	
		Set objSyncLib = new CSyncLibrary
		'objSyncLib.setbDebugOutput(True)
		objSyncLib.debugMessage("CSyncOffline Created")	
	End Sub
		
	Private Sub Class_Terminate()
		objSyncLib.debugMessage("CSyncOffline Destroyed")
		Set objSyncLib = nothing
	End Sub
	'************************************************************
	
	'******************	
	' addOrder
	'******************		
	Public Sub addOrder(nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		strMessage = "No message from addOrder"
		
		objSyncLib.GetOfflineID nOnlineID, 5, nOfflineID

		If objSyncLib.isInsert(ORDERS_ID,nOfflineID,"number") Then
			nOfflineID = objSyncLib.saveOrder(nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,objSyncLib.createUniqueID())
			objSyncLib.setErrorResponse True
		Else
			objSyncLib.updateOrder nOfflineID,nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID
		End If
		
		objSyncLib.CreateOfflineSyncLookup nOnlineID,nOfflineID,5,0,1 
		
		'objSyncLib.CreateXMLResponse bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID
		objSyncLib.CreateXMLResponse_SaveOrder 1,strMessage,nOfflineID,nNewZappCustomerID,nOrderEntryCustomerID
	
	End Sub	
	
	'******************	
	' addOrderDetails
	'******************	
	Public Sub addOrderDetails(nID,nOrderID,nOrderEntryOrderID,nProductID,nOrderEntryProductID,nQuantity,nUnitPrice,nDiscount,nAccountType,nCustomerID,nNewZappCID,dRenewalDate,bStatistics,bFrontPage,bServerSideIncludes,bSSL,bIndexServer,bCurrent,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
			
		Dim nOrderEntryOrderDetailsID
		Dim strMessage
		
		strMessage = "No message from addOrderDetails"
			
		'Uses sub to set values
		'objSyncLib.GetOnlineID nOrderID, 5, nNewZappOrderID
		objSyncLib.GetOfflineID nID, 4, nOfflineID
			
		objSyncLib.logSoapDebug("nOfflineID ==== " & nOfflineID)
		'True, a new record is required, False, update the record 
		'If objSyncLib.isInsert(ORDERDETAILS_ORDERID,nOrderEntryOrderID,"number") Then
		If objSyncLib.isInsert(ORDERDETAILS_ID,nOfflineID,"number") Then
			nOfflineID = objSyncLib.saveOrderDetails(nOrderEntryOrderID,nOrderEntryProductID,nQuantity,nUnitPrice,nDiscount,objSyncLib.createUniqueID())
			objSyncLib.setErrorResponse True
		Else		
			objSyncLib.updateOrderDetails nOfflineID,nOrderEntryOrderID,nProductID,nQuantity,nUnitPrice,nDiscount,nOnlineID
			objSyncLib.setErrorResponse True
		End If
			
		objSyncLib.CreateOfflineSyncLookup nID,nOfflineID,4,0,1 
			
		' Return a message
		objSyncLib.CreateXMLResponse bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID
			
	End Sub	
	
	Public Sub addPayment(nPaymentsID,nOrderID,nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Dim strMessage
		
		objSyncLib.GetOfflineID nOnlineID, 3, nOfflineID
		
		strMessage = "No message from addPaymentPayments"
		
		If objSyncLib.isInsert(PAYMENTS_ID,nOfflineID,"number") Then
			nOfflineID = objSyncLib.savePayment(nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
			objSyncLib.setErrorResponse True
		Else
			objSyncLib.updatePayments nOfflineID,nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge
			objSyncLib.setErrorResponse True
		End If
		
		objSyncLib.CreateOfflineSyncLookup nOnlineID,nOfflineID,3,0,1 
		
		' Return a message
		objSyncLib.CreateXMLResponse bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID
			
	End Sub
	
	Public Sub addProduct(nProductID,nOnlineProductID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive,nOrderEntryProductID,nNewZappCustomerID,nOrderEntryCustomerID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Dim strMessage
		
		objSyncLib.GetOfflineID nOnlineID, 2, nOfflineID
		
		strMessage = "No message from addProduct"
		
		If objSyncLib.isInsert(PRODUCT_ID,nOfflineID,"number") Then
			nOfflineID = objSyncLib.saveProduct(strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive)
			objSyncLib.setErrorResponse True
		Else	
			objSyncLib.updateProduct nOfflineID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive
			objSyncLib.setErrorResponse True
		End If
		
		objSyncLib.CreateOfflineSyncLookup nOnlineID,nOfflineID,2,0,1 
		
		' Return a message
		objSyncLib.CreateXMLResponse bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID
		
	End Sub
	
	Public Sub addCustomer(nCustomerID,nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit,strCustomerFirstName,strCustomerLastName,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted) 
	
	'***** WARNING *****
	' THIS SUBROUTINE ASSUMES ONLY 1 CUSTOMER EXISTS IN NEWZAPP!!!!!!!!
	
	Dim strMessage
		
	objSyncLib.GetOfflineID nOnlineID, 6, nOfflineID
	
	strMessage = "No message from addCustomer"
	
	If objSyncLib.isInsert(CUSTOMERS_ID,nOfflineID,"number") AND objSyncLib.isInsert(CUSTOMERS_NEWZAPPCID,nNewZappCID,"number") Then
		nOfflineID = objSyncLib.saveCustomer(nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit,strCustomerFirstName,strCustomerLastName,objSyncLib.createUniqueID())
		objSyncLib.setErrorResponse True
	Else	
		If Len(nOfflineID)>0 Then
			' Nothing we are good
		Else
			nOfflineID = objSyncLib.getAutoGeneratedID(nNewZappCID,CUSTOMERS_ID,CUSTOMERS_NEWZAPPCID)
		End If
		objSyncLib.updateCustomer nOfflineID,nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit
		objSyncLib.setErrorResponse True
	End If
	
	objSyncLib.CreateOfflineSyncLookup nOnlineID,nOfflineID,6,0,1 
	
	' Return a message
	nOrderEntryCustomerID = nOfflineID
	objSyncLib.CreateXMLResponse bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID

	End Sub
			
End Class

'response.write "<br>Test response.write"
%>
