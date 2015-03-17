<%

Class CSyncLibrary

	'************************************************************
	'** Copyright DestiNet Limited
	'** CSyncLibrary	v0.0l
	'**
	'**	Dependencies:
	'**	none
	'**
	'************************************************************

	' Declare Private Objects / Variables
	
	Private objSoap
	Private objDBInterface
	
	Public m_nIndex ' represents a member variable m_
	Public m_bDebugOutput
	Public m_bSendResponse
	
	 ' By default we dont send responses
	
	'************************************************************
	'**	Initialize and Terminate Sub Required by every class
	'************************************************************
	Private Sub Class_Initialize()
		m_bDebugOutput = False ' default value
		m_bSendResponse = False
		Set objSoap = New CSoap
		objSoap.Init
		
		Set objDBInterface = Nothing
		Set objDBInterface = New CDBInterface
		objDBInterface.Init arrCustDetails, nTotCustDetails	
		
	End Sub
	
	Private Sub Class_Terminate()
		set objSoap = Nothing
		Set objDBInterface = Nothing
	End Sub
	'************************************************************
	
	Public Sub setbDebugOutput(bValue)
		m_bDebugOutput = bValue
	End Sub
	
	Public Sub debugMessage(strMessage)
		If m_bDebugOutput = True Then
			'response.write "<br>" & strMessage
		End If
	End Sub
	
	Public Sub setErrorResponse(bValue)
		m_bSendResponse = True
	End Sub
	
	Public Sub CreateXMLResponse(bReportError,strMessage,nNewZappCustomerID,nOrderEntryCustomerID)
		If m_bSendResponse = True Then
			responseXML = responseXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			responseXML = responseXML & "<SOAP:Body>"
			responseXML = responseXML & "<MessageReturn>"
			responseXML = responseXML & "<Time>"&Now()&"</Time>"
			responseXML = responseXML & "<Message>"&bReportError&"</Message>"
			responseXML = responseXML & "<MessageResponse>"&strMessage&"</MessageResponse>"
			responseXML = responseXML & "<NewZappCustomerID>"&nNewZappCustomerID&"</NewZappCustomerID>"
			responseXML = responseXML & "<OrderEntryCustomerID>"&nOrderEntryCustomerID&"</OrderEntryCustomerID>"
			responseXML = responseXML & "</MessageReturn>"
			responseXML = responseXML & "</SOAP:Body>"
			responseXML = responseXML & "</SOAP:Envelope>"
			
			'Writes the XML to screen
			Response.Write (responseXML) 
			'logSoapDebug "WRITTEN BACK: "&responseXML 
		End If
	End Sub	
	
	Public Sub CreateXMLResponse_SaveOrder(bAction,strMessage,nOrderID,nNewZappCustomerID,nOrderEntryCustomerID)
	
		returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		returnXML = returnXML & "<SOAP:Body>"
		returnXML = returnXML & "<MessageReturn>"
		returnXML = returnXML & "<Version>DC-JM 0.01</Version>"
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
	
	Public Sub CreateOfflineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		'Response.Write "<br>Inside Sub"
		If Len(nTableID) > 0 AND Len(nOfflineID) > 0 Then
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery SYNCLOOKUP_ID
			objDBInterface.AddWhere SYNCLOOKUP_ONLINEID, "=", nOnlineID, "number", "and"
			objDBInterface.AddWhere SYNCLOOKUP_TABLEID, "=", nTableID, "number", "and"
			objDBInterface.QueryDatabase
		
			nRecords = objDBInterface.GetNumberOfRecords()
				
			If nRecords = 0 Then
				saveSyncData nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			Else
				nSyncLookupID = objDBInterface.GetValue(SYNCLOOKUP_ID)
				updateSyncData nSyncLookupID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			End If
			
			logSoapDebug("****VARIABLES OK nTableID = " &nTableID& " - nOfflineID = " & nOfflineID & " - nOnlineID = " & nOnlineID)
			CompleteSyncProcess nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			
		Else
			'Variables required to complete the process are missing
			logSoapDebug("****VARIABLES MISSING nTableID = " &nTableID& " - nOfflineID = " & nOfflineID & " - nOnlineID = " & nOnlineID)
		End If

	End Sub	
	
	Public Sub InitOfflineSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)

			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery SYNCLOOKUP_ID
			objDBInterface.AddWhere SYNCLOOKUP_OFFLINEID, "=", nOfflineID, "number", "and"
			objDBInterface.AddWhere SYNCLOOKUP_TABLEID, "=", nTableID, "number", "and"
			objDBInterface.QueryDatabase
		
			nRecords = objDBInterface.GetNumberOfRecords()
				
			If nRecords = 0 Then
				saveSyncData nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			Else
				nSyncLookupID = objDBInterface.GetValue(SYNCLOOKUP_ID)
				updateSyncData nSyncLookupID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			End If

	End Sub	
	
	Public Sub saveSyncData(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		objDBInterface.NewSQL         
		objDBInterface.StoreValue SYNCLOOKUP_ONLINEID, nOnlineID
		objDBInterface.StoreValue SYNCLOOKUP_OFFLINEID, nOfflineID
		objDBInterface.StoreValue SYNCLOOKUP_TABLEID, nTableID
		objDBInterface.StoreValue SYNCLOOKUP_LOCKED, bLocked
		objDBInterface.StoreValue SYNCLOOKUP_EXECUTED, bExecuted
		objDBInterface.InsertUsingStoredValues "SyncLookUp"
	End Sub
	
	Public Sub updateSyncData(nSyncLookupID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		UpdateSingleRecordDB nOnlineID, nSyncLookupID, "SyncLookUp", SYNCLOOKUP_ONLINEID, SYNCLOOKUP_ID
		UpdateSingleRecordDB OfflineID, nSyncLookupID, "SyncLookUp", SYNCLOOKUP_OFFLINEID, SYNCLOOKUP_ID
		'***** DO NOT COMMENT IN THIS LINE - reason: Changing tableID will associate IDs to wrong records !!!!!!
		'UpdateSingleRecordDB nTableID, nSyncLookupID, "SyncLookUp", SYNCLOOKUP_TABLEID, SYNCLOOKUP_ID 
		UpdateSingleRecordDB bLocked, nSyncLookupID, "SyncLookUp", SYNCLOOKUP_LOCKED, SYNCLOOKUP_ID
		UpdateSingleRecordDB bExecuted, nSyncLookupID, "SyncLookUp", SYNCLOOKUP_EXECUTED, SYNCLOOKUP_ID
	End Sub
		
	Public Sub CompleteSyncProcess(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Dim soapURL
		Dim soapXML
		
		'soapURL = "http://www01.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp" 
		soapURL = "http://system.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp" 
		
		objSoap.clearSendNodes()
		objSoap.AddSendNode "OnlineID", nOnlineID
		objSoap.AddSendNode "OfflineID", nOfflineID
		objSoap.AddSendNode "TableID", nTableID
		objSoap.AddSendNode "Locked", bLocked
		objSoap.AddSendNode "Executed", bExecuted
		
		soapXML = objSoap.buildSendSoap
		'build the soap to send
		
		soapResponseCode 	= objSoap.SendSoap(soapXML, soapURL)	' this returns 200 or 404 etc accordingly
		soapResponse 		= objSoap.getResponseText				' this returns the actual response as text
		
		Response.Write soapResponse
		
	End Sub
		
	Public Sub GetOfflineID(nNonLocalID, nTable, nLocalID)			
		If Len(nNonLocalID) > 0 Then
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery SYNCLOOKUP_OFFLINEID
			objDBInterface.AddWhere SYNCLOOKUP_ONLINEID, "=", nNonLocalID, "number", "and"
			objDBInterface.AddWhere SYNCLOOKUP_TABLEID, "=", nTable, "number", "and"	
			objDBInterface.QueryDatabase

			nRecords = objDBInterface.GetNumberOfRecords()
				
			If nRecords > 0 Then
				nLocalID = objDBInterface.GetValue(SYNCLOOKUP_OFFLINEID)
			Else
				nLocalID = ""
			End If
			
			logSoapDebug("****nRecords ==== " & nRecords)
			logSoapDebug("nNonLocalID ==== " & nNonLocalID)
			logSoapDebug("nTable ==== " & nTable)
			logSoapDebug("nLocalID ==== " & nLocalID)
		End If
	End Sub
		
	Public Function GetOfflineProductID(nOrderEntryProductID)
				
		objDBInterface.NewSQL
		objDBInterface.AddFieldToQuery PRODUCT_ID
		objDBInterface.AddWhere PRODUCT_ID, "=", nOrderEntryProductID, "number", "and"
		objDBInterface.QueryDatabase

		nRecords = objDBInterface.GetNumberOfRecords()
			
		If nRecords > 0 Then
			GetOnlineProductID = objDBInterface.GetValue(PRODUCT_ID)
		End If
		
	End Function
		
	Public Function createUniqueID()
		
		Set objTypeLib = Server.CreateObject("Scriptlet.TypeLib")
		strGUID = objTypeLib.Guid
		createUniqueID = Left( strGUID, Len( strGUID ) -2 )
		Set objTypeLib = Nothing
	
	End Function
	
	Public Function getAutoGeneratedPaymentID(strUniqueOEID,nRecordID,nUniqueID,nOrderID) ' rename vars
		If Len(strUniqueOEID)>0 Then
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery nRecordID
			objDBInterface.AddWhere nUniqueID, "LIKE", strUniqueOEID, "text", "and"
			objDBInterface.AddWhere PAYMENTS_ORDERID, "=", nOrderID, "number", "and"
			objDBInterface.QueryDatabase
		
			objDBInterface.GetLast() ' because we only want the last one added in the rare case of duplicates
			If objDBInterface.GetNumberOfRecords() > 0 Then
				getAutoGeneratedPaymentID = objDBInterface.GetValue(nRecordID)
			End If
		Else
			'Error
		End If
	
	End Function
	
	Public Function getAutoGeneratedID(strUniqueOEID,nRecordID,nUniqueID) ' rename vars
		If Len(strUniqueOEID)>0 Then
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery nRecordID
			objDBInterface.AddWhere nUniqueID, "LIKE", strUniqueOEID, "text", "and"
			objDBInterface.QueryDatabase
		
			If objDBInterface.GetNumberOfRecords() > 0 Then
				getAutoGeneratedID = objDBInterface.GetValue(nRecordID)
			End If
		Else
			'Error
		End If
	
	End Function
	
	Public Function saveOrder(nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,strUniqueOEID,nInstallCID,nResellerID)
		
		If Len(nOrderEntryCID)=0 Then	
			
			GetOfflineID nCustomerID, 6, nOrderEntryCID
			
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery WEBHOSTING_CUSTOMERID
			objDBInterface.AddWhere WEBHOSTING_CID, "=", nNewZappCID, "number", "and"
			objDBInterface.QueryDatabase
		
			nOrderEntryCID = objDBInterface.GetValue(WEBHOSTING_CUSTOMERID)
			
		End If
			
		objDBInterface.NewSQL
		objDBInterface.StoreValue ORDERS_CUSTOMERID, nOrderEntryCID
		objDBInterface.StoreValue ORDERS_QUOTATIONID, nQuotationID
		objDBInterface.StoreValue ORDERS_EMPLOYEEID, nEmployeeID
		
		If Len(dOrderDate) > 0 Then
			objDBInterface.StoreValue ORDERS_ORDERDATE, dOrderDate
		Else
			objDBInterface.StoreValue ORDERS_ORDERDATE, Date()
		End If
		
		objDBInterface.StoreValue ORDERS_PURCHASEORDERNUMBER, strPurchaseOrderNumber
		objDBInterface.StoreValue ORDERS_INVOICEDESCRIPTION, strInvoiceDescription
		objDBInterface.StoreValue ORDERS_SHIPNAME, strShipName
		objDBInterface.StoreValue ORDERS_SHIPADDRESS, strShipAddress
		objDBInterface.StoreValue ORDERS_SHIPCITY, strShipCity
		objDBInterface.StoreValue ORDERS_SHIPSTATEORPROVINCE, strShipStateOrProvince
		objDBInterface.StoreValue ORDERS_SHIPPOSTCODE, strShipPostalCode
		objDBInterface.StoreValue ORDERS_SHIPCOUNTRY, strShipCountry
		objDBInterface.StoreValue ORDERS_SHIPPHONENUMBER, strShipPhoneNumber
		objDBInterface.StoreValue ORDERS_SHIPDATE, dShipDate
		objDBInterface.StoreValue ORDERS_SHIPPINGMETHODID, nShippingMethodID 
		objDBInterface.StoreValue ORDERS_FREIGHTCHARGE, nFreightCharge
		objDBInterface.StoreValue ORDERS_SALESTAXRATE, nSalesTaxRate
		
		If Len(nLeadReferralID) > 0 Then
			objDBInterface.StoreValue ORDERS_LEADREFERRALID, nLeadReferralID
		End If
		
		objDBInterface.StoreValue ORDERS_PROPOSALREFERENCE, strProposalReference
		
		If Len(dEntryDate) > 0 Then
			objDBInterface.StoreValue ORDERS_ENTRYDATE, dEntryDate
		Else
			objDBInterface.StoreValue ORDERS_ENTRYDATE, Date()
		End If
		
		objDBInterface.StoreValue ORDERS_COMMISSIONPAID, bCommissionPaid
		objDBInterface.StoreValue ORDERS_INVOICEID, nInvoiceID
		objDBInterface.StoreValue ORDERS_UNIQUEKEY, strUniqueOEID
		objDBInterface.StoreValue ORDERS_INSTALLCID, nInstallCID
		objDBInterface.StoreValue ORDERS_RESELLERID, nResellerID
		objDBInterface.InsertUsingStoredValues "Orders"
		
		saveOrder = getAutoGeneratedID(strUniqueOEID,ORDERS_ID,ORDERS_UNIQUEKEY)
		
		LogOrder saveOrder, -1, "Order Received from NZ and INSERTED", 1
	
	End Function
	
	Public Function saveOrderDetails(nOrderEntryOrderID,nOrderEntryProductID,nQuantity,nUnitPrice,nDiscount,strUniqueOEID)
			
		objDBInterface.NewSQL         
		objDBInterface.StoreValue ORDERDETAILS_ORDERID, nOrderEntryOrderID
		objDBInterface.StoreValue ORDERDETAILS_PRODUCTID, nOrderEntryProductID
		objDBInterface.StoreValue ORDERDETAILS_QUANTITY, nQuantity
		objDBInterface.StoreValue ORDERDETAILS_UNITPRICE, nUnitPrice
		objDBInterface.StoreValue ORDERDETAILS_DISCOUNT, nDiscount
		objDBInterface.StoreValue ORDERDETAILS_UNIQUEKEY, strUniqueOEID
		objDBInterface.StoreValue ORDERDETAILS_NEWZAPPCID, nNewZappCID
		objDBInterface.StoreValue ORDERDETAILS_NEWZAPPCID, nNewZappCID
		objDBInterface.InsertUsingStoredValues "Order Details"
	
		saveOrderDetails = getAutoGeneratedID(strUniqueOEID,ORDERDETAILS_ID,ORDERDETAILS_UNIQUEKEY)
		
		LogOrder nOrderEntryOrderID, saveOrderDetails, "OrderDetail Received from NZ and INSERTED", 2	
	End Function
	
	Public Function savePayment(nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
	
		objDBInterface.NewSQL  
		objDBInterface.StoreValue PAYMENTS_ORDERID, nOrderEntryOrderID
		objDBInterface.StoreValue PAYMENTS_PAYMENTAMOUNT, nPaymentAmount
		objDBInterface.StoreValue PAYMENTS_PAYMENTDATE, dPaymentDate
		'objDBInterface.StoreValue PAYMENTS_CREDITCARDID, nCreditCardID
		objDBInterface.StoreValue PAYMENTS_PAYMENTMETHODID, nPaymentMethodID
		objDBInterface.StoreValue PAYMENTS_TRANSACTIONCODE, strTransactionCode
		objDBInterface.StoreValue PAYMENTS_PROCESSINGVALID, strProcessingValid
		objDBInterface.StoreValue PAYMENTS_PROCESSINGCODE, strProcessingCode
		objDBInterface.StoreValue PAYMENTS_PROCESSINGCODEMESSAGE, strProcessingCodeMessage
		objDBInterface.StoreValue PAYMENTS_AUTHCODE, strAuthCode
		objDBInterface.StoreValue PAYMENTS_AUTHMESSAGE, strAuthMessage
		objDBInterface.StoreValue PAYMENTS_RESPONSECODE, strResponseCode
		objDBInterface.StoreValue PAYMENTS_RESPONSECODEMESSAGE, strResponseCodeMessage
		objDBInterface.StoreValue PAYMENTS_AMOUNTTOCHARGE, nAmountToCharge
		objDBInterface.InsertUsingStoredValues "Payments"
		
		'savePayment = getAutoGeneratedID(strTransactionCode,PAYMENTS_ID,PAYMENTS_TRANSACTIONCODE)
		savePayment = getAutoGeneratedPaymentID(strTransactionCode,PAYMENTS_ID,PAYMENTS_TRANSACTIONCODE,nOrderEntryOrderID)
		
		LogOrder nOrderEntryOrderID, -1, "Payment Received from NZ and INSERTED", 3
		
	End Function
	
	Public Function saveProduct(strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive)
		
		objDBInterface.NewSQL			
		objDBInterface.StoreValue PRODUCTS_PRODUCTNAME, strProductName
		objDBInterface.StoreValue PRODUCTS_PRODUCTCATEGORYID, nProductCategoryID
		objDBInterface.StoreValue PRODUCTS_UNITPRICE, nUnitPrice
		objDBInterface.StoreValue PRODUCTS_UNITCOST, nUnitCost
		objDBInterface.StoreValue PRODUCTS_ONEOFFCOST, strOneOffCost
		objDBInterface.StoreValue PRODUCTS_SUPPLIERID, nSupplierID
		objDBInterface.StoreValue PRODUCTS_DESCRIPTION, strDescription
		objDBInterface.StoreValue PRODUCTS_SYSTEMDESCRIPTION, strSystemDescription
		objDBInterface.StoreValue PRODUCTS_SETUPPRICE, nSetupPrice
		objDBInterface.StoreValue PRODUCTS_ORDER, nOrder
		objDBInterface.StoreValue PRODUCTS_ACTIVE, bActive
		objDBInterface.InsertUsingStoredValues "Products"
		
		' ideally we'd like a unique identifier - hopefully it's name is hmmmm
		savePayment = getAutoGeneratedID(strProductName,PRODUCTS_ID,PRODUCTS_PRODUCTNAME)
		
	End Function
	
	Public Function saveCustomer(nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit,strCustomerFirstName,strCustomerLastName,strUniqueID)
	
		objDBInterface.NewSQL	
				
		objDBInterface.StoreValue CUSTOMERS_COMPANYNAME, strUniqueID ' strCompanyName - we'll update to this
		objDBInterface.StoreValue CUSTOMERS_BILLINGADDRESS, strBillingAddress
		objDBInterface.StoreValue CUSTOMERS_CITY, strCity
		objDBInterface.StoreValue CUSTOMERS_STATEORPROVINCE, strStateOrProvince
		objDBInterface.StoreValue CUSTOMERS_POSTALCODE, strPostalCode
		objDBInterface.StoreValue CUSTOMERS_COUNTRYID, nCountryID
		objDBInterface.StoreValue CUSTOMERS_CONTACTTITLE, strContactTitle
		objDBInterface.StoreValue CUSTOMERS_PHONENUMBER, strPhoneNumber
		objDBInterface.StoreValue CUSTOMERS_FAXNUMBER, strFaxNumber
		objDBInterface.StoreValue CUSTOMERS_EMAILADDRESS, strEmailAddress
		objDBInterface.StoreValue CUSTOMERS_WEBADDRESS, strWebAddress
		
		objDBInterface.StoreValue CUSTOMERS_NOTES, strNotes
		objDBInterface.StoreValue CUSTOMERS_WEBSITEUSERNAME, strWebSiteUserName
		objDBInterface.StoreValue CUSTOMERS_WEBSITEPASSWORD, strWebSitePassword
		
		If Len(bLeadGenerator) > 0 Then
			objDBInterface.StoreValue CUSTOMERS_LEADGENERATOR, bLeadGenerator
		Else
			objDBInterface.StoreValue CUSTOMERS_LEADGENERATOR, 0
		End If
		
		objDBInterface.StoreValue CUSTOMERS_LEADREFERRALID, nLeadReferralID
		objDBInterface.StoreValue CUSTOMERS_VATCODE, nVATCode
		objDBInterface.StoreValue CUSTOMERS_VATNO, strVATNo
		
		If Len(bReseller) > 0 Then
			objDBInterface.StoreValue CUSTOMERS_RESELLER, bReseller
		Else
			objDBInterface.StoreValue CUSTOMERS_RESELLER, 0
		End If
		
		objDBInterface.StoreValue CUSTOMERS_SHOWCUSTOMER, bShowCustomer
		objDBInterface.StoreValue CUSTOMERS_CREDITCARDCUSTOMER, bCreditCardCustomer
		objDBInterface.StoreValue CUSTOMERS_CREDITLIMIT, nCreditLimit
		objDBInterface.StoreValue CUSTOMERS_NEWZAPPCID, nNewZappCID
		objDBInterface.InsertUsingStoredValues "Customers"
		
		'getAutoGeneratedID(strUniqueID,CUSTOMERS_ID,CUSTOMERS_COMPANYNAME)
		
		
		
		saveCustomer = getAutoGeneratedID(strUniqueID,CUSTOMERS_ID,CUSTOMERS_COMPANYNAME)
		
		objDBInterface.NewSQL	
		objDBInterface.StoreValue CONTACTS_CUSTOMERID, saveCustomer
		objDBInterface.StoreValue CONTACTS_FIRSTNAME, strCustomerFirstName
		objDBInterface.StoreValue CONTACTS_LASTNAME, strCustomerLastName
		objDBInterface.StoreValue CONTACTS_EMAILADDRESS, strEmailAddress
		objDBInterface.InsertUsingStoredValues "Contacts"
		
		' Have done this because there is no unique identifier otherwise
		' alertnatively a full update could be run as below, but that seems a bit wastefull
		
		' updateCustomer saveCustomer,nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit
		 
		UpdateSingleRecordDB strCompanyName, saveCustomer, "Customers", CUSTOMERS_COMPANYNAME, CUSTOMERS_ID
	
	End Function
	
	'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	Public Function saveWebHosting(nAccountType,nCustomerID,nNewZappCID,dRenewalDate,bStatistics,bFrontPage,bServerSideIncludes,bSSL,bIndexServer,bCurrent)
	
		' SHOULD THIS BE saveAccountData??????????? NO!!!
		objDBInterface.NewSQL         
		objDBInterface.StoreValue WEBHOSTING_ACCOUNTTYPE, nAccountType
		objDBInterface.StoreValue WEBHOSTING_CUSTOMERID, nCustomerID
		objDBInterface.StoreValue WEBHOSTING_CID, nNewZappCID
		objDBInterface.StoreValue WEBHOSTING_RENEWALDATE, dRenewalDate
		objDBInterface.StoreValue WEBHOSTING_STATISTICS, bStatistics
		objDBInterface.StoreValue WEBHOSTING_FRONTPAGE, bFrontPage
		objDBInterface.StoreValue WEBHOSTING_SERVERSIDEINCLUDES, bServerSideIncludes
		objDBInterface.StoreValue WEBHOSTING_SSL, bSSL
		objDBInterface.StoreValue WEBHOSTING_INDEXSERVER, bIndexServer
		objDBInterface.StoreValue WEBHOSTING_CURRENT, bCurrent
		objDBInterface.InsertUsingStoredValues "Web Hosting"
		
		'''Need to work out how to add Pop data here after the above save based on new records data
		'''At this point we have a new Web Hosting Record
        'nPopRefID = getWebHostingID(nNewZappCID)

		'objDBInterface.NewSQL         
		'objDBInterface.StoreValue POPACCOUNTS_WEBHOSTINGID, nPopRefID
		'objDBInterface.StoreValue POPACCOUNTS_USERNAME, strUserName
		'objDBInterface.StoreValue POPACCOUNTS_PASSWORD, strPassword
		'objDBInterface.StoreValue POPACCOUNTS_SITEADMIN, 1
		'objDBInterface.StoreValue POPACCOUNTS_PRIMARYACCOUNT, 1
		'objDBInterface.StoreValue POPACCOUNTS_SHELLACCESS, 0
		'objDBInterface.StoreValue POPACCOUNTS_USERFPX, 0
		'objDBInterface.InsertUsingStoredValues "POP Accounts"
        
	End Function
	
	
	'Function getWebHostingID(nNewZappCID)
	
		'objDBInterface.NewSQL
		'objDBInterface.AddFieldToQuery WEBHOSTING_ID
		'objDBInterface.AddWhere WEBHOSTING_CID, "=", nNewZappCID, "number", "and"
		'objDBInterface.QueryDatabase
		
		'getWebHostingID = objDBInterface.GetValue(WEBHOSTING_ID)
	'End Function
	'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	
	Public Sub updateOrder(nOfflineID,nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,nInstallCID,nResellerID)
		
		If Len(nOrderEntryCID)=0 Then
		
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery WEBHOSTING_CUSTOMERID
			objDBInterface.AddWhere WEBHOSTING_CID, "=", nNewZappCID, "number", "and"
			objDBInterface.QueryDatabase
		
			nOrderEntryCID = objDBInterface.GetValue(WEBHOSTING_CUSTOMERID)
			
		End If
	
		'UpdateSingleRecordDB nOrderEntryCID, nOfflineID, "Orders", ORDERS_CUSTOMERID, ORDERS_ID
		UpdateSingleRecordDB nQuotationID, nOfflineID, "Orders", ORDERS_QUOTATIONID, ORDERS_ID
		'UpdateSingleRecordDB nEmployeeID, nOfflineID, "Orders", ORDERS_EMPLOYEEID, ORDERS_ID
		UpdateSingleRecordDB dOrderDate, nOfflineID, "Orders", ORDERS_ORDERDATE, ORDERS_ID
		UpdateSingleRecordDB strPurchaseOrderNumber, nOfflineID, "Orders", ORDERS_PURCHASEORDERNUMBER, ORDERS_ID
		UpdateSingleRecordDB strInvoiceDescription, nOfflineID, "Orders", ORDERS_INVOICEDESCRIPTION, ORDERS_ID
		UpdateSingleRecordDB strShipName, nOfflineID, "Orders", ORDERS_SHIPNAME, ORDERS_ID
		UpdateSingleRecordDB strShipAddress, nOfflineID, "Orders", ORDERS_SHIPADDRESS, ORDERS_ID
		UpdateSingleRecordDB strShipCity, nOfflineID, "Orders", ORDERS_SHIPCITY, ORDERS_ID
		UpdateSingleRecordDB strShipStateOrProvince, nOfflineID, "Orders", ORDERS_SHIPSTATEORPROVINCE, ORDERS_ID
		UpdateSingleRecordDB strShipPostalCode, nOfflineID, "Orders", ORDERS_SHIPPOSTCODE, ORDERS_ID
		UpdateSingleRecordDB strShipCountry, nOfflineID, "Orders", ORDERS_SHIPCOUNTRY, ORDERS_ID
		UpdateSingleRecordDB strShipPhoneNumber, nOfflineID, "Orders", ORDERS_SHIPPHONENUMBER, ORDERS_ID
		UpdateSingleRecordDB dShipDate, nOfflineID, "Orders", ORDERS_SHIPDATE, ORDERS_ID
		'UpdateSingleRecordDB nShippingMethodID, nOfflineID, "Orders", ORDERS_SHIPPINGMETHODID, ORDERS_ID
		UpdateSingleRecordDB nFreightCharge, nOfflineID, "Orders", ORDERS_FREIGHTCHARGE, ORDERS_ID
		UpdateSingleRecordDB nSalesTaxRate, nOfflineID, "Orders", ORDERS_SALESTAXRATE, ORDERS_ID
		UpdateSingleRecordDB nLeadReferralID, nOfflineID, "Orders", ORDERS_LEADREFERRALID, ORDERS_ID
		UpdateSingleRecordDB strProposalReference, nOfflineID, "Orders", ORDERS_PROPOSALREFERENCE, ORDERS_ID
		UpdateSingleRecordDB dEntryDate, nOfflineID, "Orders", ORDERS_ENTRYDATE, ORDERS_ID
		UpdateSingleRecordDB bCommissionPaid, nOfflineID, "Orders", ORDERS_COMMISSIONPAID, ORDERS_ID
		UpdateSingleRecordDB nInvoiceID, nOfflineID, "Orders", ORDERS_INVOICEID, ORDERS_ID
		UpdateSingleRecordDB nInstallCID, nOfflineID, "Orders", ORDERS_INSTALLCID, ORDERS_ID
		UpdateSingleRecordDB nResellerID, nOfflineID, "Orders", ORDERS_RESELLERID, ORDERS_ID
		
		LogOrder nOfflineID, -1, "Order Received from NZ and UPDATED", 1

	End Sub
	
	Public Sub updateOrderDetails(nOfflineID,nOrderID,nProductID,nQuantity,nUnitPrice,nDiscount,nOnlineID)
		'// commented out these updates as the are on FOREIGN KEYS
		'UpdateSingleRecordDB nOrderID, nOfflineID, "Order Details", ORDERDETAILS_ORDERID, ORDERDETAILS_ID
		'UpdateSingleRecordDB nProductID, nOfflineID, "Order Details", ORDERDETAILS_PRODUCTID, ORDERDETAILS_ID
		UpdateSingleRecordDB nQuantity, nOfflineID, "Order Details", ORDERDETAILS_QUANTITY, ORDERDETAILS_ID
		UpdateSingleRecordDB nUnitPrice, nOfflineID, "Order Details", ORDERDETAILS_UNITPRICE, ORDERDETAILS_ID
		UpdateSingleRecordDB nDiscount, nOfflineID, "Order Details", ORDERDETAILS_DISCOUNT, ORDERDETAILS_ID
		'UpdateSingleRecordDB strUniqueOEID, nOfflineID, "Order Details", ORDERDETAILS_UNIQUEKEY, ORDERDETAILS_ID
		'UpdateSingleRecordDB nNewZappCID, nOfflineID, "Order Details", ORDERDETAILS_NEWZAPPCID, ORDERDETAILS_ID
		
		LogOrder nOrderID, nOfflineID, "OrderDetail Received from NZ and UPDATED", 2
	
	End Sub
	
	Public Sub updatePayments(nOfflineID,nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
		
		UpdateSingleRecordDB nOrderEntryOrderID, nOfflineID, "Payments", PAYMENTS_ORDERID, PAYMENTS_ID
		UpdateSingleRecordDB nPaymentAmount, nOfflineID, "Payments", PAYMENTS_PAYMENTAMOUNT, PAYMENTS_ID
		UpdateSingleRecordDB dPaymentDate, nOfflineID, "Payments", PAYMENTS_PAYMENTDATE, PAYMENTS_ID
		'UpdateSingleRecordDB nCreditCardID, nOfflineID, "Payments", PAYMENTS_CREDITCARDID, PAYMENTS_ID
		UpdateSingleRecordDB nPaymentMethodID, nOfflineID, "Payments", PAYMENTS_PAYMENTMETHODID, PAYMENTS_ID
		UpdateSingleRecordDB strTransactionCode, nOfflineID, "Payments", PAYMENTS_TRANSACTIONCODE, PAYMENTS_ID
		UpdateSingleRecordDB strProcessingValid, nOfflineID, "Payments", PAYMENTS_PROCESSINGVALID, PAYMENTS_ID
		UpdateSingleRecordDB strProcessingCode, nOfflineID, "Payments", PAYMENTS_PROCESSINGCODE, PAYMENTS_ID
		UpdateSingleRecordDB strProcessingCodeMessage, nOfflineID, "Payments", PAYMENTS_PROCESSINGCODEMESSAGE, PAYMENTS_ID
		UpdateSingleRecordDB strAuthCode, nOfflineID, "Payments", PAYMENTS_AUTHCODE, PAYMENTS_ID
		UpdateSingleRecordDB strAuthMessage, nOfflineID, "Payments", PAYMENTS_AUTHMESSAGE, PAYMENTS_ID
		UpdateSingleRecordDB strResponseCode, nOfflineID, "Payments", PAYMENTS_RESPONSECODE, PAYMENTS_ID
		UpdateSingleRecordDB strResponseCodeMessage, nOfflineID, "Payments", PAYMENTS_RESPONSECODEMESSAGE, PAYMENTS_ID
		UpdateSingleRecordDB nAmountToCharge, nOfflineID, "Payments", PAYMENTS_AMOUNTTOCHARGE, PAYMENTS_ID
		
		LogOrder nOrderEntryOrderID, -1, "Payment Received from NZ and UPDATED", 3
	
	End Sub
	
	Public Sub updateProduct(nOfflineID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive)
		
		UpdateSingleRecordDB strProductName, nOfflineID, "Products", PRODUCT_NAME, PRODUCT_ID
		UpdateSingleRecordDB nProductCategoryID, nOfflineID, "Products", PRODUCTS_PRODUCTCATEGORYID, PRODUCT_ID
		UpdateSingleRecordDB nUnitPrice, nOfflineID, "Products", PRODUCT_USAGELIMIT, PRODUCT_ID
		UpdateSingleRecordDB nUnitCost, nOfflineID, "Products", PRODUCTS_UNITCOST, PRODUCT_ID
		UpdateSingleRecordDB strOneOffCost, nOfflineID, "Products", PRODUCTS_ONEOFFCOST, PRODUCT_ID
		UpdateSingleRecordDB nSupplierID, nOfflineID, "Products", PRODUCTS_SUPPLIERID, PRODUCT_ID
		UpdateSingleRecordDB strDescription, nOfflineID, "Products", PRODUCTS_SYSTEMDESCRIPTION, PRODUCT_ID
		UpdateSingleRecordDB strSystemDescription, nOfflineID, "Products", PRODUCTS_SYSTEMDESCRIPTION, PRODUCT_ID
		UpdateSingleRecordDB nSetupPrice, nOfflineID, "Products", PRODUCTS_SETUPPRICE, PRODUCT_ID
		UpdateSingleRecordDB nOrder, nOfflineID, "Products", PRODUCTS_ORDER, PRODUCT_ID
		UpdateSingleRecordDB bActive, nOfflineID, "Products", PRODUCT_ACTIVE, PRODUCT_ID

	End Sub
	
	Public Sub updateCustomer (nOfflineID,nNewZappCID,strCompanyName,strBillingAddress,strCity,strStateOrProvince,strPostalCode,nCountryID,strContactTitle,strPhoneNumber,strFaxNumber,strEmailAddress,strWebAddress,strNotes,strWebSiteUserName,strWebSitePassword,bLeadGenerator,nLeadReferralID,nVATCode,strVATNo,bReseller,bShowCustomer,bCreditCardCustomer,nCreditLimit)
	
		UpdateSingleRecordDB strCompanyName, nOfflineID, "Customers", CUSTOMERS_COMPANYNAME, CUSTOMERS_ID
		UpdateSingleRecordDB strBillingAddress, nOfflineID, "Customers", CUSTOMERS_BILLINGADDRESS, CUSTOMERS_ID
		UpdateSingleRecordDB strCity, nOfflineID, "Customers", CUSTOMERS_CITY, CUSTOMERS_ID
		UpdateSingleRecordDB strStateOrProvince, nOfflineID, "Customers", CUSTOMERS_STATEORPROVINCE, CUSTOMERS_ID
		UpdateSingleRecordDB strPostalCode, nOfflineID, "Customers", CUSTOMERS_POSTALCODE, CUSTOMERS_ID
		UpdateSingleRecordDB nCountryID, nOfflineID, "Customers", CUSTOMERS_COUNTRYID, CUSTOMERS_ID
		UpdateSingleRecordDB strContactTitle, nOfflineID, "Customers", CUSTOMERS_CONTACTTITLE, CUSTOMERS_ID
		UpdateSingleRecordDB strPhoneNumber, nOfflineID, "Customers", CUSTOMERS_PHONENUMBER, CUSTOMERS_ID
		UpdateSingleRecordDB strFaxNumber, nOfflineID, "Customers", CUSTOMERS_FAXNUMBER, CUSTOMERS_ID
		UpdateSingleRecordDB strEmailAddress, nOfflineID, "Customers", CUSTOMERS_EMAILADDRESS, CUSTOMERS_ID
		UpdateSingleRecordDB strWebAddress, nOfflineID, "Customers", CUSTOMERS_WEBADDRESS, CUSTOMERS_ID
		UpdateSingleRecordDB strNotes, nOfflineID, "Customers", CUSTOMERS_NOTES, CUSTOMERS_ID
		UpdateSingleRecordDB strWebSiteUserName, nOfflineID, "Customers", CUSTOMERS_WEBSITEUSERNAME, CUSTOMERS_ID
		UpdateSingleRecordDB strWebSitePassword, nOfflineID, "Customers", CUSTOMERS_WEBSITEPASSWORD, CUSTOMERS_ID
		UpdateSingleRecordDB bLeadGenerator, nOfflineID, "Customers", CUSTOMERS_LEADGENERATOR, CUSTOMERS_ID
		UpdateSingleRecordDB nLeadReferralID, nOfflineID, "Customers", CUSTOMERS_LEADREFERRALID, CUSTOMERS_ID
		UpdateSingleRecordDB nVATCode, nOfflineID, "Customers", CUSTOMERS_VATCODE, CUSTOMERS_ID
		UpdateSingleRecordDB strVATNo, nOfflineID, "Customers", CUSTOMERS_VATNO, CUSTOMERS_ID
		UpdateSingleRecordDB bReseller, nOfflineID, "Customers", CUSTOMERS_RESELLER, CUSTOMERS_ID
		UpdateSingleRecordDB bShowCustomer, nOfflineID, "Customers", CUSTOMERS_SHOWCUSTOMER, CUSTOMERS_ID
		UpdateSingleRecordDB bCreditCardCustomer, nOfflineID, "Customers", CUSTOMERS_CREDITCARDCUSTOMER, CUSTOMERS_ID
		UpdateSingleRecordDB nCreditLimit, nOfflineID, "Customers", CUSTOMERS_CREDITLIMIT, CUSTOMERS_ID
		UpdateSingleRecordDB nNewZappCID, nOfflineID, "Customers", CUSTOMERS_NEWZAPPCID, CUSTOMERS_ID
		
	End Sub
	
	'it follows that this function will be used to update a table with a value
	Public Sub UpdateSingleRecordDB(objValue,objID,objTable,nItem,nLookup)
		
		If Len(objValue) > 0 Then	
			objDBInterface.NewSQL
			objDBInterface.StoreValue nItem, objValue
			objDBInterface.SelectRecordForUpdate nLookup, objID
			objDBInterface.UpdateUsingStoredValues objTable	
		End If
					
	End Sub
	
	Public Function isInsert(nLookup,objValue,strLookupType)
		
		isInsert = False 'Default position is False until a record is found
		If Len(objValue)>0 Then
		
			objDBInterface.NewSQL
			objDBInterface.AddFieldToQuery nLookup
			objDBInterface.AddWhere nLookup, "=", objValue, strLookupType, "and"
			objDBInterface.QueryDatabase
			
			'response.write "<br> nRecords = " & objDBInterface.GetNumberOfRecords()
			
			If objDBInterface.GetNumberOfRecords() = 0 Then
				isInsert = True 'No records found so one can be added
			End If
		Else
			isInsert = True 'No records found so one can be added
		End If
		
	End Function
	
	
	
	'************************************************************
	'**	Process the initial soap request from ORDER ENTRY - all requests go through processSoap()
	'************************************************************
	
	Public Sub processSoap()
	
		Const bLocked = 1
		Const bExecuted = 0
		'Const strSoapFolder = "http://www01.newzapp.co.uk/SOAP/"
		Const strSoapFolder = "http://system.newzapp.co.uk/SOAP/"
		
		'response.write objSoap.ReadSingleNode("FunctionCall")
	
		Select Case objSoap.ReadSingleNode("FunctionCall")
			Case "SendCustomerDetailsXML"
				
				nOnlineID	= ""
				nOfflineID	= objSoap.ReadSingleNode ("CustomerID")
				nTableID	= 6
				
				soapURL = strSoapFolder & "AddCustomerDetailsSOAPOO.asp"
				
				objSoap.clearSendNodes()
				addNodeWithSameTagAsNodeRead "ID"
				addNodeWithSameTagAsNodeRead "ResellerID"
				addNodeWithSameTagAsNodeRead "CustomerID"
				addNodeWithSameTagAsNodeRead "OrderEntryCID"
				addNodeWithSameTagAsNodeRead "CustomerName"
				addNodeWithSameTagAsNodeRead "Company"
				addNodeWithSameTagAsNodeRead "Address1"
				addNodeWithSameTagAsNodeRead "Address2"
				addNodeWithSameTagAsNodeRead "City"
				addNodeWithSameTagAsNodeRead "County"
				addNodeWithSameTagAsNodeRead "PostCode"
				addNodeWithSameTagAsNodeRead "EmailAddress"
				addNodeWithSameTagAsNodeRead "TelNo"
				addNodeWithSameTagAsNodeRead "NewsletterTitle"
				addNodeWithSameTagAsNodeRead "NewsletterEmail"
				addNodeWithSameTagAsNodeRead "StartDate"
				addNodeWithSameTagAsNodeRead "LastDate"
				addNodeWithSameTagAsNodeRead "TotalNewsSent"
				addNodeWithSameTagAsNodeRead "LastUsed"
				addNodeWithSameTagAsNodeRead "TimesLoggedIn"
				addNodeWithSameTagAsNodeRead "Active"
				addNodeWithSameTagAsNodeRead "Demo"
				addNodeWithSameTagAsNodeRead "HashKey"
				addNodeWithSameTagAsNodeRead "MTA"
				' addNodeWithSameTagAsNodeRead "NewZappCID"
				addNodeWithSameTagAsNodeRead "CreditLimit"
				addNodeWithSameTagAsNodeRead "VATCode"
				addNodeWithSameTagAsNodeRead "VATNo"
				
				' THIS WOULD READ IN MULTIPLE NEWZAPP CIDS AND THEN ADD THEM TO THE SOAP TO BE SENT TO NZ
				objSoap.AddOpenNode ("NewZappCIDs")
				objSoap.SetNodeList ("NewZappCID")

					For i = 0 to (objSoap.GetNodeListLength()-1)
						objSoap.AddSendNode "NewZappCID", objSoap.getItemText(i)
					Next
					
				objSoap.AddCloseNode ("NewZappCIDs")

				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
				
			Case "SendOrderXML"
			
				nOnlineID	= ""
				nOfflineID 	= objSoap.ReadSingleNode ("OrderID")
				nTableID 	= 5
				
				soapURL = strSoapFolder & "AddOrderSOAPOO.asp"
				
				objSoap.clearSendNodes()
				addNodeWithSameTagAsNodeRead "CustomerID"
				addNodeWithSameTagAsNodeRead "NewZappCID"
				addNodeWithSameTagAsNodeRead "QuotationID"
				addNodeWithSameTagAsNodeRead "EmployeeID"
				addNodeWithSameTagAsNodeRead "OrderDate"
				addNodeWithSameTagAsNodeRead "PurchaseOrderNumber"
				addNodeWithSameTagAsNodeRead "InvoiceDescription"
				addNodeWithSameTagAsNodeRead "ShipName"
				addNodeWithSameTagAsNodeRead "ShipAddress"
				addNodeWithSameTagAsNodeRead "ShipCity"
				addNodeWithSameTagAsNodeRead "ShipStateOrProvince"
				addNodeWithSameTagAsNodeRead "ShipCountry"
				addNodeWithSameTagAsNodeRead "ShipPhoneNumber"
				addNodeWithSameTagAsNodeRead "ShipDate"
				addNodeWithSameTagAsNodeRead "ShippingMethodID"
				addNodeWithSameTagAsNodeRead "FreightCharge"
				addNodeWithSameTagAsNodeRead "SalesTaxRate"
				addNodeWithSameTagAsNodeRead "LeadReferralID"
				addNodeWithSameTagAsNodeRead "ProposalReference"
				addNodeWithSameTagAsNodeRead "EntryDate"
				addNodeWithSameTagAsNodeRead "CommissionPaid"
				addNodeWithSameTagAsNodeRead "ShowCustomer"
				addNodeWithSameTagAsNodeRead "InvoiceID"
				addNodeWithSameTagAsNodeRead "InstallCID"
				addNodeWithSameTagAsNodeRead "ResellerID"
		
				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
				
				LogOrder nOfflineID,-1,"Order Posted to NZ",1
					
			Case "SendOrderDetailsXML"
			
				nOnlineID	= objSoap.ReadSingleNode ("NewZappOrderDetailID")
				nOfflineID 	= objSoap.ReadSingleNode ("OrderDetailID")
				nTableID 	= 4
				
				nJustificationOrderID = objSoap.ReadSingleNode ("OrderID")
				LogOrder nJustificationOrderID,nOfflineID,"OrderDetail Posted to NZ",2
				
				soapURL = strSoapFolder & "AddOrderDetailsSOAPOO.asp"
				
				objSoap.clearSendNodes()
				addNodeWithSameTagAsNodeRead "ID"
				addNodeWithSameTagAsNodeRead "OrderDetailID"
				addNodeWithSameTagAsNodeRead "NewZappOrderDetailID"
				addNodeWithSameTagAsNodeRead "OrderID"
				addNodeWithSameTagAsNodeRead "NewZappOrderID"
				addNodeWithSameTagAsNodeRead "ProductID"
				addNodeWithSameTagAsNodeRead "NewZappProductID"
				addNodeWithSameTagAsNodeRead "Quantity"
				addNodeWithSameTagAsNodeRead "UnitPrice"
				addNodeWithSameTagAsNodeRead "Discount"
				addNodeWithSameTagAsNodeRead "AccountType"
				addNodeWithSameTagAsNodeRead "CustomerID"
				addNodeWithSameTagAsNodeRead "CID"
				addNodeWithSameTagAsNodeRead "RenewalDate"
				addNodeWithSameTagAsNodeRead "Statistics"
				addNodeWithSameTagAsNodeRead "FrontPage"
				addNodeWithSameTagAsNodeRead "ServerSideIncludes"
				addNodeWithSameTagAsNodeRead "SSL"
				addNodeWithSameTagAsNodeRead "ndexServer"
				addNodeWithSameTagAsNodeRead "Current"
				addNodeWithSameTagAsNodeRead "NewZappCID"
				
				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
				
			Case "SendDeleteOrderXML"
			
				nOnlineID	= ""
				nOfflineID 	= objSoap.ReadSingleNode ("InvoiceID")
				nTableID 	= 5
				
				soapURL = strSoapFolder & "DeleteOrderSOAPOO.asp"
				
				objSoap.clearSendNodes()

				addNodeWithSameTagAsNodeRead "InvoiceID"
		
				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			
			Case "SendPaymentXML"
				nOnlineID	= objSoap.ReadSingleNode ("OrderEntryPaymentsID")
				nOfflineID 	= objSoap.ReadSingleNode ("PaymentsID")
				nTableID 	= 3
				
				nJustificationOrderID = objSoap.ReadSingleNode ("OrderID")
				LogOrder nJustificationOrderID,-1,"Payment Posted to NZ",3
				
				soapURL = strSoapFolder & "AddPaymentSOAPOO.asp"
				
				objSoap.clearSendNodes()
				addNodeWithSameTagAsNodeRead "PaymentsID"
				addNodeWithSameTagAsNodeRead "OrderID"
				addNodeWithSameTagAsNodeRead "NewZappOrderID"
				addNodeWithSameTagAsNodeRead "PaymentAmount"
				addNodeWithSameTagAsNodeRead "PaymentDate"
				'addNodeWithSameTagAsNodeRead "CreditCardID"   ' NOTE: BUG CREDIT CARD IDS CANNOT BE PASSED THEY ARE NOT SYNCD
				objSoap.AddSendNode "CreditCardID", ""
				addNodeWithSameTagAsNodeRead "PaymentMethodID"
				addNodeWithSameTagAsNodeRead "TransactionCode"
				addNodeWithSameTagAsNodeRead "ProcessingValid"
				addNodeWithSameTagAsNodeRead "ProcessingCode"
				addNodeWithSameTagAsNodeRead "ProcessingCodeMessage"
				addNodeWithSameTagAsNodeRead "AuthCode"
				addNodeWithSameTagAsNodeRead "AuthMessage"
				addNodeWithSameTagAsNodeRead "ResponseCode"
				addNodeWithSameTagAsNodeRead "ResponseCodeMessage"
				
				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			
			Case "SendProductXML"
				nOnlineID	= objSoap.ReadSingleNode ("ProductID")
				nOfflineID 	= objSoap.ReadSingleNode ("OrderEntryProductID")
				nTableID 	= 2
				
				soapURL = strSoapFolder & "AddProductSOAP.asp"
				objSoap.clearSendNodes()
				addNodeWithSameTagAsNodeRead "ProductID"
				addNodeWithSameTagAsNodeRead "ProductName"
				addNodeWithSameTagAsNodeRead "ProductType"
				addNodeWithSameTagAsNodeRead "UsageLimit"
				addNodeWithSameTagAsNodeRead "Cost"
				addNodeWithSameTagAsNodeRead "Active"
				addNodeWithSameTagAsNodeRead "OrderEntryProductID"
				
				sendSoapXML soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			Case Else
				' Do nothing havent been told what to do
				
		End Select
	End Sub
	
	Public Sub addNodeWithSameTagAsNodeRead(strTag)
		objSoap.AddSendNode strTag, objSoap.ReadSingleNode (strTag)
	End Sub
	
	Sub addSyncNodes(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		objSoap.AddSendNode "OnlineID", nOnlineID
		objSoap.AddSendNode "OfflineID", nOfflineID
		objSoap.AddSendNode "TableID", nTableID
		objSoap.AddSendNode "Locked", bLocked
		objSoap.AddSendNode "Executed", bExecuted
	End Sub
	
	Public Sub sendSoapXML(soapURL,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		If AllowCID(nOfflineID,nTableID) = True Then
		
			InitOfflineSyncLookup nOnlineID,nOfflineID,nTableID,bLocked,bExecuted 
			
			addSyncNodes nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
			soapXML = objSoap.buildSendSoap
		
			logSoapDebug "SENT: "&soapXML
			soapResponseCode 	= objSoap.SendSoap(soapXML, soapURL)	' this returns 200 or 404 etc accordingly
			logSoapDebug "RESPONSE CODE: "&soapResponseCode
			soapResponse 		= objSoap.getResponseText				' this returns the actual response as text
			logSoapDebug "RETURNED: "&soapResponse
			Response.write soapResponse	
		Else
			responseXML = ""
			responseXML = responseXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			responseXML = responseXML & "<SOAP:Body>"
			responseXML = responseXML & "<MessageReturn>"
			responseXML = responseXML & "<Time>"&Now()&"</Time>"
			responseXML = responseXML & "<Message>False</Message>"
			responseXML = responseXML & "<MessageResponse>NOT UPLOADED</MessageResponse>"
			responseXML = responseXML & "<NewZappCustomerID>0</NewZappCustomerID>"
			responseXML = responseXML & "<OrderEntryCustomerID>0</OrderEntryCustomerID>"
			responseXML = responseXML & "</MessageReturn>"
			responseXML = responseXML & "</SOAP:Body>"
			responseXML = responseXML & "</SOAP:Envelope>"
			Response.write responseXML
			
		End If
		
	End Sub 
	
	Public Sub logSoapDebug(strLine)
		'If m_bDebugOutput = True Then
			' this could be change to write to the database
			dim strFilePath
			strFilePath = server.MapPath("logSoap.txt")
			
			Dim fso, f
			
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set f = fso.OpenTextFile(strFilePath, 8, True)
			
			f.Write strLine & " - " & vbcrlf & now() & vbcrlf & "----------------" & vbcrlf & vbcrlf
			
			Set f = Nothing
			Set fso = Nothing
		'End If
	End Sub
	
	'************************************************************
	'**	SYNC THE DATABASES  
	'************************************************************
	Public Sub syncCheck()
	
		Set objDBsyncCheck = Nothing
		Set objDBsyncCheck = New CDBInterface
		objDBsyncCheck.Init arrCustDetails, nTotCustDetails	
	
		objDBsyncCheck.NewSQL
		objDBsyncCheck.AddFieldToQuery SYNCLOOKUP_ID
		objDBsyncCheck.AddFieldToQuery SYNCLOOKUP_ONLINEID
		objDBsyncCheck.AddFieldToQuery SYNCLOOKUP_OFFLINEID
		objDBsyncCheck.AddFieldToQuery SYNCLOOKUP_TABLEID
		objDBsyncCheck.AddWhere SYNCLOOKUP_LOCKED, "=", 1, "number", "and"
		objDBsyncCheck.AddWhere SYNCLOOKUP_EXECUTED, "=", 0, "number", "and"	
		objDBsyncCheck.QueryDatabase
		
		nRecords = objDBsyncCheck.GetNumberOfRecords()
		
		If nRecords > 0 Then
			' We have to sync the records
			bRecordOutOfSync = True
			nCount = 0
			objDBsyncCheck.GetFirst()
			Do While nCount < nRecords		
		
				nOnlineID 	= objDBsyncCheck.GetValue(SYNCLOOKUP_ONLINEID)
				nOfflineID 	= objDBsyncCheck.GetValue(SYNCLOOKUP_OFFLINEID)
				nTableID 	= objDBsyncCheck.GetValue(SYNCLOOKUP_TABLEID)
				
				SyncData nOnlineID,nOfflineID,nTableID ' This does the work
				
				nCount = nCount + 1
				objDBsyncCheck.GetNext()
			Loop
		End If
		
	End Sub
	
	Public Sub SyncData(nOnlineID,nOfflineID,nTableID)
	'response.write "SyncData: <br>"
		Select Case nTableID
			Case 2
			
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery PRODUCTS_ID
				objDBInterface.AddFieldToQuery PRODUCTS_PRODUCTNAME
				objDBInterface.AddFieldToQuery PRODUCTS_PRODUCTCATEGORYID
				objDBInterface.AddFieldToQuery PRODUCTS_UNITCOST
				objDBInterface.AddFieldToQuery PRODUCTS_ACTIVE
				objDBInterface.AddWhere PRODUCTS_ID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				objSoap.clearSendNodes()
				objSoap.AddSendNode "FunctionCall", "SendProductXML"
				objSoap.AddSendNode "ProductID", ""
				objSoap.AddSendNode "ProductName", objDBInterface.GetValue(PRODUCTS_PRODUCTNAME)
				objSoap.AddSendNode "OrderEntryProductID", objDBInterface.GetValue(PRODUCTS_ID)
				objSoap.AddSendNode "ProductType", objDBInterface.GetValue(PRODUCTS_PRODUCTCATEGORYID)
				objSoap.AddSendNode "UsageLimit", ""
				objSoap.AddSendNode "Cost", objDBInterface.GetValue(PRODUCTS_UNITCOST)
				objSoap.AddSendNode "Active", objDBInterface.GetValue(PRODUCTS_ACTIVE)
		
			Case 3
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery PAYMENTS_ID
				objDBInterface.AddFieldToQuery PAYMENTS_ORDERID
				objDBInterface.AddFieldToQuery PAYMENTS_PAYMENTAMOUNT
				objDBInterface.AddFieldToQuery PAYMENTS_PAYMENTDATE
				objDBInterface.AddFieldToQuery PAYMENTS_CREDITCARDID
				objDBInterface.AddFieldToQuery PAYMENTS_PAYMENTMETHODID
				objDBInterface.AddFieldToQuery PAYMENTS_TRANSACTIONCODE
				objDBInterface.AddFieldToQuery PAYMENTS_PROCESSINGVALID
				objDBInterface.AddFieldToQuery PAYMENTS_PROCESSINGCODE
				objDBInterface.AddFieldToQuery PAYMENTS_PROCESSINGCODEMESSAGE
				objDBInterface.AddFieldToQuery PAYMENTS_AUTHCODE
				objDBInterface.AddFieldToQuery PAYMENTS_AUTHMESSAGE
				objDBInterface.AddFieldToQuery PAYMENTS_RESPONSECODE
				objDBInterface.AddFieldToQuery PAYMENTS_RESPONSECODEMESSAGE
				objDBInterface.AddFieldToQuery PAYMENTS_AMOUNTTOCHARGE
				objDBInterface.AddWhere PAYMENTS_ID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				objSoap.clearSendNodes()
				objSoap.AddSendNode "FunctionCall", "SendPaymentXML"
				objSoap.AddSendNode "PaymentsID", objDBInterface.GetValue(PAYMENTS_ID)
				objSoap.AddSendNode "OrderEntryPaymentsID", ""
				objSoap.AddSendNode "OrderID", objDBInterface.GetValue(PAYMENTS_ORDERID)
				objSoap.AddSendNode "PaymentAmount", objDBInterface.GetValue(PAYMENTS_PAYMENTAMOUNT)
				objSoap.AddSendNode "PaymentDate", objDBInterface.GetValue(PAYMENTS_PAYMENTDATE)
				objSoap.AddSendNode "CreditCardID", objDBInterface.GetValue(PAYMENTS_CREDITCARDID)
				objSoap.AddSendNode "PaymentMethodID", objDBInterface.GetValue(PAYMENTS_PAYMENTMETHODID)
				objSoap.AddSendNode "TransactionCode", objDBInterface.GetValue(PAYMENTS_TRANSACTIONCODE)
				objSoap.AddSendNode "ProcessingValid", objDBInterface.GetValue(PAYMENTS_PROCESSINGVALID)
				objSoap.AddSendNode "ProcessingCode", objDBInterface.GetValue(PAYMENTS_PROCESSINGCODE)
				objSoap.AddSendNode "ProcessingCodeMessage", objDBInterface.GetValue(PAYMENTS_PROCESSINGCODEMESSAGE)
				objSoap.AddSendNode "AuthCode", objDBInterface.GetValue(PAYMENTS_AUTHCODE)
				objSoap.AddSendNode "AuthMessage", objDBInterface.GetValue(PAYMENTS_AUTHMESSAGE)
				objSoap.AddSendNode "ResponseCode", objDBInterface.GetValue(PAYMENTS_RESPONSECODE)
				objSoap.AddSendNode "ResponseCodeMessage", objDBInterface.GetValue(PAYMENTS_RESPONSECODEMESSAGE)
				objSoap.AddSendNode "AmountToCharge", objDBInterface.GetValue(PAYMENTS_AMOUNTTOCHARGE)

			Case 4
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery ORDERDETAILS_ID
				objDBInterface.AddFieldToQuery ORDERDETAILS_ORDERID
				objDBInterface.AddFieldToQuery ORDERDETAILS_PRODUCTID
				objDBInterface.AddFieldToQuery ORDERDETAILS_QUANTITY
				objDBInterface.AddFieldToQuery ORDERDETAILS_UNITPRICE
				objDBInterface.AddFieldToQuery ORDERDETAILS_DISCOUNT
				objDBInterface.AddFieldToQuery ORDERDETAILS_UNIQUEKEY
				objDBInterface.AddFieldToQuery ORDERDETAILS_NEWZAPPCID
				objDBInterface.AddWhere ORDERDETAILS_ID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				objSoap.clearSendNodes()
				objSoap.AddSendNode "FunctionCall", "SendOrderDetailsXML"
				objSoap.AddSendNode "ID", objDBInterface.GetValue(ORDERDETAILS_ORDERID)
				objSoap.AddSendNode "OrderID", objDBInterface.GetValue(ORDERDETAILS_ORDERID)
				objSoap.AddSendNode "ProductID", objDBInterface.GetValue(ORDERDETAILS_PRODUCTID)
				objSoap.AddSendNode "Quantity", objDBInterface.GetValue(ORDERDETAILS_QUANTITY)
				objSoap.AddSendNode "UnitPrice", objDBInterface.GetValue(ORDERDETAILS_UNITPRICE)
				objSoap.AddSendNode "Discount", objDBInterface.GetValue(ORDERDETAILS_DISCOUNT)
				objSoap.AddSendNode "UniqueKey", objDBInterface.GetValue(ORDERDETAILS_UNIQUEKEY)
				objSoap.AddSendNode "NewZappCID", objDBInterface.GetValue(ORDERDETAILS_NEWZAPPCID)
				
			Case 5
				Dim nOrderDetailsNewZappCID
			
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery ORDERS_ID
				objDBInterface.AddFieldToQuery ORDERS_QUOTATIONID
				objDBInterface.AddFieldToQuery ORDERS_CUSTOMERID
				objDBInterface.AddFieldToQuery ORDERS_EMPLOYEEID
				objDBInterface.AddFieldToQuery ORDERS_ORDERDATE
				objDBInterface.AddFieldToQuery ORDERS_PURCHASEORDERNUMBER
				objDBInterface.AddFieldToQuery ORDERS_INVOICEDESCRIPTION
				objDBInterface.AddFieldToQuery ORDERS_SHIPNAME
				objDBInterface.AddFieldToQuery ORDERS_SHIPADDRESS
				objDBInterface.AddFieldToQuery ORDERS_SHIPCITY
				objDBInterface.AddFieldToQuery ORDERS_SHIPSTATEORPROVINCE
				objDBInterface.AddFieldToQuery ORDERS_SHIPPOSTCODE
				objDBInterface.AddFieldToQuery ORDERS_SHIPCOUNTRY
				objDBInterface.AddFieldToQuery ORDERS_SHIPPHONENUMBER
				objDBInterface.AddFieldToQuery ORDERS_SHIPDATE
				objDBInterface.AddFieldToQuery ORDERS_SHIPPINGMETHODID
				objDBInterface.AddFieldToQuery ORDERS_FREIGHTCHARGE
				objDBInterface.AddFieldToQuery ORDERS_SALESTAXRATE
				objDBInterface.AddFieldToQuery ORDERS_LEADREFERRALID
				objDBInterface.AddFieldToQuery ORDERS_PROPOSALREFERENCE
				objDBInterface.AddFieldToQuery ORDERS_ENTRYDATE
				objDBInterface.AddFieldToQuery ORDERS_COMMISSIONPAID
				objDBInterface.AddFieldToQuery ORDERS_INVOICEID
				objDBInterface.AddFieldToQuery ORDERS_UNIQUEKEY
				objDBInterface.AddFieldToQuery ORDERS_INSTALLCID
				objDBInterface.AddFieldToQuery ORDERS_RESELLERID
				objDBInterface.AddWhere ORDERS_ID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				' add this in, not too much overhead
				Set objDBOrderDetails = Nothing
				Set objDBOrderDetails = New CDBInterface
				objDBOrderDetails.Init arrCustDetails, nTotCustDetails	
		
				objDBOrderDetails.NewSQL
				objDBOrderDetails.AddFieldToQuery ORDERDETAILS_NEWZAPPCID
				objDBOrderDetails.AddWhere ORDERDETAILS_ORDERID, "=", objDBInterface.GetValue(ORDERS_ID), "number", "and"
				objDBOrderDetails.QueryDatabase
				
				nRecords = objDBOrderDetails.GetNumberOfRecords()
		
				If nRecords > 0 Then
					nCount = 0
					objDBOrderDetails.GetFirst()
					nOrderDetailsNewZappCID = objDBOrderDetails.GetValue(ORDERDETAILS_NEWZAPPCID)
					Do While nCount < nRecords		
							If nOrderDetailsNewZappCID <> objDBOrderDetails.GetValue(ORDERDETAILS_NEWZAPPCID) Then
								' We have a problem the order details are for more than 1 NewZappCID
								' better write this to a log file
								logSoapDebug "SYSTEM ERROR: Order: " & objDBInterface.GetValue(ORDERS_ID) & " - Contains Multiple NewZappCIDs within Order Details OR NewZappCID is Missing from some order details"
							End If
						nCount = nCount + 1
						objDBOrderDetails.GetNext()
					Loop
				End If
				
				If Len(nOrderDetailsNewZappCID)>0 Then
					objSoap.clearSendNodes()
					objSoap.AddSendNode "FunctionCall", "SendOrderXML"
					objSoap.AddSendNode "OrderID", objDBInterface.GetValue(ORDERS_ID)
					objSoap.AddSendNode "CustomerID", objDBInterface.GetValue(ORDERS_CUSTOMERID)
					objSoap.AddSendNode "NewZappCID", nOrderDetailsNewZappCID
					objSoap.AddSendNode "QuotationID", objDBInterface.GetValue(ORDERS_QUOTATIONID)
					objSoap.AddSendNode "EmployeeID", objDBInterface.GetValue(ORDERS_EMPLOYEEID)
					objSoap.AddSendNode "OrderDate", objDBInterface.GetValue(ORDERS_ORDERDATE)
					objSoap.AddSendNode "PurchaseOrderNumber", objDBInterface.GetValue(ORDERS_PURCHASEORDERNUMBER)
					objSoap.AddSendNode "InvoiceDescription", objDBInterface.GetValue(ORDERS_INVOICEDESCRIPTION)
					objSoap.AddSendNode "ShipName", objDBInterface.GetValue(ORDERS_SHIPNAME)
					objSoap.AddSendNode "ShipAddress", objDBInterface.GetValue(ORDERS_SHIPADDRESS)
					objSoap.AddSendNode "ShipCity", objDBInterface.GetValue(ORDERS_SHIPCITY)
					objSoap.AddSendNode "ShipStateOrProvince", objDBInterface.GetValue(ORDERS_SHIPSTATEORPROVINCE)
					objSoap.AddSendNode "ShipPostalCode", objDBInterface.GetValue(ORDERS_SHIPPOSTCODE)
					objSoap.AddSendNode "ShipCountry", objDBInterface.GetValue(ORDERS_SHIPCOUNTRY)
					objSoap.AddSendNode "ShipPhoneNumber", objDBInterface.GetValue(ORDERS_SHIPPHONENUMBER)
					objSoap.AddSendNode "ShipDate", objDBInterface.GetValue(ORDERS_SHIPDATE)
					objSoap.AddSendNode "ShippingMethodID", objDBInterface.GetValue(ORDERS_SHIPPINGMETHODID)
					objSoap.AddSendNode "FreightCharge", objDBInterface.GetValue(ORDERS_FREIGHTCHARGE)
					objSoap.AddSendNode "SalesTaxRate", objDBInterface.GetValue(ORDERS_SALESTAXRATE)
					objSoap.AddSendNode "LeadReferralID", objDBInterface.GetValue(ORDERS_LEADREFERRALID)
					objSoap.AddSendNode "ProposalReference", objDBInterface.GetValue(ORDERS_PROPOSALREFERENCE)
					objSoap.AddSendNode "EntryDate", objDBInterface.GetValue(ORDERS_ENTRYDATE)
					objSoap.AddSendNode "CommissionPaid", objDBInterface.GetValue(ORDERS_COMMISSIONPAID)
					objSoap.AddSendNode "InvoiceID", objDBInterface.GetValue(ORDERS_INVOICEID)
					objSoap.AddSendNode "InstallCID", objDBInterface.GetValue(ORDERS_INSTALLCID)
					objSoap.AddSendNode "ResellerID", objDBInterface.GetValue(ORDERS_RESELLERID)
				Else
					logSoapDebug "SYSTEM ERROR: No NewZappCID passed with order from order entry"
				End If
				Set objDBOrderDetails = Nothing

			Case 6
			
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery CUSTOMERS_ID
				objDBInterface.AddFieldToQuery CUSTOMERS_COMPANYNAME
				objDBInterface.AddFieldToQuery CUSTOMERS_BILLINGADDRESS
				objDBInterface.AddFieldToQuery CUSTOMERS_CITY
				objDBInterface.AddFieldToQuery CUSTOMERS_STATEORPROVINCE
				objDBInterface.AddFieldToQuery CUSTOMERS_POSTALCODE
				objDBInterface.AddFieldToQuery CUSTOMERS_COUNTRYID
				objDBInterface.AddFieldToQuery CUSTOMERS_PHONENUMBER
				objDBInterface.AddFieldToQuery CUSTOMERS_EMAILADDRESS
				objDBInterface.AddFieldToQuery CUSTOMERS_SHOWCUSTOMER
				objDBInterface.AddFieldToQuery CUSTOMERS_CREDITLIMIT
				'objDBInterface.AddFieldToQuery CUSTOMERS_NEWZAPPCID
				objDBInterface.AddWhere CUSTOMERS_ID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				objSoap.clearSendNodes()
				objSoap.AddSendNode "FunctionCall", "SendCustomerDetailsXML"
				objSoap.AddSendNode "CustomerID", objDBInterface.GetValue(CUSTOMERS_ID)
				objSoap.AddSendNode "CustomerName", objDBInterface.GetValue(CUSTOMERS_COMPANYNAME)
				objSoap.AddSendNode "Address1", objDBInterface.GetValue(CUSTOMERS_BILLINGADDRESS)
				objSoap.AddSendNode "City", objDBInterface.GetValue(CUSTOMERS_CITY)
				objSoap.AddSendNode "County", objDBInterface.GetValue(CUSTOMERS_STATEORPROVINCE)
				objSoap.AddSendNode "PostCode", objDBInterface.GetValue(CUSTOMERS_POSTALCODE)
				objSoap.AddSendNode "Country", objDBInterface.GetValue(CUSTOMERS_COUNTRYID)
				objSoap.AddSendNode "EmailAddress", objDBInterface.GetValue(CUSTOMERS_EMAILADDRESS)
				objSoap.AddSendNode "TelNo", objDBInterface.GetValue(CUSTOMERS_PHONENUMBER)
				objSoap.AddSendNode "CreditLimit", objDBInterface.GetValue(CUSTOMERS_CREDITLIMIT)
				'objSoap.AddSendNode "NewZappCID", objDBInterface.GetValue(CUSTOMERS_NEWZAPPCID)
				
				
				objDBInterface.NewSQL
				objDBInterface.AddFieldToQuery WEBHOSTING_CID
				objDBInterface.AddWhere WEBHOSTING_CUSTOMERID, "=", nOfflineID, "number", "and"
				objDBInterface.QueryDatabase
				
				nRecords = objDBInterface.GetNumberOfRecords()
				'response.write "nRecords: " & nRecords
				nCount = 0
				
				objSoap.AddOpenNode ("NewZappCIDs")
				If nRecords>0 Then
					Do While nCount < nRecords	
						'response.write "<br>REC: " & objDBInterface.GetValue(WEBHOSTING_CID)
						objSoap.AddSendNode "NewZappCID", objDBInterface.GetValue(WEBHOSTING_CID)
						
						nCount = nCount + 1
						objDBInterface.GetNext()
					Loop
				End If	
				objSoap.AddCloseNode ("NewZappCIDs")
			
			Case Else
				' do nothing
		End Select
		'response.write "<br>SOAP: " & objSoap.buildSendSoap
		sendSoapToUrl "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 
		
	End Sub
	
	Public Sub sendSoapToUrl(strURL)
		'response.write objSoap.buildSendSoap
		soapResponseCode 	= objSoap.SendSoap(objSoap.buildSendSoap, strURL)	' this returns 200 or 404 etc accordingly
		soapResponse 		= objSoap.getResponseText							' this returns the actual response as text
	
		'Response.write "RESPONSE TEXT: " & soapResponse	
	End Sub
	
	
	
	' GET NewZapp CID from OrderEntry
	' need to know online account this relates to!!!
	' is it better to look up form sync table????
	
	Public Sub LogOrder (nOrderID,nOrderDetailsID,strActionText,nActionType)
		Dim objSPConnEcom
		Dim objSPCmdEcom
		
		Set objSPConnEcom = Server.CreateObject("ADODB.Connection")
		Set objSPCmdEcom = Server.CreateObject("ADODB.Command")
		 
		objSPConnEcom.ConnectionString = strDB 
		objSPConnEcom.CursorLocation = adUseServer
		objSPConnEcom.open
		
		objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nOrderID", adInteger, adParamInput, , nOrderID)
		objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nOrderDetailID", adInteger, adParamInput, , nOrderDetailsID)
		objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("strAction", adVarChar, adParamInput, Len(strActionText), strActionText)
		objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nActionType", adInteger, adParamInput, , nActionType)
		
		objSPCmdEcom.ActiveConnection = objSPConnEcom
		objSPCmdEcom.CommandType = adCmdStoredProc
		objSPCmdEcom.CommandText="ECOM_AddLogEcommerce"
		objSPCmdEcom.Execute
		
		Set objSPCmdEcom = Nothing
		Set objSPConnEcom = Nothing
	End Sub
	
	Function AllowCID(nOfflineID,nTableID)
		
		AllowCID = False
		nOrderEntryCustomerID = 0
		
		If nTableID = 6 Then
			nOrderEntryCustomerID = nOfflineID
		End If
		
		If nTableID = 5 Then
			nOrderEntryCustomerID = getCustomerIDFromOrders(nOfflineID)
		End If
		
		If nTableID = 4 Then
			nOrderID = getOrderIDFromOrderDetails(nOfflineID)
			nOrderEntryCustomerID = getCustomerIDFromOrders(nOrderID)
		End If
		
		If nTableID = 3 Then
			nOrderID = getOrderIDFromPayments(nOfflineID)
			nOrderEntryCustomerID = getCustomerIDFromOrders(nOrderID)
		End If
		
		' HERE WE CAN ADD A CHECK 1 MANY ALL CIDS ETC ETC
		'If nOrderEntryCustomerID > 1419 Then ' 1420 to be the first customer allowed through.
			AllowCID = True
		'End If
		
	End Function
	
	Function getCustomerIDFromOrders(nOrderID)
		
		objDBInterface.NewSQL
		objDBInterface.AddFieldToQuery ORDERS_ID
		objDBInterface.AddFieldToQuery ORDERS_CUSTOMERID
		objDBInterface.AddWhere ORDERS_ID, "=", nOrderID, "number", "and"
		objDBInterface.QueryDatabase
		
		getCustomerIDFromOrders = objDBInterface.GetValue(ORDERS_CUSTOMERID)
	End Function
	
	Function getOrderIDFromOrderDetails(nOrderDetailsID)
		
		objDBInterface.NewSQL
		objDBInterface.AddFieldToQuery ORDERDETAILS_ID
		objDBInterface.AddFieldToQuery ORDERDETAILS_ORDERID
		objDBInterface.AddWhere ORDERDETAILS_ID, "=", nOrderDetailsID, "number", "and"
		objDBInterface.QueryDatabase
		
		getOrderIDFromOrderDetails = objDBInterface.GetValue(ORDERDETAILS_ORDERID)
	End Function
	
	Function getOrderIDFromPayments(nPaymentsID)
		
		objDBInterface.NewSQL
		objDBInterface.AddFieldToQuery PAYMENTS_ID
		objDBInterface.AddFieldToQuery PAYMENTS_ORDERID
		objDBInterface.AddWhere PAYMENTS_ID, "=", nPaymentsID, "number", "and"
		objDBInterface.QueryDatabase
		
		getOrderIDFromPayments = objDBInterface.GetValue(PAYMENTS_ORDERID)
	End Function
	
End Class
%>
