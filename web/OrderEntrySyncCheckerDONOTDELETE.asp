<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->
<%
	bDebug = True
	Response.Buffer = True 
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	strDB = strOEDB

	'Check native db for incomplete syncs
	Set objCheckSyncLookup = Nothing
	Set objCheckSyncLookup = New CDBInterface
	objCheckSyncLookup.Init arrCustDetails, nTotCustDetails
		
	objCheckSyncLookup.NewSQL
	objCheckSyncLookup.AddFieldToQuery SYNCLOOKUP_ID
	objCheckSyncLookup.AddFieldToQuery SYNCLOOKUP_ONLINEID
	objCheckSyncLookup.AddFieldToQuery SYNCLOOKUP_OFFLINEID
	objCheckSyncLookup.AddFieldToQuery SYNCLOOKUP_TABLEID
	objCheckSyncLookup.AddWhere SYNCLOOKUP_LOCKED, "=", 1, "number", "and"
	objCheckSyncLookup.AddWhere SYNCLOOKUP_EXECUTED, "=", 0, "number", "and"	
	objCheckSyncLookup.QueryDatabase
	
	nRecords = objCheckSyncLookup.GetNumberOfRecords()
	
	'loops through records
	If nRecords > 0 Then
		
		bRecordOutOfSync = True

		nCount = 0
		objCheckSyncLookup.GetFirst()
		
		Do While nCount < nRecords		
		
		nOnlineID = objCheckSyncLookup.GetValue(SYNCLOOKUP_ONLINEID)
		nOfflineID = objCheckSyncLookup.GetValue(SYNCLOOKUP_OFFLINEID)
		nTableID = objCheckSyncLookup.GetValue(SYNCLOOKUP_TABLEID)
	
		If nTableID = 2 Then
			strTableToUpdate = "Products"
			Call SyncProducts(nOnlineID,nOfflineID,nTableID)
		End If
	
		If nTableID = 3 Then
			strTableToUpdate = "Payments"
			'nOrderEntryPaymentsID = nOfflineID
			Call SyncPayments(nOnlineID,nOfflineID,nTableID)
		End If
	
		If nTableID = 4 Then
			strTableToUpdate = "OrderDetails"
			Call SyncOrderDetails(nOnlineID,nOfflineID,nTableID)
		End If
	
		If nTableID = 5 Then
			strTableToUpdate = "Orders"
			Call SyncOrders(nOnlineID,nOfflineID,nTableID)
		End If
	
		If nTableID = 6 Then
			strTableToUpdate = "Customers"
			Call SyncCustomers(nOnlineID,nOfflineID,nTableID)
		End If

		If nTableID = 8 Then
			strTableToUpdate = "Contacts"
			Call SyncContacts(nOnlineID,nOfflineID,nTableID)
		End If

		nCount = nCount + 1
		objCheckSyncLookup.GetNext()
		
		Loop

	End If		
	
	'functions to get local data and pass it to the soap local soap handler
	Function SyncProducts(nOnlineID,nOfflineID,nTableID)
	
		Set objGetProductData = Nothing
		Set objGetProductData = New CDBInterface
		objGetProductData.Init arrCustDetails, nTotCustDetails
		
		objGetProductData.NewSQL
		objGetProductData.AddFieldToQuery PRODUCTS_ID
		objGetProductData.AddFieldToQuery PRODUCTS_PRODUCTNAME
		objGetProductData.AddFieldToQuery PRODUCTS_PRODUCTCATEGORYID
		objGetProductData.AddFieldToQuery PRODUCTS_UNITPRICE
		objGetProductData.AddFieldToQuery PRODUCTS_UNITCOST
		objGetProductData.AddFieldToQuery PRODUCTS_ONEOFFCOST
		objGetProductData.AddFieldToQuery PRODUCTS_SUPPLIERID
		objGetProductData.AddFieldToQuery PRODUCTS_DESCRIPTION
		objGetProductData.AddFieldToQuery PRODUCTS_SYSTEMDESCRIPTION
		objGetProductData.AddFieldToQuery PRODUCTS_SETUPPRICE
		objGetProductData.AddFieldToQuery PRODUCTS_ORDER
		objGetProductData.AddFieldToQuery PRODUCTS_ACTIVE
		objGetProductData.AddWhere PRODUCTS_ID, "=", nOfflineID, "number", "and"
		objGetProductData.QueryDatabase
		
		nOrderEntryProductID = objGetProductData.GetValue(PRODUCTS_ID)
		strProductName = objGetProductData.GetValue(PRODUCTS_PRODUCTNAME)
		nProductCategoryID = objGetProductData.GetValue(PRODUCTS_PRODUCTCATEGORYID)
		nUnitPrice = objGetProductData.GetValue(PRODUCTS_UNITPRICE)
		nUnitCost = objGetProductData.GetValue(PRODUCTS_UNITCOST)
		nOneOffCost = objGetProductData.GetValue(PRODUCTS_ONEOFFCOST)
		nSupplierID = objGetProductData.GetValue(PRODUCTS_SUPPLIERID)
		strDescription = objGetProductData.GetValue(PRODUCTS_DESCRIPTION)
		strSystemDescription = objGetProductData.GetValue(PRODUCTS_SYSTEMDESCRIPTION)
		nSetupPrice = objGetProductData.GetValue(PRODUCTS_SETUPPRICE)
		nOrder = objGetProductData.GetValue(PRODUCTS_ORDER)
		bActive = objGetProductData.GetValue(PRODUCTS_ACTIVE)
		
		Call SendProductData(nOrderEntryProductID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,nOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive)
	
	End Function
	
	Function SyncPayments(nOnlineID,nOfflineID,nTableID)
	
		Set objGetPaymentData = Nothing
		Set objGetPaymentData = New CDBInterface
		objGetPaymentData.Init arrCustDetails, nTotCustDetails
		
		objGetPaymentData.NewSQL
		objGetPaymentData.AddFieldToQuery PAYMENTS_ID
		objGetPaymentData.AddFieldToQuery PAYMENTS_ORDERID
		objGetPaymentData.AddFieldToQuery PAYMENTS_PAYMENTAMOUNT
		objGetPaymentData.AddFieldToQuery PAYMENTS_PAYMENTDATE
		objGetPaymentData.AddFieldToQuery PAYMENTS_CREDITCARDID
		objGetPaymentData.AddFieldToQuery PAYMENTS_PAYMENTMETHODID
		objGetPaymentData.AddFieldToQuery PAYMENTS_TRANSACTIONCODE
		objGetPaymentData.AddFieldToQuery PAYMENTS_PROCESSINGVALID
		objGetPaymentData.AddFieldToQuery PAYMENTS_PROCESSINGCODE
		objGetPaymentData.AddFieldToQuery PAYMENTS_PROCESSINGCODEMESSAGE
		objGetPaymentData.AddFieldToQuery PAYMENTS_AUTHCODE
		objGetPaymentData.AddFieldToQuery PAYMENTS_AUTHMESSAGE
		objGetPaymentData.AddFieldToQuery PAYMENTS_RESPONSECODE
		objGetPaymentData.AddFieldToQuery PAYMENTS_RESPONSECODEMESSAGE
		objGetPaymentData.AddFieldToQuery PAYMENTS_AMOUNTTOCHARGE
		objGetPaymentData.AddWhere PAYMENTS_ID, "=", nOfflineID, "number", "and"
		objGetPaymentData.QueryDatabase
		
		nPaymentsID = objGetPaymentData.GetValue(PAYMENTS_ID)
		nOrderID = objGetPaymentData.GetValue(PAYMENTS_ORDERID)
		nPaymentAmount = objGetPaymentData.GetValue(PAYMENTS_PAYMENTAMOUNT)
		dPaymentDate = objGetPaymentData.GetValue(PAYMENTS_PAYMENTDATE)
		nCreditCardID = objGetPaymentData.GetValue(PAYMENTS_CREDITCARDID)
		nPaymentMethodID = objGetPaymentData.GetValue(PAYMENTS_PAYMENTMETHODID)
		strTransactionCode = objGetPaymentData.GetValue(PAYMENTS_TRANSACTIONCODE)
		strProcessingValid = objGetPaymentData.GetValue(PAYMENTS_PROCESSINGVALID)
		strProcessingCode = objGetPaymentData.GetValue(PAYMENTS_PROCESSINGCODE)
		strProcessingCodeMessage = objGetPaymentData.GetValue(PAYMENTS_PROCESSINGCODEMESSAGE)
		strAuthCode = objGetPaymentData.GetValue(PAYMENTS_AUTHCODE)
		strAuthMessage = objGetPaymentData.GetValue(PAYMENTS_AUTHMESSAGE)
		strResponseCode = objGetPaymentData.GetValue(PAYMENTS_RESPONSECODE)
		strResponseCodeMessage = objGetPaymentData.GetValue(PAYMENTS_RESPONSECODEMESSAGE)
		nAmountToCharge = objGetPaymentData.GetValue(PAYMENTS_AMOUNTTOCHARGE)
		
		Call SendPaymentData(nPaymentsID,nOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
	
	End Function
	
	Function SyncOrderDetails(nOnlineID,nOfflineID,nTableID)
		
		bDebug = True
	
		Set objGetOrderDetailsData = Nothing
		Set objGetOrderDetailsData = New CDBInterface
		objGetOrderDetailsData.Init arrCustDetails, nTotCustDetails
		
		objGetOrderDetailsData.NewSQL
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_ID
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_ORDERID
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_PRODUCTID
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_QUANTITY
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_UNITPRICE
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_DISCOUNT
		objGetOrderDetailsData.AddFieldToQuery ORDERDETAILS_UNIQUEKEY
		objGetOrderDetailsData.AddWhere ORDERDETAILS_ID, "=", nOfflineID, "number", "and"
		objGetOrderDetailsData.QueryDatabase
		
		nOrderDetailsID = objGetOrderDetailsData.GetValue(ORDERDETAILS_ID)
		nOrderID = objGetOrderDetailsData.GetValue(ORDERDETAILS_ORDERID)
		nProductID = objGetOrderDetailsData.GetValue(ORDERDETAILS_PRODUCTID)
		nQuantity = objGetOrderDetailsData.GetValue(ORDERDETAILS_QUANTITY)
		nUnitPrice = objGetOrderDetailsData.GetValue(ORDERDETAILS_UNITPRICE)
		nDiscount = objGetOrderDetailsData.GetValue(ORDERDETAILS_DISCOUNT)
		strUniqueKey = objGetOrderDetailsData.GetValue(ORDERDETAILS_UNIQUEKEY)
		
		Call SendOrderDetailsData(nOrderDetailsID,nOrderID,nProductID,nQuantity,nUnitPrice,nDiscount,strUniqueKey)
		
	End Function
	
	Function SyncOrders(nOnlineID,nOfflineID,nTableID)
		
		Set objGetOrdersData = Nothing
		Set objGetOrdersData = New CDBInterface
		objGetOrdersData.Init arrCustDetails, nTotCustDetails
		
		objGetOrdersData.NewSQL
		objGetOrdersData.AddFieldToQuery ORDERS_ID
		objGetOrdersData.AddFieldToQuery ORDERS_QUOTATIONID
		objGetOrdersData.AddFieldToQuery ORDERS_CUSTOMERID
		objGetOrdersData.AddFieldToQuery ORDERS_EMPLOYEEID
		objGetOrdersData.AddFieldToQuery ORDERS_ORDERDATE
		objGetOrdersData.AddFieldToQuery ORDERS_PURCHASEORDERNUMBER
		objGetOrdersData.AddFieldToQuery ORDERS_INVOICEDESCRIPTION
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPNAME
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPADDRESS
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPCITY
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPSTATEORPROVINCE
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPPOSTCODE
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPCOUNTRY
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPPHONENUMBER
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPDATE
		objGetOrdersData.AddFieldToQuery ORDERS_SHIPPINGMETHODID
		objGetOrdersData.AddFieldToQuery ORDERS_FREIGHTCHARGE
		objGetOrdersData.AddFieldToQuery ORDERS_SALESTAXRATE
		objGetOrdersData.AddFieldToQuery ORDERS_LEADREFERRALID
		objGetOrdersData.AddFieldToQuery ORDERS_PROPOSALREFERENCE
		objGetOrdersData.AddFieldToQuery ORDERS_ENTRYDATE
		objGetOrdersData.AddFieldToQuery ORDERS_COMMISSIONPAID
		'objGetOrdersData.AddFieldToQuery ORDERS_ACTIVE
		'objGetOrdersData.AddFieldToQuery ORDERS_EXPIREDATE
		objGetOrdersData.AddFieldToQuery ORDERS_INVOICEID
		objGetOrdersData.AddFieldToQuery ORDERS_UNIQUEKEY
		objGetOrdersData.AddWhere ORDERS_ID, "=", nOfflineID, "number", "and"
		objGetOrdersData.QueryDatabase
	
		nOrderID = objGetOrdersData.GetValue(ORDERS_ID)
		nQuotationID = objGetOrdersData.GetValue(ORDERS_QUOTATIONID)
		nCustomerID = objGetOrdersData.GetValue(ORDERS_CUSTOMERID)
		nEmployeeID = objGetOrdersData.GetValue(ORDERS_EMPLOYEEID)
		dOrderDate = objGetOrdersData.GetValue(ORDERS_ORDERDATE)
		strPurchaseOrderNumber = objGetOrdersData.GetValue(ORDERS_PURCHASEORDERNUMBER)
		strInvoiceDescription = objGetOrdersData.GetValue(ORDERS_INVOICEDESCRIPTION)
		strShipName = objGetOrdersData.GetValue(ORDERS_SHIPNAME)
		strShipAddress = objGetOrdersData.GetValue(ORDERS_SHIPADDRESS)
		strShipCity = objGetOrdersData.GetValue(ORDERS_SHIPCITY)
		strShipStateOrProvince = objGetOrdersData.GetValue(ORDERS_SHIPSTATEORPROVINCE)
		strShipPostalCode = objGetOrdersData.GetValue(ORDERS_SHIPPOSTCODE)
		strShipCountry = objGetOrdersData.GetValue(ORDERS_SHIPCOUNTRY)
		strShipPhoneNumber = objGetOrdersData.GetValue(ORDERS_SHIPPHONENUMBER)
		dShipDate = objGetOrdersData.GetValue(ORDERS_SHIPDATE)
		nShippingMethodID = objGetOrdersData.GetValue(ORDERS_SHIPPINGMETHODID)
		nFreightCharge = objGetOrdersData.GetValue(ORDERS_FREIGHTCHARGE)
		nSalesTaxRate = objGetOrdersData.GetValue(ORDERS_SALESTAXRATE)
		nLeadReferralID = objGetOrdersData.GetValue(ORDERS_LEADREFERRALID)
		strProposalReference = objGetOrdersData.GetValue(ORDERS_PROPOSALREFERENCE)
		dEntryDate = objGetOrdersData.GetValue(ORDERS_ENTRYDATE)
		bCommissionPaid = objGetOrdersData.GetValue(ORDERS_COMMISSIONPAID)
		'bActive = objGetOrdersData.GetValue(ORDERS_ACTIVE)
		'dExpireDate = objGetOrdersData.GetValue(ORDERS_EXPIREDATE)
		nInvoiceID = objGetOrdersData.GetValue(ORDERS_INVOICEID)
		nUniqueKey = objGetOrdersData.GetValue(ORDERS_UNIQUEKEY)
		nNewZappCID = 2067
		
		Call SendOrdersData(nOrderID,nQuotationID,nCustomerID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,bActive,dExpireDate,nInvoiceID,nUniqueKey,nNewZappCID)
	
	End Function
	
	Function SyncCustomers(nOnlineID,nOfflineID,nTableID)
		
		Response.Write "<br>Running correct function"
		
		Set objGetCustomerData = Nothing
		Set objGetCustomerData = New CDBInterface
		objGetCustomerData.Init arrCustDetails, nTotCustDetails
		
		objGetCustomerData.NewSQL
		objGetCustomerData.AddFieldToQuery CUSTOMERS_ID
		objGetCustomerData.AddFieldToQuery CUSTOMERS_COMPANYNAME
		objGetCustomerData.AddFieldToQuery CUSTOMERS_BILLINGADDRESS
		objGetCustomerData.AddFieldToQuery CUSTOMERS_CITY
		objGetCustomerData.AddFieldToQuery CUSTOMERS_STATEORPROVINCE
		objGetCustomerData.AddFieldToQuery CUSTOMERS_POSTALCODE
		objGetCustomerData.AddFieldToQuery CUSTOMERS_COUNTRYID
		objGetCustomerData.AddFieldToQuery CUSTOMERS_PHONENUMBER
		objGetCustomerData.AddFieldToQuery CUSTOMERS_EMAILADDRESS
		objGetCustomerData.AddFieldToQuery CUSTOMERS_SHOWCUSTOMER
		objGetCustomerData.AddFieldToQuery CUSTOMERS_CREDITLIMIT
		objGetCustomerData.AddFieldToQuery CUSTOMERS_NEWZAPPCID
		objGetCustomerData.AddWhere CUSTOMERS_ID, "=", nOfflineID, "number", "and"
		objGetCustomerData.QueryDatabase
		
		nCustomerID = objGetCustomerData.GetValue(CUSTOMERS_ID)
		strCustomerName = objGetCustomerData.GetValue(CUSTOMERS_COMPANYNAME)
		strAddress1 = objGetCustomerData.GetValue(CUSTOMERS_BILLINGADDRESS)
		strCity = objGetCustomerData.GetValue(CUSTOMERS_CITY)
		strCounty = objGetCustomerData.GetValue(CUSTOMERS_STATEORPROVINCE)
		strPostCode = objGetCustomerData.GetValue(CUSTOMERS_POSTALCODE)
		strCountry = objGetCustomerData.GetValue(CUSTOMERS_COUNTRYID)
		strTelNo = objGetCustomerData.GetValue(CUSTOMERS_PHONENUMBER)
		strEmailAddress = objGetCustomerData.GetValue(CUSTOMERS_EMAILADDRESS)
		bActive = objGetCustomerData.GetValue(CUSTOMERS_SHOWCUSTOMER)
		nCreditLimit = objGetCustomerData.GetValue(CUSTOMERS_CREDITLIMIT)
		nNewZappCID = objGetCustomerData.GetValue(CUSTOMERS_NEWZAPPCID)
		
		Call SendCustomerData(nCustomerID,strCustomerName,strAddress1,strCity,strCounty,strPostCode,strCountry,strTelNo,strEmailAddress,bActive,nCreditLimit,nNewZappCID)
	
	End Function
	
	Function SyncContacts(nOnlineID,nOfflineID,nTableID)
	
		Response.Write "<br>Running correct function again"
		
		'CONTACTS_ID
		'CONTACTS_CUSTOMERID
		'CONTACTS_FIRSTNAME
		'CONTACTS_LASTNAME
		'CONTACTS_JOBTITLE
		'CONTACTS_MAINCONTACT
		'CONTACTS_ACCOUNTCONTACT
		'CONTACTS_EXTENSION
		'CONTACTS_MOBILENO
		'CONTACTS_EMAILADDRESS
		'CONTACTS_USERNAME
		'CONTACTS_PASSWORD
		'CONTACTS_WEBSITE
		'CONTACTS_NEWSLETTER
		'CONTACTS_NEWZAPPCID
	
	End Function
	
	'Send data to the local SOAP Handler
	Function SendProductData(nOrderEntryProductID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,nOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive)	
	
		strFunctionCall = "SendProductXML"
	
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)
		
		Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 
		
		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
		sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
		sendReq = sendReq & "<ProductName>"&strProductName&"</ProductName>"
		sendReq = sendReq & "<OrderEntryProductID>"&nOrderEntryProductID&"</OrderEntryProductID>"
		sendReq = sendReq & "<ProductType>"&nProductCategoryID&"</ProductType>"
		sendReq = sendReq & "<UsageLimit>"&nUsageLimit&"</UsageLimit>"
		sendReq = sendReq & "<Cost>"&nUnitCost&"</Cost>"
		sendReq = sendReq & "<Active>"&bActive&"</Active>"
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
	
	Function SendPaymentData(nPaymentsID,nOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
		
		strFunctionCall = "SendPaymentXML"
	
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)
		
		Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 
		
		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
		sendReq = sendReq & "<PaymentsID>"&nPaymentsID&"</PaymentsID>"
		sendReq = sendReq & "<OrderEntryPaymentsID>"&nOrderEntryPaymentsID&"</OrderEntryPaymentsID>"
		sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
		sendReq = sendReq & "<PaymentAmount>"&nPaymentAmount&"</PaymentAmount>"
		sendReq = sendReq & "<PaymentDate>"&dPaymentDate&"</PaymentDate>"
		sendReq = sendReq & "<CreditCardID>"&nCreditCardID&"</CreditCardID>"
		sendReq = sendReq & "<PaymentMethodID>"&nPaymentMethodID&"</PaymentMethodID>"
		sendReq = sendReq & "<TransactionCode>"&strTransactionCode&"</TransactionCode>"
		sendReq = sendReq & "<ProcessingValid>"&strProcessingValid&"</ProcessingValid>"
		sendReq = sendReq & "<ProcessingCode>"&strProcessingCode&"</ProcessingCode>"
		sendReq = sendReq & "<ProcessingCodeMessage>"&strProcessingCodeMessage&"</ProcessingCodeMessage>"
		sendReq = sendReq & "<AuthCode>"&strAuthCode&"</AuthCode>"
		sendReq = sendReq & "<AuthMessage>"&strAuthMessage&"</AuthMessage>"
		sendReq = sendReq & "<ResponseCode>"&strResponseCode&"</ResponseCode>"
		sendReq = sendReq & "<ResponseCodeMessage>"&strResponseCodeMessage&"</ResponseCodeMessage>"
		sendReq = sendReq & "<AmountToCharge>"&nAmountToCharge&"</AmountToCharge>"
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
	
	Function SendOrderDetailsData(nOrderDetailsID,nOrderID,nProductID,nQuantity,nUnitPrice,nDiscount,strUniqueKey)
		
		strFunctionCall = "SendOrderDetailsXML"
		
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)

		Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 
	
		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
		sendReq = sendReq & "<ID>"&nOrderDetailsID&"</ID>"
		sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
		sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
		sendReq = sendReq & "<Quantity>"&nQuantity&"</Quantity>"
		sendReq = sendReq & "<UnitPrice>"&nUnitPrice&"</UnitPrice>"
		sendReq = sendReq & "<Discount>"&nDiscount&"</Discount>"
		sendReq = sendReq & "<UniqueKey>"&strUniqueKey&"</UniqueKey>"
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
	
	Function SendOrdersData(nOrderID,nQuotationID,nCustomerID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,bActive,dExpireDate,nInvoiceID,nUniqueKey,nNewZappCID)
	
		strFunctionCall = "SendOrderXML"
	
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

		Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
		sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
		sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
		sendReq = sendReq & "<NewZappCID>"&nNewZappCID&"</NewZappCID>"
		sendReq = sendReq & "<QuotationID>"&nQuotationID&"</QuotationID>"
		sendReq = sendReq & "<EmployeeID>"&nEmployeeID&"</EmployeeID>"
		sendReq = sendReq & "<OrderDate>"&dOrderDate&"</OrderDate>"
		sendReq = sendReq & "<PurchaseOrderNumber>"&strPurchaseOrderNumber&"</PurchaseOrderNumber>"
		sendReq = sendReq & "<InvoiceDescription>"&strInvoiceDescription&"</InvoiceDescription>"
		sendReq = sendReq & "<ShipName>"&strShipName&"</ShipName>"
		sendReq = sendReq & "<ShipAddress>"&strShipAddress&"</ShipAddress>"
		sendReq = sendReq & "<ShipCity>"&strShipCity&"</ShipCity>"
		sendReq = sendReq & "<ShipStateOrProvince>"&strShipStateOrProvince&"</ShipStateOrProvince>"
		sendReq = sendReq & "<ShipPostalCode>"&strShipPostalCode&"</ShipPostalCode>"
		sendReq = sendReq & "<ShipCountry>"&strShipCountry&"</ShipCountry>"
		sendReq = sendReq & "<ShipPhoneNumber>"&strShipPhoneNumber&"</ShipPhoneNumber>"
		sendReq = sendReq & "<ShipDate>"&dShipDate&"</ShipDate>"
		sendReq = sendReq & "<ShippingMethodID>"&nShippingMethodID&"</ShippingMethodID>"
		sendReq = sendReq & "<FreightCharge>"&nFreightCharge&"</FreightCharge>"
		sendReq = sendReq & "<SalesTaxRate>"&nSalesTaxRate&"</SalesTaxRate>"
		sendReq = sendReq & "<LeadReferralID>"&nLeadReferralID&"</LeadReferralID>"
		sendReq = sendReq & "<ProposalReference>"&strProposalReference&"</ProposalReference>"
		sendReq = sendReq & "<EntryDate>"&dEntryDate&"</EntryDate>"
		sendReq = sendReq & "<CommissionPaid>"&bCommissionPaid&"</CommissionPaid>"
		sendReq = sendReq & "<ShowCustomer>"&bShowCustomer&"</ShowCustomer>"
		sendReq = sendReq & "<InvoiceID>"&nInvoiceID&"</InvoiceID>"
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
	
	Function SendCustomerData(nCustomerID,strCustomerName,strAddress1,strCity,strCounty,strPostCode,strCountry,strTelNo,strEmailAddress,bActive,nCreditLimit,nNewZappCID)
	
		Response.Write "<br>nCustomerID: " & nCustomerID
		Response.Write "<br>strCustomerName: " & strCustomerName
		Response.Write "<br>strAddress1: " & strAddress1
		Response.Write "<br>strCity: " & strCity
		Response.Write "<br>strCounty: " & strCounty
		Response.Write "<br>strPostCode: " & strPostCode
		Response.Write "<br>strCountry: " & strCountry
		Response.Write "<br>strTelNo: " & strTelNo
		Response.Write "<br>strEmailAddress: " & strEmailAddress
		Response.Write "<br>bActive: " & bActive
		Response.Write "<br>nCreditLimit: " & nCreditLimit
		Response.Write "<br>nNewZappCID: " & nNewZappCID
		
		strFunctionCall = "SendCustomerDetailsXML"
	
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

		Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
		sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
		sendReq = sendReq & "<CustomerName>"&strCustomerName&"</CustomerName>"
		sendReq = sendReq & "<Address1>"&strAddress1&"</Address1>"
		sendReq = sendReq & "<City>"&strCity&"</City>"
		sendReq = sendReq & "<County>"&strCounty&"</County>"
		sendReq = sendReq & "<PostCode>"&strPostCode&"</PostCode>"
		sendReq = sendReq & "<Country>"&strCountry&"</Country>"
		sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
		sendReq = sendReq & "<TelNo>"&strTelNo&"</TelNo>"
		sendReq = sendReq & "<CreditLimit>"&nCreditLimit&"</CreditLimit>"
		sendReq = sendReq & "<NewZappCID>"&nNewZappCID&"</NewZappCID>"
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
	
	Set objxmldom = Nothing
	Set objxmlhttp = Nothing

%>