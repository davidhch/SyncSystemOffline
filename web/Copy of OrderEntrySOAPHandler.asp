<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->
<%
'actions this file needs to preform
'collect data when an action is sent to it
'identify where the data has come form
'run the required function
'save data to its host db
'send data to other handler
'preform db lookup for any items that are not in sync
	
	Response.Buffer = True 
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	strDB = strOEDB
	strHandlerName = "OrderEntryHandler"
	'Response.Write strHandlerName
	
	On Error Resume Next

	strFunctionCall = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FunctionCall").Text

	nID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ID").Text
	nResellerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ResellerID").Text
	nCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nOrderEntryCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryCID").Text
	strCustomerName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerName").Text
	strCompany = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Company").Text
	strAddress1 = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Address1").Text
	strAddress2 = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Address2").Text
	strCity = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/City").Text
	strCounty = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/County").Text 
	strPostCode = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PostCode").Text 
	strCountry = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Country").Text
	strEmailAddress = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/EmailAddress").Text
	strTelNo = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TelNo").Text
	strNewsletterTitle = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewsletterTitle").Text
	strNewsletterEmail = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewsletterEmail").Text
	dLastDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/LastDate").Text
	nTotalNewsSent = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TotalNewsSent").Text
	bActive = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Active").Text
	nTimesLoggedIn = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TimesLoggedIn").Text
	bDemo = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Demo").Text
	strHashKey = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/HashKey").Text
	strMTA = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/MTA").Text
	nCreditLimit = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CreditLimit").Text
	nNewZappCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappCID").Text
	
	If strFunctionCall = "SendCustomerDetailsXML" Then
		Call SendCustomerDetailsXML(nResellerID,nCustomerID,nOrderEntryCID,strCustomerName,strCompany,strAddress1,strAddress2,strCity,strCounty,strPostCode,strCountry,strEmailAddress,strTelNo,strNewsletterTitle,strNewsletterEmail,dLastDate,nTotalNewsSent,bActive,nTimesLoggedIn,bDemo,strHashKey,strMTA,nCreditLimit,nNewZappCID)
	End If
	
	Function SendCustomerDetailsXML(nResellerID,nCustomerID,nOrderEntryCID,strCustomerName,strCompany,strAddress1,strAddress2,strCity,strCounty,strPostCode,strCountry,strEmailAddress,strTelNo,strNewsletterTitle,strNewsletterEmail,dLastDate,nTotalNewsSent,bActive,nTimesLoggedIn,bDemo,strHashKey,strMTA,nCreditLimit,nNewZappCID)
		
		'Response.Write "<br>Inside correct function"
		
		nTableID = 6
		nOfflineID = nCustomerID
		'nOnlineID = nNewZappCID	
		bLocked = 1
		bExecuted = 0

		Call AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		
		'Response.Write "<br>Inside correct function after call!"
		
		'Response.Write "<br>******************nNewZappCID: " & nNewZappCID
		
		'Response.Write "<br>Coming back after function"
		
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)

		'Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddCustomerDetailsSOAP.asp" 
		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddCustomerDetailsSOAPOO.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<ID>"&nID&"</ID>"
		sendReq = sendReq & "<ResellerID>"&nResellerID&"</ResellerID>"
		sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
		sendReq = sendReq & "<OrderEntryCID>"&nOrderEntryCID&"</OrderEntryCID>"
		sendReq = sendReq & "<CustomerName>"&strCustomerName&"</CustomerName>"
		sendReq = sendReq & "<Company>"&strCompany&"</Company>"
		sendReq = sendReq & "<Address1>"&strAddress1&"</Address1>"
		sendReq = sendReq & "<Address2>"&strAddress2&"</Address2>"
		sendReq = sendReq & "<City>"&strCity&"</City>"
		sendReq = sendReq & "<County>"&strCounty&"</County>"
		sendReq = sendReq & "<PostCode>"&strPostCode&"</PostCode>"
		sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
		sendReq = sendReq & "<TelNo>"&strTelNo&"</TelNo>"
		sendReq = sendReq & "<NewsletterTitle>"&strNewsletterTitle&"</NewsletterTitle>"
		sendReq = sendReq & "<NewsletterEmail>"&strNewsletterEmail&"</NewsletterEmail>"
		sendReq = sendReq & "<StartDate>"&dStartDate&"</StartDate>"
		sendReq = sendReq & "<LastDate>"&dLastDate&"</LastDate>"
		sendReq = sendReq & "<TotalNewsSent>"&nTotalNewsSent&"</TotalNewsSent>"
		sendReq = sendReq & "<LastUsed>"&dLastUsed&"</LastUsed>"
		sendReq = sendReq & "<TimesLoggedIn>"&nTimesLoggedIn&"</TimesLoggedIn>"
		sendReq = sendReq & "<Active>"&bActive&"</Active>"
		sendReq = sendReq & "<Demo>"&bDemo&"</Demo>"
		sendReq = sendReq & "<HashKey>"&strHashKey&"</HashKey>"
		sendReq = sendReq & "<MTA>"&strMTA&"</MTA>"
		sendReq = sendReq & "<NewZappCID>"&nNewZappCID&"</NewZappCID>"
		sendReq = sendReq & "<CreditLimit>"&nCreditLimit&"</CreditLimit>"
		sendReq = sendReq & "<OnlineID>"&nOnlineID&"</OnlineID>"
		sendReq = sendReq & "<OfflineID>"&nOfflineID&"</OfflineID>"
		sendReq = sendReq & "<TableID>"&nTableID&"</TableID>"
		sendReq = sendReq & "<Locked>"&bLocked&"</Locked>"
		sendReq = sendReq & "<Executed>"&bExecuted&"</Executed>"
		sendReq = sendReq & "</Request>"
		sendReq = sendReq & "</SOAP:Body>"
		sendReq = sendReq & "</SOAP:Envelope>"	
		
		'Response.Write "<br>sendReq: " & sendReq

		objxmlhttp.open "POST", SoapServerURL, False
		objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
		objxmlhttp.setRequestHeader "MessageType", "CALL"
		objxmlhttp.setRequestHeader "Content-Type", "text/xml"
		objxmlhttp.send(sendReq)

		Response.write objxmlhttp.responseText

		Set objxmlhttp = Nothing
		Set objxmldom = Nothing
		
	End Function
	
	nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
	nCustomerID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nOrderEntryCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryCID").Text
	nNewZappCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappCID").Text
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
	
	If strFunctionCall = "SendOrderXML" Then
		Call SendOrderXML(nOrderID,nCustomerID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,nNewZappCID)
	End If
	
	Function SendOrderXML(nOrderID,nCustomerID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,nNewZappCID)
		
		nTableID = 5
		nOfflineID = nOrderID
		nOnlineID = nNewZappOrderID	
		bLocked = 1
		bExecuted = 0
		
		'Response.Write "<br>nTableID: " & nTableID
		'Response.Write "<br>nOfflineID: " & nOfflineID
		'Response.Write "<br>nOnlineID: " & nOnlineID
		'bDebug = True
		
		Call AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		'Response.Write "HERE 1:"

		objxmldom.load(Request)

		'Response.Write "HERE 2:"

		'Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddOrderSOAP.asp" 
		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddOrderSOAPOO.asp"

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
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
		sendReq = sendReq & "<OnlineID>"&nOnlineID&"</OnlineID>"
		sendReq = sendReq & "<OfflineID>"&nOfflineID&"</OfflineID>"
		sendReq = sendReq & "<TableID>"&nTableID&"</TableID>"
		sendReq = sendReq & "<Locked>"&bLocked&"</Locked>"
		sendReq = sendReq & "<Executed>"&bExecuted&"</Executed>"
		sendReq = sendReq & "</Request>"
		sendReq = sendReq & "</SOAP:Body>"
		sendReq = sendReq & "</SOAP:Envelope>"	

		'Response.Write "HERE 3:"

		objxmlhttp.open "POST", SoapServerURL, False
		objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
		objxmlhttp.setRequestHeader "MessageType", "CALL"
		objxmlhttp.setRequestHeader "Content-Type", "text/xml"
		objxmlhttp.send(sendReq)

		'Response.Write "HERE 4:"

		Response.write objxmlhttp.responseText

		'Response.Write "END: " & nTableID
		
		Set objxmlhttp = Nothing
		Set objxmldom = Nothing

	End Function
	
	nID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ID").Text
	nOrderDetailID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderDetailID").Text
	nNewZappOrderDetailID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappOrderDetailID").Text
	nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
	nNewZappOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappOrderID").Text
	nProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductID").Text
	nNewZappProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappProductID").Text
	nQuantity = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Quantity").Text
	nUnitPrice = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UnitPrice").Text
	nDiscount = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Discount").Text
	nAccountType = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/AccountType").Text
	nOrderEntryCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CustomerID").Text
	nCID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/CID").Text
	dRenewalDate = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/RenewalDate").Text
	bStatistics = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Statistics").Text
	bFrontPage = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/FrontPage").Text
	bServerSideIncludes = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ServerSideIncludes").Text
	bSSL = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/SSL").Text
	bIndexServer = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/IndexServer").Text
	bCurrent = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Current").Text
	
	'Response.Write "HERE 2:" & strFunctionCall
	
	If strFunctionCall = "SendOrderDetailsXML" Then
		'Response.Write "HERE 3:"
		Call SendOrderDetailsXML(nID,nOrderDetailID,nNewZappOrderDetailID,nOrderID,nNewZappOrderID,nProductID,nNewZappProductID,nQuantity,nUnitPrice,nDiscount,nAccountType,nOrderEntryCID,nCID,dRenewalDate,bStatistics,bFrontPage,bServerSideIncludes,bSSL,bIndexServer,bCurrent)
	End If

	Function SendOrderDetailsXML(nID,nOrderDetailID,nNewZappOrderDetailID,nOrderID,nNewZappOrderID,nProductID,nNewZappProductID,nQuantity,nUnitPrice,nDiscount,nAccountType,nOrderEntryCID,nCID,dRenewalDate,bStatistics,bFrontPage,bServerSideIncludes,bSSL,bIndexServer,bCurrent)
		
		'Response.Write "HERE 4:"
		'bDebug = True
		
		
		nTableID = 4
		nOfflineID = nOrderDetailID
		nOnlineID = nNewZappOrderDetailID	
		bLocked = 1
		bExecuted = 0
		
		'Response.Write "**nTableID: " & nTableID
		'Response.Write "**nOfflineID: " & nOfflineID
		'Response.Write "**nOnlineID: " & nOnlineID
		'Response.Write "**bLocked: " & bLocked
		'Response.Write "**bExecuted: " & bExecuted

		Call AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)

		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)

		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddOrderDetailsSOAPOO.asp" 
		'Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddOrderDetailsSOAP.asp" 
	
		

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<ID>"&nID&"</ID>"
		sendReq = sendReq & "<OrderDetailID>"&nOrderDetailID&"</OrderDetailID>"
		sendReq = sendReq & "<NewZappOrderDetailID>"&nNewZappOrderDetailID&"</NewZappOrderDetailID>"
		sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
		sendReq = sendReq & "<NewZappOrderID>"&nNewZappOrderID&"</NewZappOrderID>"
		sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
		sendReq = sendReq & "<NewZappProductID>"&nNewZappProductID&"</NewZappProductID>"
		sendReq = sendReq & "<Quantity>"&nQuantity&"</Quantity>"
		sendReq = sendReq & "<UnitPrice>"&nUnitPrice&"</UnitPrice>"
		sendReq = sendReq & "<Discount>"&nDiscount&"</Discount>"
		sendReq = sendReq & "<AccountType>"&nAccountType&"</AccountType>"
		sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
		sendReq = sendReq & "<CID>"&nCID&"</CID>"
		sendReq = sendReq & "<RenewalDate>"&dRenewalDate&"</RenewalDate>"
		sendReq = sendReq & "<Statistics>"&bStatistics&"</Statistics>"
		sendReq = sendReq & "<FrontPage>"&bFrontPage&"</FrontPage>"
		sendReq = sendReq & "<ServerSideIncludes>"&bServerSideIncludes&"</ServerSideIncludes>"
		sendReq = sendReq & "<SSL>"&bSSL&"</SSL>"
		sendReq = sendReq & "<IndexServer>"&bIndexServer&"</IndexServer>"
		sendReq = sendReq & "<Current>"&bCurrent&"</Current>"
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
		
	End Function
	
	nPaymentsID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/PaymentsID").Text
	nOrderEntryPaymentsID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryPaymentsID").Text
	nNewZappOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/NewZappOrderID").Text
	nOnlinePaymentID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlinePaymentID").Text
	nOrderID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderID").Text
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
	
	
	If strFunctionCall = "SendPaymentXML" Then
		Call SendPaymentXML(nPaymentsID,nOrderEntryPaymentsID,nNewZappOrderID,nOnlinePaymentID,nOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)
	End If

	Function SendPaymentXML(nPaymentsID,nOrderEntryPaymentsID,nNewZappOrderID,nOnlinePaymentID,nOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge)

		nTableID = 3
		nOfflineID = nPaymentsID
		nOnlineID = nOrderEntryPaymentsID
		bLocked = 1
		bExecuted = 0
		
		Call AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
		
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)

		'Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddPaymentSOAP.asp" 
		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddPaymentSOAPOO.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<PaymentsID>"&nPaymentsID&"</PaymentsID>"
		sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
		sendReq = sendReq & "<NewZappOrderID>"&nNewZappOrderID&"</NewZappOrderID>"
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
	
	nProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductID").Text
	strProductName = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductName").Text
	nOrderEntryProductID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OrderEntryProductID").Text
	nProductType = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/ProductType").Text
	nUsageLimit = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/UsageLimit").Text
	nCost = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Cost").Text
	bActive = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Active").Text
	
	'Response.Write "<br>nUsageLimit: " & nUsageLimit
	'Response.Write "<br>nCost: " & nCost

	If strFunctionCall = "SendProductXML" Then
		Call SendProductXML(nProductID,strProductName,nOrderEntryProductID,nProductType,nUsageLimit,nCost,bActive)
	End If

	Function SendProductXML(nProductID,strProductName,nOrderEntryProductID,nProductType,nUsageLimit,nCost,bActive)
	
		nTableID = 2
		nOfflineID = nOrderEntryProductID
		nOnlineID = nProductID	
		bLocked = 1
		bExecuted = 0
		
		Call AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
		
		objxmldom.load(Request)

		Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddProductSOAP.asp" 

		sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		sendReq = sendReq & "<SOAP:Body>"
		sendReq = sendReq & "<Request>"
		sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
		sendReq = sendReq & "<ProductName>"&strProductName&"</ProductName>"
		sendReq = sendReq & "<ProductType>"&nProductType&"</ProductType>"
		sendReq = sendReq & "<UsageLimit>"&nUsageLimit&"</UsageLimit>"
		sendReq = sendReq & "<Cost>"&nCost&"</Cost>"
		sendReq = sendReq & "<Active>"&bActive&"</Active>"
		sendReq = sendReq & "<OrderEntryProductID>"&nOrderEntryProductID&"</OrderEntryProductID>"
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

		If objxmlhttp.Status = 200 Then
		  Set objxmldom = objxmlhttp.responseXML
		End If

		Set objxmlhttp = Nothing
		Set objxmldom = Nothing
		
	End Function
	
	Function AddSyncLookup(nOnlineID,nOfflineID,nTableID,bLocked,bExecuted)
	
		Set objAddSyncLookup = Nothing
		Set objAddSyncLookup = New CDBInterface
		objAddSyncLookup.Init arrCustDetails, nTotCustDetails
		
		objAddSyncLookup.NewSQL
		objAddSyncLookup.AddFieldToQuery SYNCLOOKUP_ID
		objAddSyncLookup.AddWhere SYNCLOOKUP_OFFLINEID, "=", nOfflineID, "number", "and"
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
	
	End Function
	
	Set objxmldom = Nothing
	Set objxmlhttp = Nothing
		

%>
