<%@ Language=VBScript %>
<%

strFunctionCall = "SendOrderXML"
nCustomerID = 2067
nOrderEntryCID = 1117
nQuotationID = 8
nEmployeeID = 1
dOrderDate = Date()
strPurchaseOrderNumber = "JJM date test"
strInvoiceDescription = "NA"
strShipName = "NA" 
strShipAddress = "NA" 
strShipCity = "NA"
strShipStateOrProvince = "NA"
strShipPostalCode = "NA"
strShipCountry = "NA"
strShipPhoneNumber = "NA"
dShipDate = Date()
nShippingMethodID = 1
nFreightCharge = 0
nSalesTaxRate = 0.75
nLeadReferralID = 1
strProposalReference = "NA" 
dEntryDate = Now()
bCommissionPaid = 0
nInvoiceID = 1

Response.Write "<br>THIS IS DRIVING ME BONKERS!!!!!"

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 

Function sendReq()
  'sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  'sendReq = sendReq & "<SOAP:Body>"
  'sendReq = sendReq & "<Request>"
  'sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
  'sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
  'sendReq = sendReq & "<OrderEntryCID>"&nOrderEntryCID&"</OrderEntryCID>"
  'sendReq = sendReq & "<QuotationID>"&nQuotationID&"</QuotationID>"
  'sendReq = sendReq & "<EmployeeID>"&nEmployeeID&"</EmployeeID>"
  'sendReq = sendReq & "<OrderDate>"&dOrderDate&"</OrderDate>"
  'sendReq = sendReq & "<PurchaseOrderNumber>"&strPurchaseOrderNumber&"</PurchaseOrderNumber>"
  'sendReq = sendReq & "<InvoiceDescription>"&strInvoiceDescription&"</InvoiceDescription>"
  'sendReq = sendReq & "<ShipName>"&strShipName&"</ShipName>"
  'sendReq = sendReq & "<ShipAddress>"&strShipAddress&"</ShipAddress>"
  'sendReq = sendReq & "<ShipCity>"&strShipCity&"</ShipCity>"
 ' sendReq = sendReq & "<ShipStateOrProvince>"&strShipStateOrProvince&"</ShipStateOrProvince>"
  'sendReq = sendReq & "<ShipPostalCode>"&strShipPostalCode&"</ShipPostalCode>"
  'sendReq = sendReq & "<ShipCountry>"&strShipCountry&"</ShipCountry>"
  'sendReq = sendReq & "<ShipPhoneNumber>"&strShipPhoneNumber&"</ShipPhoneNumber>"
  'sendReq = sendReq & "<ShipDate>"&dShipDate&"</ShipDate>"
  'sendReq = sendReq & "<ShippingMethodID>"&nShippingMethodID&"</ShippingMethodID>"
  'sendReq = sendReq & "<FreightCharge>"&nFreightCharge&"</FreightCharge>"
  'sendReq = sendReq & "<SalesTaxRate>"&nSalesTaxRate&"</SalesTaxRate>"
  'sendReq = sendReq & "<LeadReferralID>"&nLeadReferralID&"</LeadReferralID>"
  'sendReq = sendReq & "<ProposalReference>"&strProposalReference&"</ProposalReference>"
  'sendReq = sendReq & "<EntryDate>"&dEntryDate&"</EntryDate>"
  'sendReq = sendReq & "<CommissionPaid>"&bCommissionPaid&"</CommissionPaid>"
  'sendReq = sendReq & "<ShowCustomer>"&bShowCustomer&"</ShowCustomer>"
  'sendReq = sendReq & "<InvoiceID>"&nInvoiceID&"</InvoiceID>"
  'sendReq = sendReq & "</Request>"
  'sendReq = sendReq & "</SOAP:Body>"
  'sendReq = sendReq & "</SOAP:Envelope>"	
   sendReq = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><FunctionCall>SendOrderXML</FunctionCall><CustomerID></CustomerID><NewZappCID>2467</NewZappCID><QuotationID></QuotationID><EmployeeID></EmployeeID><OrderDate>12/11/2007</OrderDate><PurchaseOrderNumber>21-9-15-2467-2007-11-12</PurchaseOrderNumber><InvoiceDescription>{1C35487C-942F-476D-9A8F-70955EAC0706}</InvoiceDescription><ShipName></ShipName><ShipAddress></ShipAddress><ShipCity></ShipCity><ShipStateOrProvince></ShipStateOrProvince><ShipPostalCode></ShipPostalCode><ShipCountry></ShipCountry><ShipPhoneNumber></ShipPhoneNumber><ShipDate></ShipDate><ShippingMethodID></ShippingMethodID><FreightCharge></FreightCharge><SalesTaxRate></SalesTaxRate><LeadReferralID></LeadReferralID><ProposalReference></ProposalReference><EntryDate></EntryDate><CommissionPaid></CommissionPaid><ShowCustomer></ShowCustomer><InvoiceID></InvoiceID></Request></SOAP:Body></SOAP:Envelope>"
End Function

objxmlhttp.open "POST", SoapServerURL, False
objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
objxmlhttp.setRequestHeader "MessageType", "CALL"
objxmlhttp.setRequestHeader "Content-Type", "text/xml"
objxmlhttp.send(sendReq)

Response.write objxmlhttp.responseText

Set objxmlhttp = Nothing
Set objxmldom = Nothing

%>