<%@ Language=VBScript %>
<%

nCustomerID = 1117
nNewZappCID = 2067
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
dShipDate = Now()
nShippingMethodID = 1
nFreightCharge = 0
nSalesTaxRate = 0.75
'nLeadReferralID = 
strProposalReference = "NA" 
dEntryDate = Now()
bCommissionPaid = 0
nInvoiceID = 1

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddOrderSOAP.asp" 

Function sendReq()
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
  'sendReq = sendReq & "<LeadReferralID>"&nLeadReferralID&"</LeadReferralID>"
  sendReq = sendReq & "<ProposalReference>"&strProposalReference&"</ProposalReference>"
  sendReq = sendReq & "<EntryDate>"&dEntryDate&"</EntryDate>"
  sendReq = sendReq & "<CommissionPaid>"&bCommissionPaid&"</CommissionPaid>"
  sendReq = sendReq & "<ShowCustomer>"&bShowCustomer&"</ShowCustomer>"
  sendReq = sendReq & "<InvoiceID>"&nInvoiceID&"</InvoiceID>"
  sendReq = sendReq & "</Request>"
  sendReq = sendReq & "</SOAP:Body>"
  sendReq = sendReq & "</SOAP:Envelope>"	
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