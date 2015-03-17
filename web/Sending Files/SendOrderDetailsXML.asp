<%@ Language=VBScript %>
<%

nID = 11183
nOrderID = 6913
nProductID = 555
nQuantity = 9
nUnitPrice = 101
nDiscount = 16
nAccountType = 24
nCustomerID = 1117
nCID = 2067
dRenewalDate = Date()
bStatistics = 0
bFrontPage = 0
bServerSideIncludes = 0
bSSL = 0
bIndexServer = 0
bCurrent = 1

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddOrderDetailsSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<ID>"&nID&"</ID>"
  sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
  sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
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