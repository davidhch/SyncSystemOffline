<%@ Language=VBScript %>
<%

'nProductID = 578
strProductName = "Justin XML test 3"
nProductCategoryID = 16
nUnitPrice = 23
nUnitCost = 5
strOneOffCost = "Working" 
nSupplierID = 45
strDescription = "dfgjkdf"
strSystemDescription = "fbiohfrejio"
nSetupPrice = 56
nOrder = 981
bActive = 0

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddProductSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
  sendReq = sendReq & "<ProductName>"&strProductName&"</ProductName>"
  sendReq = sendReq & "<ProductCategoryID>"&nProductCategoryID&"</ProductCategoryID>"
  sendReq = sendReq & "<UnitPrice>"&nUnitPrice&"</UnitPrice>"
  sendReq = sendReq & "<UnitCost>"&nUnitCost&"</UnitCost>"
  sendReq = sendReq & "<OneOffCost>"&strOneOffCost&"</OneOffCost>"
  sendReq = sendReq & "<SupplierID>"&nSupplierID&"</SupplierID>"
  sendReq = sendReq & "<Description>"&strDescription&"</Description>"
  sendReq = sendReq & "<SystemDescription>"&strSystemDescription&"</SystemDescription>"
  sendReq = sendReq & "<SetupPrice>"&nSetupPrice&"</SetupPrice>"
  sendReq = sendReq & "<Order>"&nOrder&"</Order>"
  sendReq = sendReq & "<Active>"&bActive&"</Active>"
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

If objxmlhttp.Status = 200 Then
  Set objxmldom = objxmlhttp.responseXML
End If

Set objxmlhttp = Nothing
Set objxmldom = Nothing

%>