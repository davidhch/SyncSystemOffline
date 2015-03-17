<%@ Language=VBScript %>
<%


strFunctionCall = "SendProductXML"
nProductID = 32
nOrderEntryProductID = 101
strProductName = "Justin's Test Product Update"
nProductType = 19
nUsageLimit = 10
nCost = 6
bActive = 0

Response.Write "<br>nUsageLimit: " & nUsageLimit
Response.Write "<br>nCost: " & nCost

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
  sendReq = sendReq & "<ProductID>"&nProductID&"</ProductID>"
  sendReq = sendReq & "<ProductName>"&strProductName&"</ProductName>"
  sendReq = sendReq & "<OrderEntryProductID>"&nOrderEntryProductID&"</OrderEntryProductID>"
  sendReq = sendReq & "<ProductType>"&nProductType&"</ProductType>"
  sendReq = sendReq & "<UsageLimit>"&nUsageLimit&"</UsageLimit>"
  sendReq = sendReq & "<Cost>"&nCost&"</Cost>"
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