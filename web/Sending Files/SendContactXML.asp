<%@ Language=VBScript %>
<%

'nID = 88888888
strFunctionCall = "SendContactXML"
nCID = 1117
nNewZappCID = 2067
strEmailAddress = "jmkb009@aol.com"
strFirstName = "Justin"
strLastName = "Mortimore"
strJobTitle = "Work horse"
strMobileNo = "999"
bMainContact = 1
bAccountsContact = 0
strExtension = "Blah blah"
strUsername = "fdoin"
strPassword = "sdfoj"
strWebSite = "www.ebay.co.uk"
bNewsletter = 0


Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddContactSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<ID>"&nID&"</ID>"
  sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
  sendReq = sendReq & "<CID>"&nCID&"</CID>"
  sendReq = sendReq & "<NewZappCID>"&nNewZappCID&"</NewZappCID>"
  sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
  sendReq = sendReq & "<FirstName>"&strFirstName&"</FirstName>"
  sendReq = sendReq & "<LastName>"&strLastName&"</LastName>"
  sendReq = sendReq & "<JobTitle>"&strJobTitle&"</JobTitle>"
  sendReq = sendReq & "<MobileNo>"&strMobileNo&"</MobileNo>"
  sendReq = sendReq & "<MainContact>"&bMainContact&"</MainContact>"
  sendReq = sendReq & "<AccountsContact>"&bAccountsContact&"</AccountsContact>"
  sendReq = sendReq & "<Extension>"&strExtension&"</Extension>"
  sendReq = sendReq & "<Username>"&strUsername&"</Username>"
  sendReq = sendReq & "<Password>"&strPassword&"</Password>"
  sendReq = sendReq & "<WebSite>"&strWebSite&"</WebSite>"
  sendReq = sendReq & "<Newsletter>"&bNewsletter&"</Newsletter>"
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