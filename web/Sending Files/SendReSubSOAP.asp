<%@ Language=VBScript %>
<%

nCID = 2067
strNameOfUser = "adminjustin"
strPassword = "xmltest" 
strEmailAddress = "jmkb007@aol.com"


Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

'Const SoapServerURL = "https://system.newzapp.co.uk/SOAP/ThankyouReSubscribeAPI.asp" 
Const SoapServerURL = "https://www01.newzapp.co.uk/SOAP/ThankyouReSubscribeAPI.asp" 
'Const SoapServerURL = "http://dev210.newzapp.co.uk/test/responseSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<CID>"&nCID&"</CID>"
  sendReq = sendReq & "<UserName>"&strNameOfUser&"</UserName>"
  sendReq = sendReq & "<Password>"&strPassword&"</Password>"
  sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
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