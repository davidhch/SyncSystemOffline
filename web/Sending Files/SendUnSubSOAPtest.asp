<%@ Language=VBScript %>
<%

nCID = 2067 'Use your CID
strNameOfUser = "Your NewZapp username"
strPassword = "Your API password" 
strEmailAddress = "Email address you want to remove"


Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "https://system.newzapp.co.uk/SOAP/ThankyouRemoveAPI.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<CID>"&nCID&"</CID>"
  sendReq = sendReq & "<UserName>"&strNameOfUser&"</UserName>"
  sendReq = sendReq & "<Password>"&strPassword&"</Password>"
  sendReq = sendReq & "<DeleteEmail>"&strEmailAddress&"</DeleteEmail>"
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