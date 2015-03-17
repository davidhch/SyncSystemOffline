<%@ Language=VBScript %>
<%

	nCID = 2067
	strNameOfUser = "adminjustin"
	strPassword = "xmltest"
	strEmailAddress = "jmkb007@aol.com"
	nCampID = 135949
	bRepeatEmails = 1
	bUnsubcribe = 1

    Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
    Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

    Const SoapServerURL = "https://system.newzapp.co.uk/SOAP/DoubleOptInEmailAPI.asp"
    
    Function sendReq()
      sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
      sendReq = sendReq & "<SOAP:Body>"
      sendReq = sendReq & "<Request>"
      sendReq = sendReq & "<CID>"&nCID&"</CID>"
      sendReq = sendReq & "<UserName>"&strNameOfUser&"</UserName>"
      sendReq = sendReq & "<Password>"&strPassword&"</Password>"
      sendReq = sendReq & "<SendEmail>"&strEmailAddress&"</SendEmail>"
      sendReq = sendReq & "<CampID>"&nCampID&"</CampID>"
      sendReq = sendReq & "<RepeatSends>"&bRepeatEmails&"</RepeatSends>"
      sendReq = sendReq & "<SendToUnsubscriber>"&bUnsubcribe&"</SendToUnsubscriber>"
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