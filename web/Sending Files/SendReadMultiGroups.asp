<%@ Language=VBScript %>
<%

    nCID = 2067
    strNameOfUser = "adminjustin"
    strPassword = "xmltest" 
    strField = "jmkb007885@aol.com"
    strGroup = "AOL,IMPORTED,test"
    'strGroup = "AOL,test"
    strFirstName = "Justin"
    strLastName = "Mortimore"
    strTitle = "Mr"
    strMobileNo = "999"
    strCompanyName = "Dnet"
    nEmailType = 0
    dEditDate = ""

    Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
    Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

    'Const SoapServerURL = "https://system.newzapp.co.uk/SOAP/ThankyouReSubscribeAPI.asp" 
    Const SoapServerURL = "https://www02.newzapp.co.uk/SOAP/ThankyouSubscribeToGroupAPI-Multi.asp" 
    'Const SoapServerURL = "http://dev210.newzapp.co.uk/test/responseSOAP.asp" 

    Function sendReq()
      sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
      sendReq = sendReq & "<SOAP:Body>"
      sendReq = sendReq & "<Request>"
      sendReq = sendReq & "<CID>"&nCID&"</CID>"
      sendReq = sendReq & "<UserName>"&strNameOfUser&"</UserName>"
      sendReq = sendReq & "<Password>"&strPassword&"</Password>"
      sendReq = sendReq & "<Email>"&strEmailAddress&"</Email>"
      sendReq = sendReq & "<FirstName>"&strFirstName&"</FirstName>"
      sendReq = sendReq & "<LastName>"&strLastName&"</LastName>"
      sendReq = sendReq & "<Title>"&strTitle&"</Title>"
      sendReq = sendReq & "<Mobile>"&strMobileNo&"</Mobile>"
      sendReq = sendReq & "<Company>"&strCompanyName&"</Company>"
      sendReq = sendReq & "<Group>"&strGroup&"</Group>"  
      sendReq = sendReq & "<EmailType>"&nEmailType&"</EmailType>"  
      sendReq = sendReq & "<EditDate>"&dEditDate&"</EditDate>"
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