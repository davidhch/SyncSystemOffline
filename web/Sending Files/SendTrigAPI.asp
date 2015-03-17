<%@ Language=VBScript %>
<%

    Response.Write ("test here")
    nCID = 2067
    adminjustin = "adminjustin"
    xmltest = "xmltest"
    CampID = 193464
    RepeatSends = True
    SendToUnsubscriber = False
    jmortimore = "jmortimore@destinet.co.uk"

    Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
    Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP")
    
    Const SoapServerURL = "http://nz-dev-net.destinet.co.uk/soap/TriggerCampaignAPI.aspx" 
    

    Function buildSendSoap()
        buildSendSoap = buildSendSoap & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
        buildSendSoap = buildSendSoap & "<SOAP:Body>"
        buildSendSoap = buildSendSoap & "<Request>"
        buildSendSoap = buildSendSoap & "<CID>" & nCID & "</CID>"
        buildSendSoap = buildSendSoap & "<UserName>" & adminjustin & "</UserName>"
        buildSendSoap = buildSendSoap & "<Password>" & xmltest & "</Password>"
        buildSendSoap = buildSendSoap & "<SendEmail>" & jmortimore & "</SendEmail>"
        buildSendSoap = buildSendSoap & "<CampID>" & CampID & "</CampID>"
        buildSendSoap = buildSendSoap & "<RepeatSends>" & RepeatSends & "</RepeatSends>"
        buildSendSoap = buildSendSoap & "<SendToUnsubscriber>" & SendToUnsubscriber & "</SendToUnsubscriber>"
        buildSendSoap = buildSendSoap & "<Validated></Validated>"
        buildSendSoap = buildSendSoap & "<FirstName></FirstName>"
        buildSendSoap = buildSendSoap & "<LastName></LastName>"
        buildSendSoap = buildSendSoap & "<Company></Company>"
        buildSendSoap = buildSendSoap & "<Group></Group>"
        buildSendSoap = buildSendSoap & "</Request>"
        buildSendSoap = buildSendSoap & "</SOAP:Body>"
        buildSendSoap = buildSendSoap & "</SOAP:Envelope>"
    End Function

    objxmlhttp.open "POST", SoapServerURL, False
    objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
    objxmlhttp.setRequestHeader "MessageType", "CALL"
    objxmlhttp.setRequestHeader "Content-Type", "text/xml"
    objxmlhttp.send(buildSendSoap)

    Response.write objxmlhttp.responseText

    Set objxmlhttp = Nothing
    Set objxmldom = Nothing 


%>