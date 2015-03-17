<%@ Language=VBScript %>
<%

nID = 7508
strFunctionCall = "SendCustomerDetailsXML"
nResellerID = 0
nCustomerID = 2467
nOrderEntryCID = 1170
strCustomerName = "Darren C"
strCompany = "NA"
strAddress1 = "NA"
strAddress2 = "NA"
strCity = "NA"
strCounty = "NA"
strPostCode = "NA"
strCountry = "NA"
strEmailAddress = "dcurnow@destinet.co.uk"
strTelNo = "999"
strNewsletterTitle = "NA"
strNewsletterEmail = "NA"
dStartDate = "06/01/2006 13:00:00"
dLastDate = Now()
nTotalNewsSent = 15
dLastUsed = Now()
nTimesLoggedIn = 1121
bActive = 1
bDemo = 0 
strHashKey = "NA"
strMTA = "NA"
nCreditLimit = 0

Response.Write "Testing"

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<FunctionCall>"&strFunctionCall&"</FunctionCall>"
  sendReq = sendReq & "<ID>"&nID&"</ID>"
  sendReq = sendReq & "<ResellerID>"&nResellerID&"</ResellerID>"
  sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
  sendReq = sendReq & "<OrderEntryCID>"&nOrderEntryCID&"</OrderEntryCID>"
  sendReq = sendReq & "<CustomerName>"&strCustomerName&"</CustomerName>"
  sendReq = sendReq & "<Company>"&strCompany&"</Company>"
  sendReq = sendReq & "<Address1>"&strAddress1&"</Address1>"
  sendReq = sendReq & "<Address2>"&strAddress2&"</Address2>"
  sendReq = sendReq & "<City>"&strCity&"</City>"
  sendReq = sendReq & "<County>"&strCounty&"</County>"
  sendReq = sendReq & "<PostCode>"&strPostCode&"</PostCode>"
  sendReq = sendReq & "<Country>"&strCountry&"</Country>"
  sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
  sendReq = sendReq & "<TelNo>"&strTelNo&"</TelNo>"
  sendReq = sendReq & "<NewsletterTitle>"&strNewsletterTitle&"</NewsletterTitle>"
  sendReq = sendReq & "<NewsletterEmail>"&strNewsletterEmail&"</NewsletterEmail>"
  sendReq = sendReq & "<StartDate>"&dStartDate&"</StartDate>"
  sendReq = sendReq & "<LastDate>"&dLastDate&"</LastDate>"
  sendReq = sendReq & "<TotalNewsSent>"&nTotalNewsSent&"</TotalNewsSent>"
  sendReq = sendReq & "<LastUsed>"&dLastUsed&"</LastUsed>"
  sendReq = sendReq & "<TimesLoggedIn>"&nTimesLoggedIn&"</TimesLoggedIn>"
  sendReq = sendReq & "<Active>"&bActive&"</Active>"
  sendReq = sendReq & "<Demo>"&bDemo&"</Demo>"
  sendReq = sendReq & "<HashKey>"&strHashKey&"</HashKey>"
  sendReq = sendReq & "<MTA>"&strMTA&"</MTA>"
  sendReq = sendReq & "<CreditLimit>"&nCreditLimit&"</CreditLimit>"
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