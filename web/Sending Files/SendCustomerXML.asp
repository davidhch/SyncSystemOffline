<%@ Language=VBScript %>
<%

nCustomerID = 1117
nNewZappCID = 2467
strCompanyName = "Js comp"
strBillingAddress = "BOOOOO!"
strCity = "NA"
strStateOrProvince = "NA"
strPostalCode = "NA" 
nCountryID = 0 
strContactTitle = "NA"
strPhoneNumber = "NA"
strFaxNumber = "NA"
strEmailAddress = "jmkb008@aol.com"
strWebAddress = "NA"
strNotes = "refevfr regvve ergbv"
strWebSiteUserName = "NA"
strWebSitePassword = "NA"
bLeadGenerator = 0
'nLeadReferralID = 0
nVATCode = 0
strVATNo = "NA"
bReseller = 1
bShowCustomer = 0
bCreditCardCustomer = 1
nCreditLimit = 5000

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddCustomerSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<CustomerID>"&nCustomerID&"</CustomerID>"
  sendReq = sendReq & "<NewZappCID>"&nNewZappCID&"</NewZappCID>"
  sendReq = sendReq & "<CompanyName>"&strCompanyName&"</CompanyName>"
  sendReq = sendReq & "<BillingAddress>"&strBillingAddress&"</BillingAddress>"
  sendReq = sendReq & "<City>"&strCity&"</City>"
  sendReq = sendReq & "<StateOrProvince>"&strStateOrProvince&"</StateOrProvince>"
  sendReq = sendReq & "<PostalCode>"&strPostalCode&"</PostalCode>"
  sendReq = sendReq & "<CountryID>"&nCountryID&"</CountryID>"
  sendReq = sendReq & "<ContactTitle>"&strContactTitle&"</ContactTitle>"
  sendReq = sendReq & "<PhoneNumber>"&strPhoneNumber&"</PhoneNumber>"
  sendReq = sendReq & "<FaxNumber>"&strFaxNumber&"</FaxNumber>"
  sendReq = sendReq & "<EmailAddress>"&strEmailAddress&"</EmailAddress>"
  sendReq = sendReq & "<WebAddress>"&strWebAddress&"</WebAddress>"
  sendReq = sendReq & "<Notes>"&strNotes&"</Notes>"
  sendReq = sendReq & "<WebSiteUserName>"&strWebSiteUserName&"</WebSiteUserName>"
  sendReq = sendReq & "<WebSitePassword>"&strWebSitePassword&"</WebSitePassword>"
  sendReq = sendReq & "<LeadGenerator>"&bLeadGenerator&"</LeadGenerator>"
  'sendReq = sendReq & "<LeadReferralID>"&nLeadReferralID&"</LeadReferralID>"
  sendReq = sendReq & "<VATCode>"&nVATCode&"</VATCode>"
  sendReq = sendReq & "<VATNo>"&strVATNo&"</VATNo>"
  sendReq = sendReq & "<Reseller>"&bReseller&"</Reseller>"
  sendReq = sendReq & "<ShowCustomer>"&bShowCustomer&"</ShowCustomer>"
  sendReq = sendReq & "<CreditCardCustomer>"&bCreditCardCustomer&"</CreditCardCustomer>"
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