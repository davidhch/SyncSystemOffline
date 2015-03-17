<%@ Language=VBScript %>
<%

nPaymentsID = 5868
nOrderID = 6913
nPaymentAmount = 16
dPaymentDate = Date()
nCreditCardID = 2
nPaymentMethodID = 14
strTransactionCode = "NA"
strProcessingValid = "NA"
strProcessingCode = "NA"
strProcessingCodeMessage = "NA"
strAuthCode = "NA"
strAuthMessage = "NA"
strResponseCode = "NA"
strResponseCodeMessage = "NA"
nAmountToCharge = 10


Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://orderentry-soap.destinet.co.uk/AddPaymentSOAP.asp" 

Function sendReq()
  sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
  sendReq = sendReq & "<SOAP:Body>"
  sendReq = sendReq & "<Request>"
  sendReq = sendReq & "<PaymentsID>"&nPaymentsID&"</PaymentsID>"
  sendReq = sendReq & "<OrderID>"&nOrderID&"</OrderID>"
  sendReq = sendReq & "<PaymentAmount>"&nPaymentAmount&"</PaymentAmount>"
  sendReq = sendReq & "<PaymentDate>"&dPaymentDate&"</PaymentDate>"
  sendReq = sendReq & "<CreditCardID>"&nCreditCardID&"</CreditCardID>"
  sendReq = sendReq & "<PaymentMethodID>"&nPaymentMethodID&"</PaymentMethodID>"
  sendReq = sendReq & "<TransactionCode>"&strTransactionCode&"</TransactionCode>"
  sendReq = sendReq & "<ProcessingValid>"&strProcessingValid&"</ProcessingValid>"
  sendReq = sendReq & "<ProcessingCode>"&strProcessingCode&"</ProcessingCode>"
  sendReq = sendReq & "<ProcessingCodeMessage>"&strProcessingCodeMessage&"</ProcessingCodeMessage>"
  sendReq = sendReq & "<AuthCode>"&strAuthCode&"</AuthCode>"
  sendReq = sendReq & "<AuthMessage>"&strAuthMessage&"</AuthMessage>"
  sendReq = sendReq & "<ResponseCode>"&strResponseCode&"</ResponseCode>"
  sendReq = sendReq & "<ResponseCodeMessage>"&strResponseCodeMessage&"</ResponseCodeMessage>"
  sendReq = sendReq & "<AmountToCharge>"&nAmountToCharge&"</AmountToCharge>"
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