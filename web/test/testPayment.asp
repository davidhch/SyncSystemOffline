<!--#include virtual="/_private/CSyncSoap.asp"-->
<%
nPaymentsID = 6094
nOrderID = 7197
nPaymentAmount = 15
dPaymentDate = Date()
nCreditCardID = 2
nPaymentMethodID = 14
strTransactionCode = 6094
strProcessingValid = "NA"
strProcessingCode = "NA"
strProcessingCodeMessage = "NA"
strAuthCode = "NA"
strAuthMessage = "NA"
strResponseCode = "NA"
strResponseCodeMessage = "NA"
nAmountToCharge = 15


	sendReq = ""
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
			
	soapXML = sendReq

	'soapURL = "http://mailex.destinet.co.uk/test/return.asp"
	soapURL = "http://www02.newzapp.co.uk/SOAP/AddPaymentSOAP.asp"
	'soapURL = "http://system.newzapp.co.uk/SOAP/AddPaymentSOAP.asp"
	
	'soapURL = "http://system.newzapp.co.uk/test/testnothing.asp"
	'soapURL = "http://www02.newzapp.co.uk/test/testnothing.asp"
	'soapURL = "http://www02.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp"
	
	Dim objSoapSync
	Set objSoapSync = New CSoap
	objSoapSync.Init
	
	response.write "<br><br>Sent <br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapXML&"</textarea><br><br>"
	
	soapResponseCode = objSoapSync.SendSoap(soapXML, soapURL)	
	
	response.write "<br><br>Response Code<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponseCode&"</textarea><br><br>"
	
	soapResponse = objSoapSync.getResponseText
	

	
	response.write "<br><br>Returned Text<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponse&"</textarea><br><br>"
%>