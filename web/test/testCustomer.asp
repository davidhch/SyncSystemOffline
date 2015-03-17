<!--#include virtual="/_private/CSyncSoap.asp"-->
<%




	'soapURL = "http://mailex.destinet.co.uk/test/return.asp"
	'soapURL = "http://www02.newzapp.co.uk/SOAP/AddPaymentSOAP.asp"
	'soapURL = "http://system.newzapp.co.uk/SOAP/AddPaymentSOAP.asp"
	
	'soapURL = "http://system.newzapp.co.uk/test/testnothing.asp"
	'soapURL = "http://www02.newzapp.co.uk/test/testnothing.asp"
	'soapURL = "http://www02.newzapp.co.uk/SOAP/AddSyncLookupSOAP.asp"
soapURL = "http://destinetapps.destinet.co.uk/OrderEntrySoapHandlerOO.asp"
	
Dim objSoapSync
Set objSoapSync = New CSoap
objSoapSync.Init
	
objSoapSync.clearSendNodes()
objSoapSync.AddSendNode "FunctionCall","SendCustomerDetailsXML"
objSoapSync.AddSendNode "ResellerID",""
objSoapSync.AddSendNode "CustomerID",1117
objSoapSync.AddSendNode "OrderEntryCID",2067
objSoapSync.AddSendNode "CustomerName","Justin Mortimore"
objSoapSync.AddSendNode "Company","Jays"
objSoapSync.AddSendNode "Address1","I live here"
objSoapSync.AddSendNode "Address2",""
objSoapSync.AddSendNode "City","Torquay"
objSoapSync.AddSendNode "County","NA"
objSoapSync.AddSendNode "PostCode","NA"
objSoapSync.AddSendNode "EmailAddress","jmortimore@destinet.co.uk"
objSoapSync.AddSendNode "TelNo","NA"
objSoapSync.AddSendNode "NewsletterTitle",""
objSoapSync.AddSendNode "NewsletterEmail",""
'objSoapSync.AddSendNode "StartDate",""
'objSoapSync.AddSendNode "LastDate",""
objSoapSync.AddSendNode "TotalNewsSent",0
'objSoapSync.AddSendNode "LastUsed",""
objSoapSync.AddSendNode "TimesLoggedIn",1589
objSoapSync.AddSendNode "Active",True
objSoapSync.AddSendNode "Demo",False
objSoapSync.AddSendNode "HashKey",""
objSoapSync.AddSendNode "MTA",""
objSoapSync.AddSendNode "NewZappCID",2067
objSoapSync.AddSendNode "CreditLimit",3000

soapXML = objSoapSync.buildSendSoap

soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><FunctionCall>SendCustomerDetailsXML</FunctionCall><CustomerID>1117</CustomerID><OrderEntryCID>1117</OrderEntryCID><CustomerName>Justin MortimorZXZXZX</CustomerName><Company>Justin's account</Company><Address1>I live here</Address1><City>Torquay</City><County>NA</County><PostCode>NA</PostCode><TelNo>NA</TelNo><NewZappCID>2067</NewZappCID><CreditLimit>10000</CreditLimit></Request></SOAP:Body></SOAP:Envelope>"
	
	response.write "<br><br>Sent <br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapXML&"</textarea><br><br>"
	
	response.write "<br>" & soapURL & "<br>"
	
	soapResponseCode = objSoapSync.SendSoap(soapXML, soapURL)	
	
	response.write "<br><br>Response Code<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponseCode&"</textarea><br><br>"
	
	soapResponse = objSoapSync.getResponseText
	

	
	response.write "<br><br>Returned Text<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponse&"</textarea><br><br>"
%>