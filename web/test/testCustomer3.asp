<!--#include virtual="/_private/CSyncSoap.asp"-->
<%
soapURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp"
	
Dim objSoapSync
Set objSoapSync = New CSoap
objSoapSync.Init
	
'soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><FunctionCall>SendCustomerDetailsXML</FunctionCall><CustomerID>1187</CustomerID><OrderEntryCID>7734</OrderEntryCID><CustomerName>Test Name</CustomerName><Company>Test Company</Company><Address1>Test Address 1</Address1><City>Test City</City><County>Test County</County><PostCode>ex27hu</PostCode><TelNo>0800800800</TelNo><NewZappCIDs><NewZappCID>7734</NewZappCID><NewZappCID>2467</NewZappCID></NewZappCIDs><CreditLimit>1000</CreditLimit></Request></SOAP:Body></SOAP:Envelope>"

'soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><FunctionCall>SendCustomerDetailsXML</FunctionCall><CustomerID>1187</CustomerID><OrderEntryCID>1187</OrderEntryCID><CustomerName></CustomerName><Company>Test Company DC</Company><Address1>Test Address 1, Address2</Address1><City>Test City234</City><County>Test Countyb</County><PostCode>ex27hu</PostCode><TelNo>0800800800</TelNo><NewZappCIDs><NewZappCID>7734</NewZappCID><NewZappCID>2467</NewZappCID></NewZappCIDs><CreditLimit>1000</CreditLimit></Request></SOAP:Body></SOAP:Envelope>"

soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><FunctionCall>SendCustomerDetailsXML</FunctionCall><CustomerID>1187</CustomerID><CustomerName>Test Company DC</CustomerName><Address1>Test Address 16, Address2</Address1><City>Test City</City><County>Test County</County><PostCode>tq125bg</PostCode><Country>225</Country><EmailAddress>salesdept@destinet.co.uk</EmailAddress><TelNo>0800800800</TelNo><CreditLimit>1000</CreditLimit><NewZappCIDs><NewZappCID>7734</NewZappCID><NewZappCID>2467</NewZappCID></NewZappCIDs></Request></SOAP:Body></SOAP:Envelope>"


response.write "<br><br>Sent <br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapXML&"</textarea><br><br>"
	
	soapResponseCode = objSoapSync.SendSoap(soapXML, soapURL)	
	
	response.write "<br><br>Response Code<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponseCode&"</textarea><br><br>"
	
	soapResponse = objSoapSync.getResponseText
	

	
	response.write "<br><br>Returned Text<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponse&"</textarea><br><br>"

%>
