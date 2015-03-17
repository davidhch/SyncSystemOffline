<!--#include virtual="/_private/CSyncSoap.asp"-->
<%
soapURL = "http://www04.newzapp.co.uk/SOAP/AddCustomerDetailsSOAPOO.asp"
	
Dim objSoapSync
Set objSoapSync = New CSoap
objSoapSync.Init
	
soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><ID></ID><ResellerID></ResellerID><CustomerID>1117</CustomerID><OrderEntryCID>2067</OrderEntryCID><CustomerName>Justin Mortimore</CustomerName><Company>Jays</Company><Address1>I live here</Address1><Address2></Address2><City>Torquay</City><County>NA</County><PostCode>NA</PostCode><EmailAddress>jmortimore@destinet.co.uk</EmailAddress><TelNo>NA</TelNo><NewsletterTitle></NewsletterTitle><NewsletterEmail></NewsletterEmail><TotalNewsSent>0</TotalNewsSent><TimesLoggedIn>1589</TimesLoggedIn><Active>True</Active><Demo>False</Demo><HashKey></HashKey><MTA></MTA><NewZappCID>2067</NewZappCID><CreditLimit>3000</CreditLimit><OnlineID></OnlineID><OfflineID>1117</OfflineID><TableID>6</TableID><Locked>1</Locked><Executed>0</Executed></Request></SOAP:Body></SOAP:Envelope>"


response.write "<br><br>Sent <br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapXML&"</textarea><br><br>"
	
	soapResponseCode = objSoapSync.SendSoap(soapXML, soapURL)	
	
	response.write "<br><br>Response Code<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponseCode&"</textarea><br><br>"
	
	soapResponse = objSoapSync.getResponseText
	
	response.write "<br><br>Returned Text<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponse&"</textarea><br><br>"

%>
