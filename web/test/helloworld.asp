<!--#include virtual="/_private/CSyncSoap.asp"-->
<%
'soapURL = "http://172.16.100.211/SOAP/HelloWorld.asp"
soapURL = "http://localhost/SOAP/HelloWorld.asp"
	
Dim objSoapSync
Set objSoapSync = New CSoap
objSoapSync.Init
	
soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request><ID></ID></Request></SOAP:Body></SOAP:Envelope>"


response.write "<br><br>Sent <br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapXML&"</textarea><br><br>"
	
	soapResponseCode = objSoapSync.SendSoap(soapXML, soapURL)	
	
	response.write "<br><br>Response Code<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponseCode&"</textarea><br><br>"
	
	soapResponse = objSoapSync.getResponseText
	

	
	response.write "<br><br>Returned Text<br><textarea name=""ta"" style=""width:500px; height:200px;"">"&soapResponse&"</textarea><br><br>"

%>
