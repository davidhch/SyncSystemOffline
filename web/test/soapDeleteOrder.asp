<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	Dim strURL
	Dim strXML
	Dim objxmldom
	Dim objxmlhttp
	
	soapXML = "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1""><SOAP:Body><Request>"
	soapXML = soapXML &"<FunctionCall>SendDeleteOrderXML</FunctionCall>"
	soapXML = soapXML &"<InvoiceID>7334</InvoiceID>"
	soapXML = soapXML &"</Request></SOAP:Body></SOAP:Envelope>"
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP")
	
	objxmlhttp.open "POST", "http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp", False
	objxmlhttp.setRequestHeader "Man", POST & " http://orderentry-soap.destinet.co.uk/OrderEntrySOAPHandler.asp HTTP/1.1"
	objxmlhttp.setRequestHeader "MessageType", "CALL"
	
	objxmlhttp.send(soapXML)
	
	response.write "<br>STATUS: " & objxmlhttp.Status
	response.write "<br>RESPONSE" & objxmlhttp.responseText

%>
