<!--#include virtual="\_private\CSyncSoap.asp"-->
<%
' Some soap XML
strXML = ""
strXML = strXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
strXML = strXML & "<SOAP:Body>"
strXML = strXML & "<Request>"
strXML = strXML & "<YourNodes>Your Data</YourNodes>"
strXML = strXML & "</Request>"
strXML = strXML & "</SOAP:Body>"
strXML = strXML & "</SOAP:Envelope>"
	
'Simple Return of just text
'exampleSoapSend strXML,"https://system.newzapp.co.uk/soap/sendAccountData.asp"

'Thought i'd include this one as it returns XML fromt he PTMA
exampleSoapSend strXML,"http://dev210.newzapp.co.uk/test/responseSOAP.asp"

Sub exampleSoapSend(strXML,strURL)
	
	Dim objSoap
	Set objSoap = New CSoap
	objSoap.Init
	
	nResponseCode 	= objSoap.SendSoap(strXML,strURL)
	
	Response.Write "RETURN CODE " & nResponseCode
	
	strResponseText = objSoap.getResponseText
	Response.Write "<br><br>RETURN TEXT " & strResponseText
	
	Set objSoap = Nothing
	
	' TO READ THE RESPONSE AS XML IT NEEDS TO BE LOADED
	bGETXMLSupport = false ' OE does not support the function call getXML - perhaps it should
	If bGETXMLSupport = true Then
		Dim objSoapReturn
		Set objSoapReturn = New CSoap
		objSoapReturn.Init
		objSoapReturn.loadNewXML(strResponseText)
	
		' XML IS LOADED SO CAN BE READ BY OUR OBJECT
		Response.Write "<br><br>RETURN XML " & objSoapReturn.getXML
		
		Set objSoapReturn = Nothing
	End If
End Sub
%>