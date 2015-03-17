<%@ Language=VBScript %>
<!--#include file="_private\CSoap.asp"-->
<%
	Response.Buffer = True 

	Dim objSoap
	Set objSoap = New CSoap
	objSoap.Init	
	
	objSoap.SetEnvelopePath("SOAP:Envelope/SOAP:Body/Products/Product/")
	strProductName 	= objSoap.ReadSingleNode ("ProductName")
	nProductCost 	= objSoap.ReadSingleNode ("ProductCost")
	GetNodeList  = objSoap.GetNodeList ("Product")
	
	response.write("Current Orders: " & GetNodeList)
	
%>