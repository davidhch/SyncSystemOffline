<%@ Language=VBScript %>
<!--#include file="_private\CSoap.asp"-->
<%
	Response.Buffer = True 

	Dim objSoap
	Set objSoap = New CSoap
	objSoap.Init	
	
	strProductName 	= objSoapSync.ReadSingleNode ("ProductName")
	nProductCost 	= objSoapSync.ReadSingleNode ("ProductCost")
	
	response.write strProductName
	
%>