<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
strDB = strOEDB
updateLogin
Sub updateLogin()
	
	Dim oSPConn
	Dim oSPCmd
	
	Set oSPConn = Server.CreateObject("ADODB.Connection")
	Set oSPCmd = Server.CreateObject("ADODB.Command")
	 
	oSPConn.ConnectionString = strDB
	oSPConn.CursorLocation = adUseServer
	oSPConn.open
	
	oSPCmd.ActiveConnection = oSPConn
	oSPCmd.CommandType = adCmdStoredProc
	oSPCmd.CommandText="ECOM_TodaysOrders"
	
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue", AdVarChar, adParamOutput,8000)
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue2", AdVarChar, adParamOutput,8000)
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue3", AdVarChar, adParamOutput,8000)
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue4", AdVarChar, adParamOutput,8000)
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue5", AdVarChar, adParamOutput,8000)
	oSPCmd.Parameters.Append oSPCmd.CreateParameter("ReturnValue6", AdVarChar, adParamOutput,8000)
	
	oSPCmd.Execute
	response.write oSPCmd("ReturnValue")
	response.write oSPCmd("ReturnValue2")
	response.write oSPCmd("ReturnValue3")
	response.write oSPCmd("ReturnValue4")
	response.write oSPCmd("ReturnValue5")
	response.write oSPCmd("ReturnValue6")
	Set oSPCmd = Nothing
	Set oSPConn = Nothing
	
	
	
End Sub
%>