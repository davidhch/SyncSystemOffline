<!--#include file="..\..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\_private\OrderEntry_db.asp"-->
<%
nOrderID = 1109
strActionText = "Order Detail Added from NZ"
nActionType = 2
nOrderDetailsID = 2025

strDB = strOEDB

LogOrder nOrderID,nOrderDetailsID,strActionText,nActionType

Sub LogOrder (nOrderID,nOrderDetailsID,strActionText,nActionType)
	Dim objSPConnEcom
	Dim objSPCmdEcom
	
	Set objSPConnEcom = Server.CreateObject("ADODB.Connection")
	Set objSPCmdEcom = Server.CreateObject("ADODB.Command")
	 
	objSPConnEcom.ConnectionString = strDB 
	objSPConnEcom.CursorLocation = adUseServer
	objSPConnEcom.open
	
	objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nOrderID", adInteger, adParamInput, , nOrderID)
	objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nOrderDetailID", adInteger, adParamInput, , nOrderDetailsID)
	objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("strAction", adVarChar, adParamInput, Len(strActionText), strActionText)
	objSPCmdEcom.Parameters.Append objSPCmdEcom.CreateParameter("nActionType", adInteger, adParamInput, , nActionType)
	
	objSPCmdEcom.ActiveConnection = objSPConnEcom
	objSPCmdEcom.CommandType = adCmdStoredProc
	objSPCmdEcom.CommandText="ECOM_AddLogEcommerce"
	objSPCmdEcom.Execute
	
	Set objSPCmdEcom = Nothing
	Set objSPConnEcom = Nothing
End Sub
%>