<!--#include file="..\_private\CSyncSoap.asp"-->
<%
	Dim objSoap
	Set objSoap = New CSoap
	objSoap.Init
	
	
	objSoap.SetNodeList ("NewZappCID")
	nNodeListLength =  objSoap.GetNodeListLength()
	response.write "nNodeListLength = " & nNodeListLength & "<br>"
	For i = 0 to nNodeListLength-1
		response.write objSoap.getItemText(i) & "<br>"
 	Next

%>

