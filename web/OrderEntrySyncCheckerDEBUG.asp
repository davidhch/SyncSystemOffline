<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibraryUNRESTRICTED.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
	bDebug = True
	Response.Buffer = True 
	strDB = strOEDB
	
	Dim objSyncLibrary
	Set objSyncLibrary = New CSyncLibrary

	objSyncLibrary.syncCheck	
	
	Set objSyncLibrary = Nothing
%>