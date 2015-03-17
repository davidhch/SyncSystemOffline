<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibraryUNRESTRICTEDAdditions.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
	Response.Buffer = True 
	strDB = strOEDB

	Dim objSyncLibrary
	Set objSyncLibrary = New CSyncLibrary

	objSyncLibrary.processSoap	
	
	Set objSyncLibrary = Nothing
%>
