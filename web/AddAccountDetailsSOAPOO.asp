<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibraryUNRESTRICTED.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<!--#include file="_private\CSyncOffline.asp"-->
<%
	Response.Buffer = True
	strDB = strOEDB

	Dim objSoapSync
	Set objSoapSync = New CSoap
	objSoapSync.Init
	
	nNewZappCID 		= objSoapSync.ReadSingleNode ("NewZappCID")
	nOrderEntryCID 		= objSoapSync.ReadSingleNode ("OrderEntryCID")
	nWebHostingID 		= objSoapSync.ReadSingleNode ("WebHostingID")
	'Added to strip time from RenewalDate field at TM's request	
	dSmallRenewalDate   = objSoapSync.ReadSingleNode ("RenewalDate")
	
	If isDate(dSmallRenewalDate) Then
    	dRenewalDate        = FormatDateTime(dSmallRenewalDate,vbshortdate)
	End If
	
	strUserName 		= objSoapSync.ReadSingleNode ("UserName")
	strPassword 		= objSoapSync.ReadSingleNode ("Password")
	bActiveStatus 		= objSoapSync.ReadSingleNode ("Active")
	bDemoStatus 		= objSoapSync.ReadSingleNode ("Demo")
	bUsedStatus 		= objSoapSync.ReadSingleNode ("Used")
	nCustomerProductID 	= objSoapSync.ReadSingleNode ("CustomerProductID")
	nProductID 			= objSoapSync.ReadSingleNode ("ProductID")
	nYears 				= objSoapSync.ReadSingleNode ("NumberOfYears")
	
	nOnlineID 			= objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID 			= objSoapSync.ReadSingleNode ("OfflineID")
	nTableID 			= objSoapSync.ReadSingleNode ("TableID")
	bLocked 			= objSoapSync.ReadSingleNode ("Locked")
	bExecuted 			= objSoapSync.ReadSingleNode ("Executed")
	

	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline

	objSyncOffline.addAccountData nNewZappCID,nOrderEntryCID,nWebHostingID,dRenewalDate,strUserName,strPassword,bActiveStatus,bDemoStatus,bUsedStatus,nCustomerProductID,nProductID,nYears,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
	
	Set objSyncOffline = Nothing
%>

