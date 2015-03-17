<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibrary.asp"-->
<!--#include file="_private\CSyncOffline.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
	Response.Buffer = True 
	'bDebug = True
	
	'DB Connection
	strDB = strOEDB	
	
	'1st read the nodes
	Dim objSoapSync
	Set objSoapSync = New CSoap
	objSoapSync.Init
	
	nID						= objSoapSync.ReadSingleNode ("ID")
	nOrderID				= objSoapSync.ReadSingleNode ("OrderID")
	nOrderEntryOrderID		= objSoapSync.ReadSingleNode ("OrderEntryOrderID")
	nProductID				= objSoapSync.ReadSingleNode ("ProductID")
	nOrderEntryProductID	= objSoapSync.ReadSingleNode ("OrderEntryProductID")
	nQuantity				= objSoapSync.ReadSingleNode ("Quantity")
	nUnitPrice				= objSoapSync.ReadSingleNode ("UnitPrice")
	nDiscount				= objSoapSync.ReadSingleNode ("Discount")
	nAccountType			= objSoapSync.ReadSingleNode ("AccountType")
	nCustomerID				= objSoapSync.ReadSingleNode ("CustomerID")
	nNewZappCID				= objSoapSync.ReadSingleNode ("NewZappCID")
	dRenewalDate			= objSoapSync.ReadSingleNode ("RenewalDate")
	bStatistics				= objSoapSync.ReadSingleNode ("Statistics")
	bFrontPage				= objSoapSync.ReadSingleNode ("FrontPage")
	bServerSideIncludes		= objSoapSync.ReadSingleNode ("ServerSideIncludes")
	bSSL					= objSoapSync.ReadSingleNode ("SSL")
	bIndexServer			= objSoapSync.ReadSingleNode ("IndexServer")
	bCurrent				= objSoapSync.ReadSingleNode ("Current")
	nOnlineID				= objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID				= objSoapSync.ReadSingleNode ("OfflineID")
	nTableID				= objSoapSync.ReadSingleNode ("TableID")
	bLocked					= objSoapSync.ReadSingleNode ("Locked")
	bExecuted				= objSoapSync.ReadSingleNode ("Executed")
	
	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline	

	objSyncOffline.addOrderDetails nID,nOrderID,nOrderEntryOrderID,nProductID,nOrderEntryProductID,nQuantity,nUnitPrice,nDiscount,nAccountType,nCustomerID,nNewZappCID,dRenewalDate,bStatistics,bFrontPage,bServerSideIncludes,bSSL,bIndexServer,bCurrent,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
	
	Set objSyncOffline = Nothing
%>