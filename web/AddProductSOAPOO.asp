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
	
	Dim objSoapSync
	Set objSoapSync = New CSoap
	objSoapSync.Init	
	
	nProductID 					= objSoapSync.ReadSingleNode ("ProductID")
	nOnlineProductID 			= objSoapSync.ReadSingleNode ("OnlineProductID")
	strProductName 				= objSoapSync.ReadSingleNode ("ProductName")
	nProductCategoryID 			= objSoapSync.ReadSingleNode ("ProductCategoryID")
	nUnitPrice 					= objSoapSync.ReadSingleNode ("UnitPrice")
	nUnitCost 					= objSoapSync.ReadSingleNode ("UnitCost")
	strOneOffCost 				= objSoapSync.ReadSingleNode ("OneOffCost")
	nSupplierID 				= objSoapSync.ReadSingleNode ("SupplierID")
	strDescription 				= objSoapSync.ReadSingleNode ("Description") 
	strSystemDescription 		= objSoapSync.ReadSingleNode ("SystemDescription") 
	nSetupPrice 				= objSoapSync.ReadSingleNode ("SetupPrice")
	nOrder 						= objSoapSync.ReadSingleNode ("Order")
	bActive 					= objSoapSync.ReadSingleNode ("Active")
	nOrderEntryProductID 		= objSoapSync.ReadSingleNode ("OrderEntryProductID")
	nNewZappCustomerID 			= objSoapSync.ReadSingleNode ("nNewZappCustomerID")
	nOrderEntryCustomerID 		= objSoapSync.ReadSingleNode ("nOrderEntryCustomerID")
	nOnlineID 					= objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID 					= objSoapSync.ReadSingleNode ("OfflineID")
	nTableID 					= objSoapSync.ReadSingleNode ("TableID")
	bLocked 					= objSoapSync.ReadSingleNode ("Locked")
	bExecuted 					= objSoapSync.ReadSingleNode ("Executed")
	
	'This would run on an ASP page
	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline	 

	'...now do what you want to do!!!!
	objSyncOffline.addProduct nProductID,nOnlineProductID,strProductName,nProductCategoryID,nUnitPrice,nUnitCost,strOneOffCost,nSupplierID,strDescription,strSystemDescription,nSetupPrice,nOrder,bActive,nOrderEntryProductID,nNewZappCustomerID,nOrderEntryCustomerID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
	
	Set objSyncOffline = Nothing

%>