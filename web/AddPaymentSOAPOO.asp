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
	
	nPaymentsID 				= objSoapSync.ReadSingleNode ("PaymentsID")
	nOrderID 					= objSoapSync.ReadSingleNode ("OrderID")
	nOrderEntryOrderID 			= objSoapSync.ReadSingleNode ("OrderEntryOrderID")
	nPaymentAmount 				= objSoapSync.ReadSingleNode ("PaymentAmount")
	dPaymentDate 				= objSoapSync.ReadSingleNode ("PaymentDate")
	nCreditCardID 				= objSoapSync.ReadSingleNode ("CreditCardID")
	nPaymentMethodID 			= objSoapSync.ReadSingleNode ("PaymentMethodID")
	strTransactionCode 			= objSoapSync.ReadSingleNode ("TransactionCode")
	strProcessingValid 			= objSoapSync.ReadSingleNode ("ProcessingValid")
	strProcessingCode 			= objSoapSync.ReadSingleNode ("ProcessingCode")
	strProcessingCodeMessage 	= objSoapSync.ReadSingleNode ("ProcessingCodeMessage")
	strAuthCode 				= objSoapSync.ReadSingleNode ("AuthCode")
	strAuthMessage 				= objSoapSync.ReadSingleNode ("AuthMessage")
	strResponseCode 			= objSoapSync.ReadSingleNode ("ResponseCode")
	strResponseCodeMessage 		= objSoapSync.ReadSingleNode ("ResponseCodeMessage")
	nAmountToCharge 			= objSoapSync.ReadSingleNode ("AmountToCharge")
	nOnlineID 					= objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID 					= objSoapSync.ReadSingleNode ("OfflineID")
	nTableID 					= objSoapSync.ReadSingleNode ("TableID")
	bLocked 					= objSoapSync.ReadSingleNode ("Locked")
	bExecuted 					= objSoapSync.ReadSingleNode ("Executed")	
	
	
	'This would run on an ASP page
	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline	 

	'...now do what you want to do!!!!
	objSyncOffline.addPayment nPaymentsID,nOrderID,nOrderEntryOrderID,nPaymentAmount,dPaymentDate,nCreditCardID,nPaymentMethodID,strTransactionCode,strProcessingValid,strProcessingCode,strProcessingCodeMessage,strAuthCode,strAuthMessage,strResponseCode,strResponseCodeMessage,nAmountToCharge,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
	
	Set objSyncOffline = Nothing

%>