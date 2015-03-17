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
	
	nCustomerID = objSoapSync.ReadSingleNode ("CustomerID")
	nOrderEntryCID = objSoapSync.ReadSingleNode ("OrderEntryCID")
	nNewZappCID = objSoapSync.ReadSingleNode ("NewZappCID")
	nQuotationID = objSoapSync.ReadSingleNode ("QuotationID")
	nEmployeeID = objSoapSync.ReadSingleNode ("EmployeeID")
	dOrderDate = objSoapSync.ReadSingleNode ("OrderDate")
	strPurchaseOrderNumber = objSoapSync.ReadSingleNode ("PurchaseOrderNumber")
	strInvoiceDescription = objSoapSync.ReadSingleNode ("InvoiceDescription")
	strShipName = objSoapSync.ReadSingleNode ("ShipName") 
	strShipAddress = objSoapSync.ReadSingleNode ("ShipAddress") 
	strShipCity = objSoapSync.ReadSingleNode ("ShipCity")
	strShipStateOrProvince = objSoapSync.ReadSingleNode ("ShipStateOrProvince")
	strShipPostalCode = objSoapSync.ReadSingleNode ("ShipPostalCode")
	strShipCountry = objSoapSync.ReadSingleNode ("ShipCountry")
	strShipPhoneNumber = objSoapSync.ReadSingleNode ("ShipPhoneNumber")
	dShipDate = objSoapSync.ReadSingleNode ("ShipDate")
	nShippingMethodID = objSoapSync.ReadSingleNode ("ShippingMethodID")
	nFreightCharge = objSoapSync.ReadSingleNode ("FreightCharge")
	nSalesTaxRate = objSoapSync.ReadSingleNode ("SalesTaxRate")
	nLeadReferralID = objSoapSync.ReadSingleNode ("LeadReferralID")
	strProposalReference = objSoapSync.ReadSingleNode ("ProposalReference")
	dEntryDate = objSoapSync.ReadSingleNode ("EntryDate")
	bCommissionPaid = objSoapSync.ReadSingleNode ("CommissionPaid")
	nInvoiceID = objSoapSync.ReadSingleNode ("InvoiceID")
	nInstallCID = objSoapSync.ReadSingleNode ("InstallCID")
	nResellerID = objSoapSync.ReadSingleNode ("ResellerID")
	'strUniqueKey = objSoapSync.ReadSingleNode ("UniqueKey")
	nOnlineID = objSoapSync.ReadSingleNode ("OnlineID")
	nOfflineID = objSoapSync.ReadSingleNode ("OfflineID")
	nTableID = objSoapSync.ReadSingleNode ("TableID")
	bLocked = objSoapSync.ReadSingleNode ("Locked")
	bExecuted = objSoapSync.ReadSingleNode ("Executed")
	
	'response.write "DC nNewZappCID: " & nNewZappCID & "<br>"
	
	Dim objSyncOffline
	Set objSyncOffline = New CSyncOffline	

	objSyncOffline.addOrder nCustomerID,nNewZappCID,nOrderEntryCID,nQuotationID,nEmployeeID,dOrderDate,strPurchaseOrderNumber,strInvoiceDescription,strShipName,strShipAddress,strShipCity,strShipStateOrProvince,strShipPostalCode,strShipCountry,strShipPhoneNumber,dShipDate,nShippingMethodID,nFreightCharge,nSalesTaxRate,nLeadReferralID,strProposalReference,dEntryDate,bCommissionPaid,nInvoiceID,nInstallCID,nResellerID,nOnlineID,nOfflineID,nTableID,bLocked,bExecuted
	
	Set objSyncOffline = Nothing
%>