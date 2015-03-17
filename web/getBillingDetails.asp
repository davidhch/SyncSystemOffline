<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibrary.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<%
	Response.Buffer = True 

	strDB = strOEDB	
	
	Dim objSoap
	Set objSoap = New CSoap
	objSoap.Init
	
	nNewZappCID = objSoap.ReadSingleNode ("NewZappCID")
	
	Dim oCustomerWebHost
	Set oCustomerWebHost = New CDBInterface
	oCustomerWebHost.Init arrCustDetails, nTotCustDetails	
	
	Dim oCustomerContacts
	Set oCustomerContacts = New CDBInterface
	oCustomerContacts.Init arrCustDetails, nTotCustDetails
	
	Dim oCustomerCustomers
	Set oCustomerCustomers = New CDBInterface
	oCustomerCustomers.Init arrCustDetails, nTotCustDetails
		
	oCustomerWebHost.NewSQL
	oCustomerWebHost.AddFieldToQuery WEBHOSTING_CID
	oCustomerWebHost.AddFieldToQuery WEBHOSTING_CUSTOMERID
	oCustomerWebHost.AddWhere WEBHOSTING_CID, "=", nNewZappCID, "number", "and"
	oCustomerWebHost.QueryDatabase	
	
	nNCID = oCustomerWebHost.GetValue(WEBHOSTING_CID)
	nOECID = oCustomerWebHost.GetValue(WEBHOSTING_CUSTOMERID)
	
	oCustomerContacts.NewSQL
	oCustomerContacts.AddFieldToQuery CONTACTS_FIRSTNAME
	oCustomerContacts.AddFieldToQuery CONTACTS_LASTNAME
	oCustomerContacts.AddWhere CONTACTS_CUSTOMERID, "=", nOECID, "number", "and"
	oCustomerContacts.QueryDatabase	

	strFirstname = oCustomerContacts.GetValue(CONTACTS_FIRSTNAME)
	strLastname = oCustomerContacts.GetValue(CONTACTS_LASTNAME)
	
	oCustomerCustomers.NewSQL
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_COMPANYNAME
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_BILLINGADDRESS
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_CITY
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_STATEORPROVINCE
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_POSTALCODE
	oCustomerCustomers.AddFieldToQuery CUSTOMERS_COUNTRYID
	oCustomerCustomers.AddWhere CUSTOMERS_ID, "=", nOECID, "number", "and" 
	oCustomerCustomers.QueryDatabase
	
	objSoap.clearSendNodes()
	objSoap.AddSendNode "FunctionCall", "SendCustomerDetailsXML"
	objSoap.AddSendNode "NewZappCID", nNCID
	objSoap.AddSendNode "CustomerName", strFirstname & " " & strLastname
	objSoap.AddSendNode "CompanyName", oCustomerCustomers.GetValue(CUSTOMERS_COMPANYNAME)
	objSoap.AddSendNode "BillingAddress", oCustomerCustomers.GetValue(CUSTOMERS_BILLINGADDRESS)
	objSoap.AddSendNode "City", oCustomerCustomers.GetValue(CUSTOMERS_CITY)
	objSoap.AddSendNode "County", oCustomerCustomers.GetValue(CUSTOMERS_STATEORPROVINCE)
	objSoap.AddSendNode "PostCode", oCustomerCustomers.GetValue(CUSTOMERS_POSTALCODE)
	objSoap.AddSendNode "Country",oCustomerCustomers.GetValue(CUSTOMERS_COUNTRYID)
	
	response.write objSoap.buildSendSoap
%>