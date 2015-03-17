<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->
<%
	Response.Buffer = True 
	
	Sub ReturnSub(bReportError)
	
		returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		returnXML = returnXML & "<SOAP:Body>"
		returnXML = returnXML & "<MessageReturn>"
		returnXML = returnXML & "<Time>"&Now()&"</Time>"
		returnXML = returnXML & "<Message>"&bAction&"</Message>"
		returnXML = returnXML & "<MessageResponse>"&strMessage&"</MessageResponse>"
		returnXML = returnXML & "<NewZappCustomerID>"&nNewZappCustomerID&"</NewZappCustomerID>"
		returnXML = returnXML & "<OrderEntryCustomerID>"&nOrderEntryCustomerID&"</OrderEntryCustomerID>"
		returnXML = returnXML & "</MessageReturn>"
		returnXML = returnXML & "</SOAP:Body>"
		returnXML = returnXML & "</SOAP:Envelope>"

		Response.Write (returnXML) 
		
	End Sub	
	
	On Error Resume Next
	
	Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
	Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	
	objxmldom.load(Request)
	
	nOnlineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OnlineID").Text
	nOfflineID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/OfflineID").Text
	nTableID = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/TableID").Text
	bLocked = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Locked").Text
	bExecuted = objxmldom.selectSingleNode("SOAP:Envelope/SOAP:Body/Request/Executed").Text
		
	'bDebug = True
	
	strDB = strOEDB
		
	Set objAddSyncLookup = Nothing
	Set objAddSyncLookup = New CDBInterface
	objAddSyncLookup.Init arrCustDetails, nTotCustDetails
		
	objAddSyncLookup.NewSQL
	objAddSyncLookup.AddFieldToQuery SYNCLOOKUP_ID
	objAddSyncLookup.AddWhere SYNCLOOKUP_OFFLINEID, "=", nOfflineID, "number", "and"
	objAddSyncLookup.AddWhere SYNCLOOKUP_TABLEID, "=", nTableID, "number", "and"
	objAddSyncLookup.QueryDatabase
	
	nSyncLookupID = objAddSyncLookup.GetValue(SYNCLOOKUP_ID)
	
	objAddSyncLookup.NewSQL
	objAddSyncLookup.StoreValue SYNCLOOKUP_ONLINEID, nOnlineID
	objAddSyncLookup.StoreValue SYNCLOOKUP_OFFLINEID, nOfflineID
	objAddSyncLookup.StoreValue SYNCLOOKUP_LOCKED, bLocked
	objAddSyncLookup.StoreValue SYNCLOOKUP_EXECUTED, bExecuted
	objAddSyncLookup.SelectRecordForUpdate SYNCLOOKUP_ID, nSyncLookupID
	objAddSyncLookup.UpdateUsingStoredValues "SyncLookUp"
			
	Set objxmldom = Nothing
	Set objxmlhttp = Nothing
		
%>
