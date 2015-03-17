<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db_test.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<%
	bDebug = True
	strDB = strOEDB
				
	Set objNewDB = Nothing
	Set objNewDB = New CDBInterface
	objNewDB.Init arrCustDetails, nTotCustDetails
		
	objNewDB.NewSQL
	objNewDB.AddFieldToQuery WEBHOSTING_ID
	objNewDB.AddFieldToQuery WEBHOSTING_CID
	objNewDB.AddWhere WEBHOSTING_ID, "=", 1734, "number", "and"
	objNewDB.QueryDatabase

       	nNumberOfRecords = objNewDB.GetNumberOfRecords()
	nCID = objNewDB.GetValue(WEBHOSTING_CID)

	Response.Write "<br><br>BOOM!!!! - we are working... this is new OE-DB data using a new connection class!<br><br> CID = " & nCID & "<br>Number of records = " & nNumberOfRecords & "<br><br><br><br><br><br><br><br><br>KA-BOOM....." 

%>
