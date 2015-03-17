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
		bDebug = True
		strDB = strOEDB
			
		nOEID = 1165
				
		Set objDisplayOEData = Nothing
		Set objDisplayOEData = New CDBInterface
		objDisplayOEData.Init arrCustDetails, nTotCustDetails	
		
		objDisplayOEData.NewSQL    
		objDisplayOEData.AddFieldToQuery WEBHOSTING_RENEWALDATE
		objDisplayOEData.AddWhere WEBHOSTING_CUSTOMERID, "=", nOEID, "number", "and"	
		objDisplayOEData.QueryDatabase
	
        Response.Write "If you see this then I can see the OE-DB baby!!"

        Response.Write "<br>My renewal date: " & objDisplayOEData.GetValue(WEBHOSTING_RENEWALDATE)

%>
