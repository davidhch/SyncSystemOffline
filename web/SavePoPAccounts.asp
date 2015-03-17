<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntryAdditions.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibraryUNRESTRICTEDAdditions.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<!--#include file="_private\CSyncOfflineAdditions.asp"-->
<%
		bDebug = True
		
		strDB = strOEDB
		
		response.write "//--------------------------------------------------<br>"
		response.write "strDB: "& strDB & "<br>"
		response.write "//--------------------------------------------------<br><br>"
				
		nPOPAccountID = 2226
		strUserName = "TestUN"
		strPassword = "TestPW"
		
			Set objFields = Nothing
			Set objFields = New CDBInterface
			objFields.Init arrCustDetails, nTotCustDetails	
			
			
			objFields.NewSQL    
			objFields.StoreValue POPACCOUNTS_WEBHOSTINGID, 1483
			objFields.StoreValue POPACCOUNTS_USERNAME, "TestUserName" 
			objFields.StoreValue POPACCOUNTS_PASSWORD, "TestPassword"
			response.write objFields.getQuerySQL()
			objFields.InsertUsingStoredValues "POP Accounts"
			Response.Write "<br>INSERT --------------------------------------------------"
					


%>
