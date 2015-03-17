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
		
		response.write strDB
			
		nWebHostingID = 1483
		strUserName = "admin10441"
		strPassword = "oj4514k"
				
		Set objSavePOPAccounts = Nothing
		Set objSavePOPAccounts = New CDBInterface
		objSavePOPAccounts.Init arrCustDetails, nTotCustDetails	
		
		objSavePOPAccounts.NewSQL    
		objSavePOPAccounts.StoreValue LOGECOMMERCE_CUSTOMERID, nWebHostingID

		objSavePOPAccounts.InsertUsingStoredValues "LogEcommerce"

%>
