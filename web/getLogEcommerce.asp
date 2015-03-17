<%@ Language=VBScript %>
<!--#include file="_private\OrderEntry_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\CustomerDetailsDBOrderEntry.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\emailer_db.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\obj_init.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\PublishPage.asp"-->
<!--#include file="..\..\PowerLounge\Web\_private\ProjectMethods.asp"-->

<!--#include file="_private\CSyncLibrary.asp"-->
<!--#include file="_private\CSyncSoap.asp"-->
<% 'bDebug=true

	Session.LCID=2057
	Response.ContentType = "text/xml"
	response.buffer = true

	strDB = strOEDB	
	
	
	If request.querystring("bQuery") = 1 Then
		nInvoiceID = request.querystring("InvoiceID")
		dSpecificDate = request.querystring("Date")
		nType = request.querystring("Type")
	Else
		Dim objSoap
		Set objSoap = New CSoap
		objSoap.Init
		
		nInvoiceID = objSoap.ReadSingleNode ("InvoiceID")
		dSpecificDate = objSoap.ReadSingleNode ("Date")
		nType			= objSoap.ReadSingleNode ("Type")
	End If
	
	If Len(dSpecificDate) > 0 Then

	Else
		If Len(nInvoiceID)>0 Then
			' if we have an invoiceID then allow all days if no date is supplied
		Else
			dSpecificDate = date()
		End If
	End If
	
	Dim objLog
	Set objLog = New CDBInterface
	objLog.Init arrCustDetails, nTotCustDetails	
	
	objLog.NewSQL
	objLog.AddFieldToQuery  LOGECOMMERCE_ID 
	objLog.AddFieldToQuery  LOGECOMMERCE_CUSTOMERID
	objLog.AddFieldToQuery  LOGECOMMERCE_ORDERID
	objLog.AddFieldToQuery  LOGECOMMERCE_ORDERDETAILID
	objLog.AddFieldToQuery  LOGECOMMERCE_ACTIONTEXT
	objLog.AddFieldToQuery  LOGECOMMERCE_ACTIONTYPE
	objLog.AddFieldToQuery  LOGECOMMERCE_ACTIONDATE
	objLog.AddFieldToQuery  LOGECOMMERCE_INVOICEID
	
	If Len(nInvoiceID)> 0 Then
		objLog.AddWhere LOGECOMMERCE_INVOICEID, "=", nInvoiceID, "number", "and" 
	End If
	If Len(dSpecificDate)> 0 Then
		strSQLDateTimeEnd = dateAdd("d",1,cDate(dSpecificDate))
		objLog.AddWhere LOGECOMMERCE_ACTIONDATE, ">", ""&cDate(dSpecificDate)&"", "date", "and"
		objLog.AddWhere LOGECOMMERCE_ACTIONDATE, "<", ""&strSQLDateTimeEnd&"", "date", "and"
	End If
	
	If Len(nType)>0 Then
		objLog.AddWhere LOGECOMMERCE_ACTIONTYPE, "=", nType, "number", "and" 
	End If
	
	objLog.QueryDatabase
	
	'response.write objSoap.buildSendSoap
	nTotal = objLog.GetNumberOfRecords()
	
	objLog.GetFirst
	nCountID = 0
	
	strLogXML = ""
	
	Do While nCount < nTotal

		nID 			=  objLog.GetValue(LOGECOMMERCE_ID)
		nCustomerID 	=  objLog.GetValue(LOGECOMMERCE_CUSTOMERID)
		nOrderID 		=  objLog.GetValue(LOGECOMMERCE_ORDERID)
		nOrderDetailsID =  objLog.GetValue(LOGECOMMERCE_ORDERDETAILID)
		strActionText 	=  objLog.GetValue(LOGECOMMERCE_ACTIONTEXT)
		nActionType 	=  objLog.GetValue(LOGECOMMERCE_ACTIONTYPE)
		dActionDate 	=  objLog.GetValue(LOGECOMMERCE_ACTIONDATE)
		nInvoiceID 			=  objLog.GetValue(LOGECOMMERCE_INVOICEID)
		
		strCompany		=  getCustomerName(nCustomerID)
		strProductName  =  getProductName(nOrderDetailsID)
		
		If Len(strProductName) > 0 Then
		
		Else
			strProductName = "-"
		End If
		
		strLogXML = strLogXML & buildRow(nID,nCustomerID,nOrderID,nOrderDetailsID,strActionText,nActionType,dActionDate,nInvoiceID,strCompany,strProductName)
		
		objLog.GetNext
		nCount =nCount + 1
	Loop
	
		strLogXML = buildNode("LogEcommerce",strLogXML)
		strLogXML = buildXMLOutput(strLogXML)
		
		response.write strLogXML

Function buildRow(nID,nCustomerID,nOrderID,nOrderDetailsID,strActionText,nActionType,dActionDate,nInvoiceID,strCompany,strProductName)
		
	buildRow = buildRow & "<LogRecord>"
	buildRow = buildRow & buildNode("ID",nID)
	buildRow = buildRow & buildNode("CustomerID",nCustomerID)
	buildRow = buildRow & buildNode("OrderID",nOrderID)
	buildRow = buildRow & buildNode("OrderDetailsID",nOrderDetailsID)
	buildRow = buildRow & buildNode("ActionText",strActionText)
	buildRow = buildRow & buildNode("ActionType",nActionType)
	buildRow = buildRow & buildNode("ActionDate",dActionDate)
	buildRow = buildRow & buildNode("InvoiceID",nInvoiceID)
	buildRow = buildRow & buildNode("Company","<![CDATA["&strCompany&"]]>")
	buildRow = buildRow & buildNode("ProductName",strProductName)
	buildRow = buildRow & buildNode("Location","OE")
	buildRow = buildRow & "</LogRecord>"
	
End Function

Function buildXMLOutput(strNodes)
	Dim strXML
	strXML = "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>"
	strXML = strXML & strNodes
	buildXMLOutput = strXML
End Function

Function buildNode(strNodeName,vNodeValue)
	buildNode = "<"&strNodeName&">"&vNodeValue&"</"&strNodeName&">"
End Function

Function getCustomerName(nCID)

	getCustomerName = "UNKNOWN"
	
	Dim objCustomerDetails
	Set objCustomerDetails = new CDBInterface
	objCustomerDetails.Init arrCustDetails, nTotCustDetails
	
	objCustomerDetails.NewSQL
	objCustomerDetails.AddFieldToQuery CUSTOMERS_COMPANYNAME
	objCustomerDetails.AddWhere CUSTOMERS_ID, "=", nCID, "number", "and" 
	
	objCustomerDetails.QueryDatabase
	
	If Len(objCustomerDetails.GetValue(CUSTOMERS_COMPANYNAME))>0 Then
		getCustomerName = objCustomerDetails.GetValue(CUSTOMERS_COMPANYNAME)
	End If
End Function

Function getProductName(nOrderDetailID)
	
	getProductName = ""
	If Len(nOrderDetailID)> 0 Then
		Dim oOrderDetails
		Set oOrderDetails = New CDBInterface
		oOrderDetails.Init arrCustDetails, nTotCustDetails
		
		oOrderDetails.NewSQL
		oOrderDetails.AddFieldToQuery ORDERDETAILS_PRODUCTID
		oOrderDetails.AddWhere ORDERDETAILS_ID, "=", nOrderDetailID, "number", "and" 
		oOrderDetails.QueryDatabase
		
		
		If Len(oOrderDetails.getValue(ORDERDETAILS_PRODUCTID))>0 Then
			Dim oProductDetails
			Set oProductDetails = New CDBInterface
			oProductDetails.Init arrCustDetails, nTotCustDetails
			
			oProductDetails.NewSQL
			oProductDetails.AddFieldToQuery PRODUCTS_PRODUCTNAME
			oProductDetails.AddWhere PRODUCTS_ID, "=", oOrderDetails.getValue(ORDERDETAILS_PRODUCTID), "number", "and" 
			
			oProductDetails.QueryDatabase
			If Len(oProductDetails.getValue(PRODUCTS_PRODUCTNAME)) > 0 Then
				getProductName = oProductDetails.getValue(PRODUCTS_PRODUCTNAME)
			End If
		End If
	End If
End Function
%>