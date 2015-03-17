<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Orders</title>
<script language="JavaScript" type="text/javascript" src="/test/mh/js/xmlhttprequest.js"></script>
<script language="JavaScript" type="text/javascript" src="/test/mh/js/xmlSoap.js"></script>
<script type="text/javascript" language="javascript">

var arrOrders =new Array();
var arrxmlOrders =new Array();
var soapData = ""
var nArrayItem=0;

function sendSOAP()
{
	xmlParseSoap(soapData,"http://orderentry-soap.destinet.co.uk/test/mh/saveOrder.asp");
}

function returnSoap()
{
	alert(strResponse);
}

function openSoap()
{
	soapData = '<SOAP:Envelope xmlns:SOAP="urn:schemas-xmlsoap-org:soap.v1">';
	soapData = soapData + '<SOAP:Body>';
	soapData = soapData + '<Products>';
}

function addNode()
{	
	var strNode = "";
	
	for(i=0;i<=arrOrders.length-1;i++)
	{
		strNode = strNode + '<Product>';
		strNode = strNode + '<ProductName>'+arrOrders[i].ProductName+'</ProductName>';
		strNode = strNode + '<ProductCost>'+arrOrders[i].ProductCost+'</ProductCost>';
		strNode = strNode + '</Product>';
	}
	
	arrxmlOrders.push(strNode);
	return strNode;
}

function addSoap()
{
	soapData = soapData + addNode();
}

function closeSoap()
{

	soapData = soapData + '</Products>';
	soapData = soapData + '</SOAP:Body>';
	soapData = soapData + '</SOAP:Envelope>'
}

function xmlOrder()
{
	var nQty = arrOrders.length;
	if(nQty > 0)
	{
		openSoap();	
		for(i=0;i<=nQty-1;i++)
		{
			addSoap();
		}
		closeSoap();
	}
	sendSOAP();
	if(nQty == 1)
	{
		alert(soapData + "\n\nYou have added " + nQty + " order so far");	
	}
	else if(nQty > 1)
	{
		alert(soapData + "\n\nYou have added " + nQty + " orders so far");
	}
	else
	{
		alert("You have added " + nQty + " orders so far");

	}	
}

function OrderItem(strProductName,nCost) 
{
	this.ProductName = strProductName;
	this.ProductCost = nCost;
}

function addOrderItem(strProductName,nCost)
{
	arrOrders[nArrayItem]=new OrderItem(strProductName,nCost);
	nArrayItem+=1;
}

function addOrders(strProductName,nProductCost)
{		
	addOrderItem(strProductName,nProductCost);	
				
	document.getElementById('CurrentOrders').innerHTML=document.getElementById('CurrentOrders').innerHTML +"<br />"+strProductName + " " + nProductCost;	
}

function clearOrders(CurrentOrders)
{		
	nArrayItem = 0;
	arrxmlOrders = [];
	arrOrders = [];	
	document.getElementById('CurrentOrders').innerHTML = arrOrders;	
}
</script>
<style>
.MainDiv
{	
	margin-left:20px;
	margin-top:50px;
	padding:10px;
	background-color:#C4D6EF;
	width:290px;
	
}
.FontStyle
{
	font-weight:bold;
	text-align:center;
}
.BackgroundColor
{
	background-color:#C4D6EF;
}
</style>
</head>
<body >

<div id="MainDiv" class="MainDiv">
	<div class="FontStyle">
		<label class="FontStyle" id="lblCurrentOrders">Current Orders:</label>
	</div>
	<div id="CurrentOrders">
	
	</div>
	<br />
	<p></p>
	<label id="lblProductName">Product Name: </label><input type="text" id="ProductName" onclick="this.value='';"/>
	<br />
	<label style="margin-right:60px;" id="lblProductCost">Cost: </label><input type="text" id="ProductCost" onclick="this.value='';" />
	<br />
	<p></p>
	<input type="button" id="btnAddOrder" value="Add Orders" onclick="addOrders(document.getElementById('ProductName').value,document.getElementById('ProductCost').value)"/>
	<input type="button" id="btnClearOrders" value="Clear Orders" onclick="clearOrders(document.getElementById('CurrentOrders').value)" />
	<input type="button" id="btnAlertOrders" value="Alert Orders" onclick="xmlOrder()"/>
</div>
</body>
</html>
