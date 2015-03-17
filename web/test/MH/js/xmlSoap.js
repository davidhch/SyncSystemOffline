// JavaScript Document
var strResponse = "";
function xmlParseSoap(strSoap,strURL)
{	
	var objSoap = new XMLHttpRequest();

	objSoap.open("POST", ""+strURL+"", true);
	objSoap.setRequestHeader("Man", ""+strURL+" HTTP/1.1");
	objSoap.setRequestHeader("MessageType", "CALL");
	objSoap.setRequestHeader("Content-Type", "text/xml");
	objSoap.setRequestHeader('If-Modified-Since','Wed, 15 Nov 1995 04:58:08 GMT');
	objSoap.send(strSoap);

	objSoap.onreadystatechange = function() 
	{		
		if (objSoap.readyState == 4 && objSoap.status == 200) 
		{		
			//strResponse = objSoap.responseXML;
			strResponse = objSoap.responseText;
			returnSoap();
		}

		if (objSoap.readyState == 4 && objSoap.status != 200) 
		{
			strResponse = "Error: "+objSoap.status+" \n "+objSoap.responseText;
			returnSoap();
		}
	}
}