// JavaScript Document

var nzXMLDocument = { // Start of nzXMLDocument constants to be used
	getXMLDoc: function(func){
		var nzConstXMLDocPath = "";
		/*if(secure == true){
			var nzConstHTTPPath = "https://";
			//var nzConstServerPath = "system.newzapp.co.uk";
		}
		else{
			var nzConstHTTPPath = "http://";
			//var nzConstServerPath = "dev210.newzapp.co.uk";
		}*/

		switch(func)
		{
			case "AutoChargingData":
				nzConstXMLDocPath = strHTTP + strGlobalSystemURL + "/superadmin/getAutoChargingData.asp"
				break;
			case "EmailAddress_Test":
				nzConstXMLDocPath = strHTTP + strGlobalSystemURL + "/WebServices/RegExService.asmx"
				break;
			case "GetButtons":
				nzConstXMLDocPath = strHTTP + strGlobalSystemURL + "/test/mh/NZJSFramework/GetButtons.asp"
				break;
			case "SetOnClickForID":
				nzConstXMLDocPath = strHTTP + strGlobalSystemURL + "/test/mh/NZJSFramework/SetOnClickForID.asp"
				break;
		}
		//alert("nzConstXMLDocPath: " + nzConstXMLDocPath);
		//alert("strHTTP: " + strHTTP);
		
		return nzConstXMLDocPath;
	}
} // End of nzXMLDocument constants