// JavaScript Document

var nzXSLDocument = { // Start of nzXMLDocument constants to be used
	getXSLDoc: function(func){
		var nzConstXSLDocPath = "";
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
				nzConstXSLDocPath = strHTTP + strGlobalSystemURL + "/xsl/ViewAutoChargingData.asp"
				break;
			case "EmailAddress_Test":
				nzConstXSLDocPath = ""
				break;
			case "GetButtons":
				nzConstXSLDocPath = strHTTP + strGlobalSystemURL + "/test/mh/NZJSFramework/ViewButtonData.asp"
				break;
			case "SetOnClickForID":
				nzConstXSLDocPath = strHTTP + strGlobalSystemURL + "/test/mh/NZJSFramework/ViewSetOnClickForID.asp"
				break;
		}
		//alert("nzConstXSLDocPath: " + nzConstXSLDocPath);
		return nzConstXSLDocPath;
	}
} // End of nzXMLDocument constants