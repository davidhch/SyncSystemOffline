// JavaScript Document

function object(original) {
	function F() {}
	F.prototype = original;
	return new F();
};

var nzObjBrowserDetection = {
	nzBrowser: function(){
		var strBrowser = "";
		if (/MSIE (\d+\.\d+);/.test(navigator.userAgent)){ //test for MSIE x.x;
			var ieversion=new Number(RegExp.$1) // capture x.x portion and store as a number
			if (ieversion>=8){
				//document.write("You're using IE8 or above")
				strBrowser = "IE8+";	  
			}
			else if (ieversion>=7){
				//document.write("You're using IE7.x")
				strBrowser = "IE7";
			}
			else if (ieversion>=6){
				// document.write("You're using IE6.x")
				strBrowser = "IE6";
			}
			else if (ieversion>=5){
				//document.write("You're using IE5.x")
				strBrowser = "IE5";
			}	  
		}
		else if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)){ //test for Firefox/x.x or Firefox x.x (ignoring remaining digits);
			var ffversion=new Number(RegExp.$1) // capture x.x portion and store as a number
			if (ffversion>=3){
				//document.write("You're using FF 3.x or above")
				strBrowser = "FF3+";}
			else if (ffversion>=2){
				//document.write("You're using FF 2.x")
				strBrowser = "FF2";}
			else if (ffversion>=1){
				//document.write("You're using FF 1.x")
				strBrowser = "FF1";}
			}
		else if (/Opera[\/\s](\d+\.\d+)/.test(navigator.userAgent)){ //test for Opera/x.x or Opera x.x (ignoring remaining decimal places);
			var oprversion=new Number(RegExp.$1) // capture x.x portion and store as a number
			if (oprversion>=10){
				//document.write("You're using Opera 10.x or above")
				strBrowser = "OP10+";}
			else if (oprversion>=9){
				//document.write("You're using Opera 9.x")
				strBrowser = "OP9";}
			else if (oprversion>=8){
				//document.write("You're using Opera 8.x")
				strBrowser = "OP8";}
			else if (oprversion>=7){
				//document.write("You're using Opera 7.x")
				strBrowser = "OP7";}
			}
		else
		{
			strBrowser = "n/a";
		}
		
		return strBrowser;
	}
};


var nzObjBrowser = new object(nzObjBrowserDetection);
var nzBrowser = nzObjBrowser.nzBrowser();

switch(nzBrowser)
{
	case "IE7":
		//alert("nzBrowser: " + nzBrowser);
		/*document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzIEConstants.js'></script>");
		document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzIEObjects.js'></script>");		
		document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzGetAndSetMethods.js'></script>");*/
		
		document.write("<script language='JavaScript' type='text/javascript' src='nzIEConstants.js'></script>");
		document.write("<script language='JavaScript' type='text/javascript' src='IEJSFramework.js'></script>");
		/*document.write("<script language='JavaScript' type='text/javascript' src='/_work/jsFramework/nzXMLDocument.js'></script>");
		document.write("<script language='JavaScript' type='text/javascript' src='/_work/jsFramework/nzXSLDocument.js'></script>");
		*/
		
	break;
	case "FF3+":
		//alert("nzBrowser: " + nzBrowser);
		/*document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzFFConstants.js'></script>");
		document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzFFObjects.js'></script>");		
		document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzGetAndSetMethods.js'></script>");*/
	document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptJQueryTest/NewFFJSTest.js'></script>");
	break;
}

/*if (nzBrowser == "IE7")
{
	alert("nzBrowser: " + nzBrowser);
	document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzIEObjElements.js'></script>");
}
else if (nzBrowser == "FF3+")
{
	alert("nzBrowser: " + nzBrowser);
	document.write("<script language='JavaScript' type='text/javascript' src='/test/mh/JavascriptPrototypeTest/nzFFObjElements.js'></script>");
}*/
