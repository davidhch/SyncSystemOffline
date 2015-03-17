<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<script>

var constDelay = 2000;

function setEvents(eID)
{
	document.getElementById(eID).onmouseover=function() {
	 if(this.waiting) clearTimeout(this.waiting); 
	 var that = this;
	  this.waiting = setTimeout(function() {that.className="imover";that = null;},constDelay);
	} 
	
	document.getElementById(eID).onmouseout=function() {
		if(this.waiting) clearTimeout(this.waiting);
		var that = this;
	 	this.waiting = setTimeout(function() {that.className="imout";that = null;},constDelay);
	}
}

function initEvents()
{
	document.getElementById("loader").innerHTML = "loading....";
	for (var i = 0; i < document.getElementsByTagName("span").length; i++) { 
      if (document.getElementsByTagName("span")[i].getAttribute('id') != null ){
		setEvents(document.getElementsByTagName("span")[i].getAttribute('id'));
		document.getElementById("loader").innerHTML = "loading...." + document.getElementsByTagName("span")[i].getAttribute('id');
	  }
	}
	document.getElementById("loader").innerHTML = "complete";
}

window.onload = initEvents;
</script>
<style>
.imover{color:#FF3399;}
.imout{color:#666666;}
</style>
</head>
<body>
<div id="loader"></div>
<%

Dim i
For i=1 to 1000

%>
<span id="link<%response.write i%>">hover over and out <%response.write i%></span><br />
<%

Next 

%>
</body>
</html>
