<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' create test csv / xsl files for stress testing custom fields imports
response.Buffer = true
nMax = request.querystring("n")

response.write "Email Address, First Name, Last Name, A String, A Number, A Bool, A Date, Another String, Another Number, Another Bool, Another Date"&vbcrlf

Dim i
For i=1 to nMax
	
	e = "testEmail"&i&"@destinet.co.uk"
	f = "first"&i&""
	l = "last"&i&""
	s1 = "string"&i&""
	n1 = 7
	b1 = "0"
	d1 = "09/08/2008"
	s2 = "String 2 "&i&""
	n2 = 8
	b2 = "yes"
	d2 = "May 15th 2008" 
	
	response.write ""&e&", "&f&", "&l&", "&s1&", "&n1&", "&b1&", "&d1&", "&s2&", "&n2&", "&b2&", "&d2&""&vbcrlf
	response.Flush()
Next 
%>