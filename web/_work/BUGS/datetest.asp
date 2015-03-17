<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	' dateAdd
	
	sDate = request.querystring("DATE")
	nDate = sDate
	Response.Write "START DATE = " & nDate & "<br><br>"
	i=1
	
	Do While i<13
	  nDate = dateAdd("m",1,nDate)
	  response.write nDate & "<br>"
	  i=i+1
	Loop 
	
	Response.Write "<br>ANNUAL MONTHLY EXPIRY DATE = " & nDate & "<br><br>"
	Response.Write "That's " & DateDiff("d",sDate,nDate) & " days"
%>