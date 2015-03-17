<%@ Language=VBScript %>
<%

nURL = Request.Querystring("n")


strURL =  "http://www.newzapp.co.uk"
If nURL=1 then
	strURL = "http://www.destinet.co.uk"
End if

Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", strURL
%>