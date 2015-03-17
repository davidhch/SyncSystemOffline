<% @LANGUAGE="VBSCRIPT" %>
<h1>page3.asp</h1>
<p>you entered <% = Request.Form("Name") %>... and thats not on the list</p>
<br>
LEN = <% = Len(request.form) %>

sender = <% = Request.ServerVariables ("HTTP_REFERER")%>
</body>
</html>