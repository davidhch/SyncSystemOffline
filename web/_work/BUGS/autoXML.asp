<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Session.LCID=2057

set conn = Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rs=Server.CreateObject("ADODB.recordset")

	strSQL = "SELECT * FROM bugs FOR XML AUTO"
'response.write strSQL
	rs.Open strSQL, conn
	strTROUT = ""
	nCalls = 0
	nIncomplete = 0
	do until rs.EOF
	
	rs.MoveNext
	loop
	rs.close
	
	conn.Close
	
	set conn = nothing

%>