<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Session.LCID=2057

set conn = Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rs=Server.CreateObject("ADODB.recordset")
	
	dLastWeek = DateAdd("d",-4,date())
	strSQL = "SELECT bugs.*, projects.* FROM bugs LEFT JOIN projects ON bugs.bugProject = projects.ProjectID WHERE bugType = 4 AND bugReportedDate > #"&day(dLastWeek)&" "&monthname(month(dLastWeek))&" "&year(dLastWeek)&"# ORDER BY bugReportedDate DESC"
'response.write strSQL
	rs.Open strSQL, conn
	strTROUT = ""
	nCalls = 0
	nIncomplete = 0
	do until rs.EOF
	
		bugID				= replaceNull(rs("bugID"))
		bugSection			= replaceNull(rs("bugSection"))
		bugNewZapp			= replaceNull(rs("bugNewZapp"))
		bugEditSite			= replaceNull(rs("bugEditSite"))
		bugReportedDate		= replaceNull(rs("bugReportedDate"))
		bugReportedBy		= (rs("bugReportedBy"))
		bugConfirmedDate	= replaceNull(rs("bugConfirmedDate"))
		bugConfirmedBy		= (rs("bugConfirmedBy"))
		bugAssignedDate		= replaceNull(rs("bugAssignedDate"))
		bugAssignedTo		= (rs("bugAssignedTo"))
		bugFixedDate		= replaceNull(rs("bugFixedDate"))
		bugFixConfirmedBy	= (rs("bugFixConfirmedBy"))
		bugNotes			= replaceNull(rs("bugNotes"))
		bugPriority			= replaceNull(rs("bugPriority"))
		bugCID				= replaceNull(rs("bugCID"))
		bugType				= replaceNull(rs("bugType"))
		bugDevNotes			= replaceNull(rs("bugDevNotes"))
		
		bugComplete			= replaceNull(rs("bugComplete"))
		bugNotABug			= replaceNull(rs("bugNotABug"))
		
		projectName			= replaceNull(rs("ProjectName"))
		bugCompletedDate	= replaceNull(rs("bugCompletedDate"))
		
			strTROUT = strTROUT & "<tr>"
		strTROUT = strTROUT & "<td nowrap valign=""top"">"&bugReportedDate&"</td>"
		'strTROUT = strTROUT & "<td valign=""top"">"&bugID&"</td>"
		strTROUT = strTROUT & "<td valign=""top"">"&bugSection&"</td>"
		strTROUT = strTROUT & "<td valign=""top"">"&bugReportedBy&"</td>"
		'strTROUT = strTROUT & "<td valign=""top"">"&bugNotes&"</td>"
		strTROUT = strTROUT & "<td valign=""top"">"&bugCID&"</td>"
		strTROUT = strTROUT & "<td valign=""top"">"&bugComplete&"</td>"
		'strTROUT = strTROUT & "<td valign=""top"">"&bugCompletedDate&"</td>"
		strTROUT = strTROUT & "<td valign=""top""><a href=""http://orderentry-soap.destinet.co.uk/_work/bugs/viewList.asp?workID="&bugID&""">S-"&bugID&"</a></td>"
		
			strTROUT = strTROUT & "</tr>"
			
			IF NOT bugComplete Then
				nIncomplete = nIncomplete +1
			End IF
	
	nCalls = nCalls +1
	
	rs.MoveNext
	loop
	rs.close
	
	conn.Close
	
	set conn = nothing
	
	Function replaceNull(vData)
		If Len(vData)>0 Then
			replaceNull = vData
		Else
			replaceNull = "-"
		End If
	End Function
%>
<h3><%response.write nCalls%> SUPPORT REQUESTS for period <%response.write WeekdayName(Weekday(dLastWeek)) %>&nbsp;<%response.write dLastWeek%> to <%response.write WeekdayName(Weekday(date())) %>&nbsp;<%response.write date()%></h3>
Detail of request can be accessed by the ID link. <%if nCalls> 0 and 1=2 then response.write round(100*(nIncomplete/nCalls))&"% incomplete" end if%>
<br>
<br>

<table border="1">
<tr>
<th nowrap>REPORTED</th>
<th nowrap>SECTION</th>
<th nowrap>REPORTED BY</th>
<th nowrap>CID</th>
<th nowrap>COMPLETE</th>
<th nowrap>REF ID</th>


</tr>
<%response.write strTROUT%>
</table>
