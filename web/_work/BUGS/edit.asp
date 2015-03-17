<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Edit</title>
<link rel="stylesheet" href="theme_w3c.css" type="text/css">
<style>
th{font-family:Arial, Helvetica, sans-serif; font-size:10px; background-color:#CCCCCC;}
td{font-family:Arial, Helvetica, sans-serif; font-size:10px;
border-bottom:1px solid #cccccc;border-right:1px solid #cccccc; vertical-align:top;}
.col_one{border-left:1px solid #cccccc;}
.style1 {
	color: #990000;
	font-weight: bold;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
}
</style>
</head>

<body>
<%
If request.querystring("updated")=1 Then
%>
	<span class="style1">DATA SAVED</span><br />
<br />

<%
End If
%>
<form id="form1" name="form1" method="post" action="updateEdit.asp">
<%

function buildProjectsDropDown(nProjectID)
	buildProjectsDropDown = ""
	set connProj = Server.CreateObject("ADODB.Connection")
	connProj.Provider="Microsoft.Jet.OLEDB.4.0"
	connProj.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rsProj=Server.CreateObject("ADODB.recordset")
	
	rsProj.Open "Select * from projects WHERE ProjectID <> 4 ORDER BY projectName ", connProj
	
	do until rsProj.EOF
	
		projectID			= rsProj("projectID")
		projectName			= rsProj("projectName")
		
		If nProjectID = projectID Then
			buildProjectsDropDown = buildProjectsDropDown & "<option value="""&projectID&""" selected>"&projectName&"</option>"
		Else
			buildProjectsDropDown = buildProjectsDropDown & "<option value="""&projectID&""">"&projectName&"</option>"
		End If
		
		
		rsProj.MoveNext
	loop
	
	If Len(nProjectID)=0 Then
		nProjectID = 4 
	End If
	
	If Len(nProjectID)=0 or nProjectID = 4 Then
		buildProjectsDropDown = buildProjectsDropDown & "<option value=""4"" selected>Unassigned</option>"
	Else
		buildProjectsDropDown = buildProjectsDropDown & "<option value=""4"" >Unassigned</option>"
	End If
	
	rsProj.close
	connProj.Close
	
	set connProj = nothing
	
end function

nBugID = request.querystring("ID")

set conn = Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rs=Server.CreateObject("ADODB.recordset")
	rs.Open "Select * from bugs WHERE bugID = "&nBugID&"", conn
	
	do until rs.EOF
	
		bugID				= replaceNull(rs("bugID"))
		bugSection			= replaceNull(rs("bugSection"))
		bugNewZapp			= replaceNull(rs("bugNewZapp"))
		bugEditSite			= replaceNull(rs("bugEditSite"))
		bugReportedDate		= replaceNull(rs("bugReportedDate"))
		bugReportedBy		= replaceNull(rs("bugReportedBy"))
		bugConfirmedDate	= replaceNull(rs("bugConfirmedDate"))
		bugConfirmedBy		= replaceNull(rs("bugConfirmedBy"))
		bugAssignedDate		= replaceNull(rs("bugAssignedDate"))
		bugAssignedTo		= replaceNull(rs("bugAssignedTo"))
		bugFixedDate		= replaceNull(rs("bugFixedDate"))
		bugFixConfirmedBy	= replaceNull(rs("bugFixConfirmedBy"))
		bugNotes			= replaceNull(rs("bugNotes"))
		bugPriority			= replaceNull(rs("bugPriority"))
		bugPriority			= bugPriority/2
		bugDevNotes			= replaceNull(rs("bugDevNotes"))
		
		bugComplete			= replaceNull(rs("bugComplete"))
		bugNotABug			= replaceNull(rs("bugNotABug"))
		
		bugType				= rs("bugType")
		bugProject		= replaceNull(rs("bugProject"))
		If bugComplete = true or bugComplete = "True" Then
			bugComplete = 1
		Else
			bugComplete = 0
		End If
		
		If bugNotABug = true or bugNotABug = "True" Then
			bugNotABug = 1
		Else
			bugNotABug = 0
		End If
		
		bugNotes = replace(bugNotes,"<br>",vbcrlf)
		bugDevNotes = replace(bugDevNotes,"<br>",vbcrlf)

	response.write "<table>"
		
			response.write  "<tr><td>ID:</td><td><input name=""form_bugID"" type=""hidden"" value="""& bugID &""" />"& bugID &" </td></tr>"
			
			response.write  "<tr><td>SECTION:</td><td><input disabled=""disabled"" name=""form_Section"" type=""text"" value="""& bugSection &""" /></td></tr>"
			response.write  "<tr><td>NOTES</td><td><textarea  disabled=""disabled"" name=""form_Notes"" cols=""100"" rows=""6"" />"& bugNotes &"</textarea></td></tr>"
			
			response.write  "<tr><td>REPORTED DATE</td><td>"& bugReportedDate &"</td></tr>"
			response.write  "<tr><td>REPORTED BY</td><td>"& bugReportedBy &"</td></tr>"
			
			response.write  "<tr><td>ASSIGNED TO</td><td><input name=""form_AssignedTo"" type=""text"" value="""& bugAssignedTo &"""/></td></tr>"
			'response.write  "<tr><td>CONFIRMED DATE</td><td>"& bugConfirmedDate &"</td></tr>"
			response.write  "<tr><td>CONFIRMED BY</td><td><input name=""form_ConfirmedBy"" type=""text"" value="""& bugConfirmedBy &"""/></td></tr>"
			'response.write  "<tr><td>FIXED DATE</td><td>"& bugFixedDate &"</td></tr>"
			response.write  "<tr><td>FIXED CONFORMED BY</td><td><input name=""form_FixConfirmedBy"" type=""text"" value="""& bugFixConfirmedBy &""" /></td></tr>"
			
		
			response.write  "<tr><td>PRIORITY</td><td><input name=""form_Priority"" type=""text"" value="""& bugPriority &""" /></td></tr>"
			response.write  "<tr><td>TYPE</td><td><input name=""form_Type"" type=""text"" value="""& bugType &""" /></td></tr>"
			response.write  "<tr><td>DEV NOTES</td><td><textarea name=""form_DevNotes"" cols=""100"" rows=""6"" />"& bugDevNotes &"</textarea></td></tr>"
			response.write  "<tr><td>NEWZAPP BUG</td><td>"& bugNewZapp &"</td></tr>"
			response.write  "<tr><td>EDITSITE BUG</td><td>"& bugEditSite &"</td></tr>"
			response.write  "<tr><td>COMPLETE</td><td><input name=""form_Complete"" type=""text"" value="""& bugComplete &"""/></td></tr>"
			response.write  "<tr><td>NOT A BUG</td><td><input name=""form_NotABug"" type=""text"" value="""& bugNotABug &"""/></td></tr>"
			response.write  "<tr><td>PROJECT</td><td><select name=""form_Project"">"&buildProjectsDropDown(bugProject)&"</select></td></tr>"
			
	response.write "</table>"
		
		
		rs.MoveNext
	loop
	rs.close
	
	conn.Close
	
	set conn = nothing
	
	Function replaceNull(vData)
		If Len(vData)>0 Then
			replaceNull = vData
		Else
			replaceNull = ""
		End If
	End Function
%><br />

<input name="" type="submit" /> <input type="button" value="Return to list" onclick="document.location='viewList.asp?<%response.write session.contents("viewLod")%>'" />
</form>
</body>
</html>
<%Else%>
	ERROR
<%End If%>