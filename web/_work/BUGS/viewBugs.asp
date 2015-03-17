<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<style>
th{font-family:Arial, Helvetica, sans-serif; font-size:11px; background-color:#CCCCCC; padding:5px;}
td{font-family:Arial, Helvetica, sans-serif; font-size:13px;  padding:5px;
border-bottom:1px solid #cccccc;border-right:1px solid #cccccc; vertical-align:top}
.col_one{border-left:1px solid #cccccc;}
body{padding:0px; margin:0px;}
input{font-family:Arial, Helvetica, sans-serif; font-size:13px;}
textarea{font-family:Arial, Helvetica, sans-serif; font-size:13px;border:0;}
.SectionHeader{font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:bold; background-color:#eeeeee}
</style>
<script>
function editRecord(bugID)
{
//alert(bugID);
document.location = "edit.asp?ID="+bugID+"";
}
</script>
</head>

<body>
<table  border="0" cellspacing="0" cellpadding="0">
  <tr style="position:relative;  top:expression(this.offsetParent.scrollTop); ">
    <th nowrap="nowrap">bug No</th>
	
    <th nowrap="nowrap">Section</th>
	<th nowrap="nowrap">Notes</th>
 
    <th nowrap="nowrap">Date Reported </th>
    <th nowrap="nowrap">Found by</th>  
	
	<th nowrap="nowrap">Confirmed by </th>
	
	<th nowrap="nowrap">Assigned</th>
	 
<!--// <th nowrap="nowrap">Date Confirmed</th> //-->
  

   <!--//	<th nowrap="nowrap">Date Fixed</th> //-->
    <th nowrap="nowrap">Fix Checked by</th>
    
	<th nowrap="nowrap">Priority</th>
	
	
	<th nowrap="nowrap">NZ</th>
    <th nowrap="nowrap">ES</th>
	
	<th nowrap="nowrap">Complete</th>	
	<th nowrap="nowrap">Not a bug</th>
	
	<th nowrap="nowrap"><a href="input.asp">ADD</a></th>
  </tr>
 


<%
set conn = Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rs=Server.CreateObject("ADODB.recordset")
	rs.Open "Select * from bugs ORDER BY bugID DESC", conn
	'WHERE bugAssignedTo = 'dh' 
	n = 1
	nEqual = 1
	prevBugSection = ""
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
		bugDevNotes			= replaceNull(rs("bugDevNotes"))
		
		bugComplete			= replaceNull(rs("bugComplete"))
		bugNotABug			= replaceNull(rs("bugNotABug"))
		
		'bugNotes = replace(bugNotes,"<br>",vbcrlf)
		'bugDevNotes = replace(bugDevNotes,"<br>",vbcrlf)
		
		If Len(bugAssignedTo) > 0 Then
			bugAssignedTo = replaceWithImages(bugAssignedTo)
		Else
			bugAssignedTo = "-"
		End If
		
		
		If Len(bugReportedBy) > 0 Then
			bugReportedBy = replaceWithImages(bugReportedBy)
		Else
			bugReportedBy = "-"
		End If
		
		
		If Len(bugConfirmedBy) > 0 Then
			bugConfirmedBy = replaceWithImages(bugConfirmedBy)
		Else
			bugConfirmedBy = "-"
		End If
		
		If Len(bugFixConfirmedBy) > 0 Then
			bugFixConfirmedBy = replaceWithImages(bugFixConfirmedBy)
		Else
			bugFixConfirmedBy = "-"
		End If
		
		
		If bugNewZapp = true or bugNewZapp = "True" Then
			bugNewZapp = "<img src=""iconTick.gif"" />"
		Else
			bugNewZapp = "&nbsp;"
		End If
		
		If bugEditSite = true or bugEditSite = "True" Then
			bugEditSite = "<img src=""iconTick.gif"" />"
		Else
			bugEditSite = "&nbsp;"
		End If
		
		If bugComplete = true or bugComplete = "True" Then
			bugComplete = "<img src=""iconTick.gif"" />"
		Else
			bugComplete = "&nbsp;"
		End If
		
		If bugNotABug = true or bugNotABug = "True" Then
			bugNotABug = "<img src=""iconTick.gif"" />"
		Else
			bugNotABug = "&nbsp;"
		End If
		
		If prevBugSection <> bugSection Then
			'response.write  "<tr><td colspan=14 class=""SectionHeader"">"&bugSection&"</td></tr>"
		End if
		
		bugNotes = replace(bugNotes,"<","&lt;")
		bugNotes = replace(bugNotes,">","&gt;")
		
		bugNotes = replace(bugNotes,"&lt;br&gt;","<br>")
		bugNotes = replace(bugNotes,"&lt;BR&gt;","<br>")
		
		response.write  "<tr>"
			response.write  "<td class=""col_one"">"& bugID &"</td>"
			
			response.write  "<td nowrap>"& bugSection &"</td>"
			response.write  "<td><div style=""width:200px;"">"& bugNotes &"</div><br><div style=""width:200px; color:#006699"">"& bugDevNotes &"</div></td>"
			
			response.write  "<td>"& bugReportedDate &"</td>"
			response.write  "<td >"& bugReportedBy &"</td>"
			
			response.write  "<td >"& bugConfirmedBy &"</td>"
			
			response.write  "<td >"& bugAssignedTo &"</td>"
			'response.write  "<td>"& bugConfirmedDate &"</td>"
			
			'response.write  "<td>"& bugFixedDate &"</td>"
			response.write  "<td >"& bugFixConfirmedBy &"</td>"
			
			response.write  "<td>"& bugPriority/2 &"</td>"
			'response.write  "<td>"& bugDevNotes &"</td>"
			
			
			response.write  "<td>"& bugNewZapp &"</td>"
			response.write  "<td>"& bugEditSite &"</td>"
			
			response.write  "<td>"& bugComplete &"</td>"
			response.write  "<td>"& bugNotABug &"</td>"
			
			response.write  "<td><input name="""" type=""button"" onclick=""editRecord("&bugID&")"" value=""Edit"" /></td>"
		response.write   "</tr>"
		
		
		n = n+1
		
		prevBugSection = bugSection
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
	
	Function replaceWithImages(employees)
				'split
				If 1=2 Then
				replaceWithImages = ""
				arrEmps = split(employees,",")
				
				For i = LBound(arrEmps) To UBound(arrEmps)
					replaceWithImages = replaceWithImages & "<img src="""&arrEmps(i)&".gif"" align=""absmiddle""/> "' & arrEmps(i)
				Next
				
				Else
					replaceWithImages = employees
				End If
				
	End Function
%>
</table>
</body>
</html>
<%Else%>
	ERROR
<%End If%>