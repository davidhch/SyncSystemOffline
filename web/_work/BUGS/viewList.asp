<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then

session.contents("viewLod") = request.ServerVariables("QUERY_STRING")
'response.write session.contents("viewLod")

bShowComplete = request.querystring("ShowComplete")
bPriorityOrder = request.querystring("PriorityOrder")

workID = request.querystring("workID")
workCID = request.querystring("workCID")
workType = request.querystring("workType")
viewType = request.querystring("viewType")
sortType = request.querystring("sortType")
nProjectID = request.querystring("projectID")
nbugSection = request.querystring("bugSection")

priorityOperator = request.querystring("priorityOperator")
priorityLevel = request.querystring("priorityLevel")

If Len(workType) = 0 Then
	workType = 0
End If

If Len(viewType) = 0 Then
	viewType = 0
End If

If Len(sortType) = 0 Then
	sortType = 0
End If

function buildProjectsDropDown(nProjectID)
	buildProjectsDropDown = ""
	set connProj = Server.CreateObject("ADODB.Connection")
	connProj.Provider="Microsoft.Jet.OLEDB.4.0"
	connProj.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rsProj=Server.CreateObject("ADODB.recordset")
	
	rsProj.Open "Select * from projects WHERE projectID <> 4 ORDER BY projectName ", connProj
	
	 
		do until rsProj.EOF
		
			projectID			= rsProj("projectID")
			projectName			= rsProj("projectName")
			
		response.write "<!!--//nProjectID:"&nProjectID&"//-->"
		response.write "<!!--//projectID:"&projectID&"//-->"
		If len(nProjectID)>0 and len(projectID)>0 Then
			If cInt(nProjectID) = cInt(projectID) Then
				buildProjectsDropDown = buildProjectsDropDown & "<option value="""&projectID&""" selected>"&projectName&"</option>"
			Else
				buildProjectsDropDown = buildProjectsDropDown & "<option value="""&projectID&""">"&projectName&"</option>"
			End If
		Else
			buildProjectsDropDown = buildProjectsDropDown & "<option value="""&projectID&""">"&projectName&"</option>"
		End If
			
			rsProj.MoveNext
		loop
		
		If Len(nProjectID) > 0 Then
			If nProjectID = 0 Then
				buildProjectsDropDown = buildProjectsDropDown & "<option value=""0"" selected>Unassigned</option>"
			Else
				buildProjectsDropDown = buildProjectsDropDown & "<option value=""0"">Unassigned</option>"
			End If 
		Else
			buildProjectsDropDown = buildProjectsDropDown & "<option value=""0"">Unassigned</option>"
		End If
	
	
	If Len(nProjectID) > 0 Then 
		buildProjectsDropDown = buildProjectsDropDown & "<option value="""">ALL</option>"
	Else
		buildProjectsDropDown = buildProjectsDropDown & "<option value="""" Selected>ALL</option>"
	End If
	
	rsProj.close
	connProj.Close
	
	set connProj = nothing
	
end function

function buildSectionsDropDown(nSection)
	buildSectionsDropDown = ""
	set connSect = Server.CreateObject("ADODB.Connection")
	connSect.Provider="Microsoft.Jet.OLEDB.4.0"
	connSect.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rsSect=Server.CreateObject("ADODB.recordset")
	
	rsSect.Open "SELECT DISTINCT bugSection from bugs ORDER BY bugSection ", connSect
	do until rsSect.EOF
		
		bugSection			= rsSect("bugSection")
		If Len(bugSection)>0 Then
			If nSection = bugSection Then
				buildSectionsDropDown = buildSectionsDropDown & "<option value="""&bugSection&""" selected>"&bugSection&"</option>"
			Else
				buildSectionsDropDown = buildSectionsDropDown & "<option value="""&bugSection&""">"&bugSection&"</option>"
			End If
		End If
		rsSect.MoveNext
	loop
	
	rsSect.close
	connSect.Close
	
	set connSect = nothing
	
	If Len(nSection) = 0 then
		buildSectionsDropDown = buildSectionsDropDown & "<option value="""" selected>ALL</option>"
	Else
		buildSectionsDropDown = buildSectionsDropDown & "<option value="""" >ALL</option>"
	End If
	
end function

function buildOperatorDropDown(strOperator)
	If strOperator = "<" Then
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""<"" selected><</option>"
	Else
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""<""><</option>"
	End If 
	
	If strOperator = "<=" Then
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""<="" selected><=</option>"
	Else
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""<=""><=</option>"
	End If
	
	If strOperator = "=" Then
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""="" selected>=</option>"
	Else
		buildOperatorDropDown = buildOperatorDropDown & "<option value=""="">=</option>"
	End If 
	
	If strOperator = ">=" Then
		buildOperatorDropDown = buildOperatorDropDown & "<option value="">="" selected>>=</option>"
	Else
		buildOperatorDropDown = buildOperatorDropDown & "<option value="">="">>=</option>"
	End If 
	
	If strOperator = ">" Then
		buildOperatorDropDown = buildOperatorDropDown & "<option value="">"" selected>></option>"
	Else
		buildOperatorDropDown = buildOperatorDropDown & "<option value="">"">></option>"
	End If
	
end function

function buildPriorityDropDown(nPri)
	If nPri = "0" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""0"" selected>0</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""0"">0</option>"
	End If
	
	If nPri = "1" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""1"" selected>1</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""1"">1</option>"
	End If
	
	If nPri = "2" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""2"" selected>2</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""2"">2</option>"
	End If
	
	If nPri = "3" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""3"" selected>3</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""3"">3</option>"
	End If
	
	If nPri = "4" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""4"" selected>4</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""4"">4</option>"
	End If
	
	If nPri = "5" Then
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""5"" selected>5</option>"
	Else
		buildPriorityDropDown = buildPriorityDropDown & "<option value=""5"">5</option>"
	End If
end function

%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>newzapp development tracking</title>
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
document.location.href = "edit.asp?ID="+bugID+"&bList=1";
}

function loadData()
{
	var workID = document.getElementById("workID").value;
	var workCID = document.getElementById("workCID").value;
	var workType = document.getElementById("workType").value;
	var viewType = document.getElementById("viewType").value;
	var sortType = document.getElementById("sortType").value;
	var projectID = document.getElementById("form_Project").value;
	var priorityOperator = document.getElementById("priorityOperator").value;
	var priorityLevel = document.getElementById("priorityLevel").value;
	var bugSection = document.getElementById("bugSection").value;
	
	var strURL = "";
	strURL = "viewList.asp?workID="+workID+"&workType="+workType+"&viewType="+viewType+"&sortType="+sortType+"&projectID="+projectID+"&priorityOperator="+priorityOperator+"&priorityLevel="+priorityLevel+"&bugSection="+bugSection+"&workCID="+workCID+""; 
	
	//alert(strURL);
	document.location = strURL;
}

	var divHeight;
	var divWidth;
	function calcDivHeight()
	{
		divHeight = document.documentElement.offsetHeight - 77;
		divWidth = document.documentElement.offsetWidth;
		
		if ( divHeight > 1 ) 
		{
			document.getElementById("divCont").style.height = divHeight;
			document.getElementById("divCont").style.width = divWidth;
		}
		
		document.getElementById("contTable").style.width = divWidth - 17;
	}
	//calcDivHeight();
	window.onresize = calcDivHeight;
	window.onload = calcDivHeight;
</script>
<link rel="stylesheet" href="theme_w3c.css" type="text/css">
</head>

<body scroll=no>

<h1>Work Management: No Records Displayed = <span id="TotalRecs"></span></h1>
<div style="background-color:#c7d0d9" class="subscribertoolbar">
  <table border="0" cellspacing="0" cellpadding="3">
    <tr>
      <td valign="middle">ID:</td>
      <td><input style="width:40px;" type="text" id="workID" name="workID" value="<%response.write workID%>"/></td>
	  <td>CID:</td>
      <td><input style="width:40px;" type="text" id="workCID" name="workCID" value="<%response.write workCID%>"/></td>
      <td><select name="workType" id="workType">
        <%
		If workType = 0 Then
		%>
        <option value="0" selected>ALL</option>
        <%
		End If
		%>
        <%
		If workType = 1 Then
		%>
        <option value="1" selected>TASK LIST</option>
        <%
		End If
		%>
        <%
		If workType = 2 Then
		%>
        <option value="2" selected>WISH LIST</option>
        <%
		End If
		%>
        <%
		If workType = 3 Then
		%>
        <option value="3" selected>BUG LIST</option>
        <%
		End If
		%>
		 <%
		If workType = 4 Then
		%>
        <option value="4" selected>SUPPORT</option>
        <%
		End If
		%>
        <option value="0">------------</option>
		 <option value="4">SUPPORT</option>
        <option value="3">BUG LIST</option>
        <option value="1">TASK LIST</option>
        <option value="2">WISH LIST</option>
        <option value="0">ALL</option>
      </select></td>
      <td><select name="viewType" id="viewType">
        <%
		If viewType = 0 Then
		%>
        <option value="0" selected>ALL</option>
        <%
		End If
		%>
        <%
		If viewType = 1 Then
		%>
        <option value="1" selected>Not Completed</option>
        <%
		End If
		%>
        <%
		If viewType = 2 Then
		%>
        <option value="2" selected>Completed</option>
        <%
		End If
		%>
        <option value="0">------------</option>
        <option value="2">Completed</option>
        <option value="1">Not Completed</option>
        <option value="0">ALL</option>
      </select></td>
      <td>Sort By:</td>
      <td><select name="sortType" id="sortType">
        <%
		If sortType = 0 Then
		%>
        <option value="0" selected>Bug ID Desc</option>
        <%
		End If
		%>
        <%
		If sortType = 1 Then
		%>
        <option value="1" selected>Section</option>
        <%
		End If
		%>
        <%
		If sortType = 2 Then
		%>
        <option value="2" selected>Priority</option>
        <%
		End If
		%>
        <%
		If sortType = 3 Then
		%>
        <option value="3" selected>Date</option>
        <%
		End If
		%>
        <option value="0">------------</option>
        <option value="3">Date</option>
        <option value="2">Priority</option>
        <option value="1">Section</option>
        <option value="0">Bug ID Desc</option>
      </select></td>
	   <td>Section:</td>
	    <td><select name="bugSection" style="width:100px;" id="bugSection">
        <%response.write buildSectionsDropDown(nbugSection)%>
      </select></td>
	  </tr>
	  </table>
	  </div>
	  <div style="background-color:#c7d0d9" class="subscribertoolbar">
	  <table border="0" cellspacing="0" cellpadding="3">
	  <tr>
      <td>Project:</td>
      <td><select name="form_Project" style="width:150px;" id="form_Project">
        <%response.write buildProjectsDropDown(nProjectID)%>
      </select></td>
	  <td>Priority filter:</td>
	  <td>
	  <select name="priorityOperator" style="width:30px;" id="priorityOperator">
        <%response.write buildOperatorDropDown(priorityOperator)%>
      </select></td>
	  <td>
	  <select name="priorityLevel" style="width:200px;" id="priorityLevel">
        <%response.write buildPriorityDropDown(priorityLevel)%>
      </select></td>
      <td><input id="submitQuery" name="submitQuery" type="button" value="Load Data" onClick="loadData();"  /></td>
      <td>
	  <input type="button" id="Button" name="Button" value="Add New Item" onClick="document.location='input.asp'"/>
	  </td>
    </tr>
  </table>
</div>
<div id="divCont" >
<table id="contTable" border="0" cellspacing="0" cellpadding="0">
  <tr style="position:relative;  top:expression(this.offsetParent.scrollTop); ">
    <th nowrap="nowrap" width="">DEV ID</th>
	
    <th nowrap="nowrap">Section</th>
	<th nowrap="nowrap">Notes</th>
 
    <th nowrap="nowrap">Date Reported </th>
    <th nowrap="nowrap">Found by</th>  
	
	<th nowrap="nowrap">Confirmed by </th>
	
	<th nowrap="nowrap">Assigned</th>

    <th nowrap="nowrap">Fix Checked by</th>
   
	<th nowrap="nowrap">Priority</th>
	

	<th nowrap="nowrap">NZ</th>
    <th nowrap="nowrap">ES</th>
	
	<th nowrap="nowrap">Complete</th>	
	<th nowrap="nowrap">Not a bug</th>
	<th nowrap="nowrap">Type</th>
	<th nowrap="nowrap">Project</th>
	<th nowrap="nowrap">&nbsp;</th>

  </tr>
 


<%


set conn = Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open server.mappath("/_work/bugs/_private/bugs.mdb")
	
	set rs=Server.CreateObject("ADODB.recordset")
	'rs.Open "Select * from bugs WHERE bugType = 2 ORDER BY bugPriority, bugSection, bugID ASC", conn
	If workType = 0 Then
	strSQL = "SELECT bugs.*, projects.* FROM bugs LEFT JOIN projects ON bugs.bugProject = projects.ProjectID WHERE bugType <> -1 "
	Else
	strSQL = "SELECT bugs.*, projects.* FROM bugs LEFT JOIN projects ON bugs.bugProject = projects.ProjectID WHERE bugType = "&workType&" "
	End If
	'ORDER BY bugSection, bugPriority, bugID ASC
	
	

	If viewType = 1 Then
		strSQL = strSQL & " AND bugComplete <> true "
	End If
	
	If viewType = 2 Then
		strSQL = strSQL & " AND bugComplete = true "
	End If
	
	If Len(workID)> 0 Then
		strSQL = strSQL & " AND bugID = "&workID&" "
	End If
	
	If Len(workCID)> 0 Then
		strSQL = strSQL & " AND bugCID = "&workCID&" "
	End If
	
	If Len(nbugSection)> 0 Then
		strSQL = strSQL & " AND bugSection = '"&nbugSection&"' "
	End If

	If Len(nProjectID)> 0 Then
		If nProjectID = 0 Then
			strSQL = strSQL & " AND (ProjectID is null or ProjectID = 4) "
		Else
			strSQL = strSQL & " AND ProjectID = "&nProjectID&" "
		End If
	End If
	
	If Len(priorityOperator) > 0 AND Len(priorityLevel) > 0 Then 
		strSQL = strSQL & " AND bugPriority "&priorityOperator&" "&(priorityLevel*2)&" "
	End If

	If sortType = 0 Then
		strSQL = strSQL & " ORDER BY bugID ASC "
	End If
	
	If sortType = 1 Then
		strSQL = strSQL & " ORDER BY bugSection, bugPriority, bugID ASC "
	End If
	
	If sortType = 2 Then
		strSQL = strSQL & " ORDER BY bugPriority, bugSection, bugID ASC "
	End If
	
	If sortType = 3 Then
		strSQL = strSQL & " ORDER BY bugReportedDate DESC "
	End If
	
	
	response.write "<!--//" & strSQL & "//-->"
	rs.Open strSQL, conn
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
		bugCID				= replaceNull(rs("bugCID"))
		bugType				= replaceNull(rs("bugType"))
		bugDevNotes			= replaceNull(rs("bugDevNotes"))
		
		bugComplete			= replaceNull(rs("bugComplete"))
		bugNotABug			= replaceNull(rs("bugNotABug"))
		
		projectName			= replaceNull(rs("ProjectName"))
		
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
			response.write  "<tr><td colspan=16 class=""SectionHeader"">"&bugSection&"</td></tr>"
		End if
		
		bugNotes = replace(bugNotes,"<","&lt;")
		bugNotes = replace(bugNotes,">","&gt;")
		
		bugNotes = replace(bugNotes,"&lt;br&gt;","<br>")
		bugNotes = replace(bugNotes,"&lt;BR&gt;","<br>")
		
		response.write  "<tr>"
			response.write  "<td class=""col_one"">"& bugID &"</td>"
			
			response.write  "<td nowrap>"& bugSection &"</td>"
			response.write  "<td><div style=""width:220px; height:100px; overflow:auto;""><div style=""width:200px;"">CID: "&bugCID&"<br> "& bugNotes &"</div></div><br><div style=""width:200px; color:#006699"">"& bugDevNotes &"</div></td>"
			
			response.write  "<td>"& bugReportedDate &"</td>"
			response.write  "<td >"& bugReportedBy &"</td>"
			
			response.write  "<td >"& bugConfirmedBy &"</td>"
			
			response.write  "<td >"& bugAssignedTo &"</td>"
			
			response.write  "<td >"& bugFixConfirmedBy &"</td>"
			If bugPriority <= 10 Then
				bgCol = "#cccccc"
			End If	
			
			If bugPriority <= 9 Then
				bgCol = "#E5E566"
			End If	
			
			If bugPriority <= 8 Then
				bgCol = "#ffff00"
			End If	
			
			If bugPriority <= 6 Then
				bgCol = "#ffcc00"
			End If	
			
			If bugPriority <= 5 Then
				bgCol = "#ff9900"
			End If	
			
			If bugPriority <= 4 Then
				bgCol = "#ff6600"
			End If	
			
			If bugPriority <= 2 Then
				bgCol = "#ff0000"
			End If	
				
			
			response.write  "<td bgColor="""&bgCol&""">"& bugPriority/2 &"</td>"
			'response.write  "<td>"& bugDevNotes &"</td>"
			
			
			response.write  "<td>"& bugNewZapp &"</td>"
			response.write  "<td>"& bugEditSite &"</td>"
			
			response.write  "<td>"& bugComplete &"</td>"
			response.write  "<td>"& bugNotABug &"</td>"
			response.write "<td>"
			If bugType = 1 Then
				response.write "task"
			End If
			If bugType = 2 Then
				response.write "wish"
			End If
			If bugType = 3 Then
				response.write "bug"
			End If
			If bugType = 4 Then
				response.write "support"
			End If
			If Len(bugType) = 0 Then
				response.write "unknown"
			End If
			
			response.write "</td>"
			response.write  "<td>"&projectName&"</td>"
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
</div>
<script>
function loadTotal()
{
document.getElementById("TotalRecs").innerHTML = <%response.write (n-1)%>;
}

loadTotal();
</script>

</body>
</html>
<%Else%>
	ERROR
<%End If%>