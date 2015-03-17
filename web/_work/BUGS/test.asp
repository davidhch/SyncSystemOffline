<!--#include virtual="\_work\bugs\_private\cMDB.asp"-->
<%
	Session.LCID = 2057
	'Read the form values
	strSection 		= request.form("form_Section")
	bNewZapp 		= request.form("form_NewZapp")
	bEditSite 		= request.form("form_EditSite")
	dReported 		= request.form("form_ReportedDate")
	strReportedBy 	= request.form("form_ReportedBy")
	strNotes 		= request.form("form_Notes")
	nPriority 		= request.form("form_Priority")
	
	nFormCID = request.form("form_CID")
	
	form_Type 		= request.form("form_Type")
	
	bNotABug = 1
	nPriority = nPriority*2
	
	dReported = CDate(dReported)
	strNotes = replace(strNotes,"'","''")
	strNotes = replace(strNotes,vbcrlf,"<br>")
	' build the SQL
	strSQL = ""
	strSQL = strSQL & "INSERT INTO "
	strSQL = strSQL & "bugs (bugSection,bugNewZapp,bugEditSite,bugReportedDate,bugReportedBy,bugNotes,bugType,bugPriority,bugCID) "
	strSQL = strSQL & "VALUES ('"&strSection&"',"&bNewZapp&","&bEditSite&",'"&dReported&"','"&strReportedBy&"','"&strNotes&"',"&form_Type&","&nPriority&","&nFormCID&")"
	
	'response.write strSQL
	
	Dim objDB
	Set objDB = New CDBConn
	
	objDB.openDBConnection
	objDB.executeSQL strSQL
	objDB.closeDBConnection
	
	response.redirect "default.asp"
	
%>