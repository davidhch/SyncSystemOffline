<!--#include virtual="\_work\bugs\_private\cMDB.asp"-->
<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then%>
<%
	Session.LCID = 2057
	'Read the form values
	form_bugID = request.form("form_bugID")
	form_AssignedTo 		= request.form("form_AssignedTo")
	form_ConfirmedBy 		= request.form("form_ConfirmedBy")
	form_FixConfirmedBy 		= request.form("form_FixConfirmedBy")
	form_Priority 		= request.form("form_Priority")
	form_Priority = form_Priority*2
	form_DevNotes 	= request.form("form_DevNotes")
	
	form_Complete 	= request.form("form_Complete")
	form_NotABug 	= request.form("form_NotABug")
	
	form_DevNotes = replace(form_DevNotes,"'","''")
	form_DevNotes = replace(form_DevNotes,vbcrlf,"<br>")
	
	form_Type 	= request.form("form_Type")
	
	form_Project		= request.form("form_Project")
	
	' build the SQL
	If Len(form_Type) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugType="&form_Type&" "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_Project) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugProject="&form_Project&" "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_AssignedTo) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugAssignedTo='"&form_AssignedTo&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_ConfirmedBy) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugConfirmedBy='"&form_ConfirmedBy&"', bugConfirmedDate='"&now()&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_FixConfirmedBy) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugFixConfirmedBy='"&form_FixConfirmedBy&"', bugFixedDate='"&now()&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If

	If Len(form_Priority) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugPriority="&form_Priority&" "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_DevNotes) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugDevNotes='"&form_DevNotes&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_Complete) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugComplete="&form_Complete&", bugCompletedDate='"&now()&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	If Len(form_NotABug) > 0 Then
		strSQL = ""
		strSQL = strSQL & "UPDATE bugs "
		strSQL = strSQL & "SET bugNotABug="&form_NotABug&", bugCompletedDate='"&now()&"' "
		strSQL = strSQL & "WHERE bugID="&form_bugID&" "
		
		updateRecord strSQL
	End If
	
	Sub updateRecord(strSQL)
		response.write strSQL & "<br><br>"
		
		Dim objDB
		Set objDB = New CDBConn
		
		
		
		objDB.openDBConnection
		objDB.executeSQL strSQL
		objDB.closeDBConnection
	End Sub
	response.redirect "edit.asp?ID="&form_bugID&"&updated=1"
	
%>
<%Else%>
	ERROR
<%End If%>