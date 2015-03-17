<%
'////////////////////////////////////////////////////////////////////////////
'// Super Admin Config File   										 ////////
'////////////////////////////////////////////////////////////////////////////

Dim strIPs
Dim strBreaker
Dim arrIPs
Dim bSuperAdminAccess

strIPs = ""
strBreaker = ""

'////////////////////////////////////////////////////////////////////////////
' Allowed IPs
addIP "192.168.0."	' DestiNet Exeter Office
addIP "91.84.227.61"	' DestiNet Exeter Office

bSuperAdminAccess = isSuperAdmin(strIPs)

'//////////////////////////////////////////////////////////////////////////////

Sub addIP(strIP)
	strIPs = strIPs & strBreaker & strIP
	strBreaker = ","
End Sub

Function isSuperAdmin(str)
	isSuperAdmin = False
	arrIPs = split(str,",")
	
	For i = LBound(arrIPs) To UBound(arrIPs)
	'response.write Request.ServerVariables("remote_addr") & "<br>"
		'If Request.ServerVariables("remote_addr") = arrIPs(i) Then
		If InStr(Request.ServerVariables("remote_addr"),arrIPs(i)) Then
			isSuperAdmin = True
		End If
	Next
	
	If Session.Contents("Plus") AND Session.Contents("SuperAdmin") Then

	Else
		'isSuperAdmin = False
	End If
End Function
%>