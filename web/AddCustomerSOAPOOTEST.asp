<%
	strLine = ""
	strLine = strLine & "FILE STARTED AddCustomerSoapOO.asp"
	logSoap(strLine)
	
	
	Sub logSoap(strLine)
		dim strFilePath
		strFilePath = server.MapPath("logSoap.txt")
		
		Dim fso, f
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(strFilePath, 8, True)
		f.Write strLine & " - " & vbcrlf & now() & vbcrlf & "----------------" & vbcrlf & vbcrlf
	End Sub
	
	Response.write "Yo Dube"
%>