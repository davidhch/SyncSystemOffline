<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim strFilePath
strFilePath = server.MapPath("logSoap.txt")

Dim fso, f

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(strFilePath, 8, True)
f.Write "Hello world!" & now()

%>