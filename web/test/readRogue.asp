<html>
<body>
<p>This is the text in the text file:</p>
<%
Set fs=Server.CreateObject("Scripting.FileSystemObject")

Set f=fs.OpenTextFile(Server.MapPath("rogue.txt"), 1)
Response.Write(f.ReadAll)
f.Close

Set f=Nothing
Set fs=Nothing
%>
</body>
</html>