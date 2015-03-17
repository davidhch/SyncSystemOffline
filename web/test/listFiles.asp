<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%   
Whichfolder=server.mappath("\") &"../../../PowerLounge/web-210209/_private/"  
Dim fs, f, f1, fc  
Set fs = CreateObject("Scripting.FileSystemObject")  
Set f = fs.GetFolder(Whichfolder)  
Set fc = f.files 

For Each f1 in fc  
 Response.write (f1.name & "<BR>")  
Next  
%>
 
