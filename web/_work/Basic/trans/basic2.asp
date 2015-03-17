<% 
Dim xmldoc 
Dim xsldoc

'Use the MSXML 4.0 Version dependent PROGID 
'MSXML2.DOMDocument.4.0 if you wish to create
'an instance of the MSXML 4.0 DOMDocument object
 
Set xmldoc = Server.CreateObject("MSXML2.DOMDocument") 
Set xsldoc = Server.CreateObject("MSXML2.DOMDocument")

xmldoc.Load Server.MapPath("dcMHoutputb.xml")

'Check for a successful load of the XML Document.
if xmldoc.parseerror.errorcode <> 0 then 
  Response.Write "Error loading XML Document :" & "<BR>"
  Response.Write "----------------------------" & "<BR>"
  Response.Write "Error Code : " & xmldoc.parseerror.errorcode & "<BR>"
  Response.Write "Reason : " & xmldoc.parseerror.reason & "<BR>"
  Response.End 
End If


xsldoc.Load Server.MapPath("reportExport_EACUsimple.asp")

'Check for a successful load of the XSL Document.
if xsldoc.parseerror.errorcode <> 0 then 
  Response.Write "Error loading XSL Document :" & "<BR>"
  Response.Write "----------------------------" & "<BR>"
  Response.Write "Error Code : " & xsldoc.parseerror.errorcode & "<BR>"
  Response.Write "Reason : " & xsldoc.parseerror.reason & "<BR>"
  Response.End 
End If

Response.Write xmldoc.TransformNode(xsldoc)

%>