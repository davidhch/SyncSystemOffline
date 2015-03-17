<%
Dim xmldoc 
Dim xsldoc
 
Set xmldoc = Server.CreateObject("MSXML2.DOMDocument") 
Set xsldoc = Server.CreateObject("MSXML2.DOMDocument")

xmldoc.Load Server.MapPath("xmldata/start.asp") ' xml delivered via a processing page
xsldoc.Load Server.MapPath("xsl/start.xsl")

Response.Write xmldoc.TransformNode(xsldoc)
%>
