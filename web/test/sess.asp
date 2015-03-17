<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'session
If Len(session.contents("nFigure"))>0 Then
	session.contents("nFigure") = session.contents("nFigure") + 1
Else
	session.contents("nFigure") = 0
End If

response.write session.contents("nFigure")
%>