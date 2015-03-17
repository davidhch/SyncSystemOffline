<%
   Select Case Request.Form("Name")
   		Case "default"
         Server.Transfer "/_work/"
		 
      Case "flex"
         Server.Transfer "page4.asp"
	  Case "trial"
         Server.Transfer "page4.asp"
      Case Else
         Server.Transfer "page3.asp" ' why not have a default
   End Select
%>