<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dSmallRenewalDate = "24/10/2006"
If isDate(dSmallRenewalDate) Then
	dRenewalDate = FormatDateTime(dSmallRenewalDate,vbshortdate)
Else
	dRenewalDate = "BAD DATE"
End If
response.write dRenewalDate
%>