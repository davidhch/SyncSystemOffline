<%
Set myMail=CreateObject("CDO.Message")
myMail.Subject="Test Email"
myMail.From="dcurnow@gmail.com"
myMail.To="dcurnow@gmail.com"
myMail.TextBody="Test nbs."
myMail.Send
set myMail=nothing

%>