Sub MyASPJob()

Dim oXMLHttp
Dim sURL

on error resume next

Set oXMLHttp = CreateObject("MSXML2.XMLHTTP.3.0")
sURL = "http://orderentry-soap.destinet.co.uk/OrderEntrySyncChecker.asp"
oXMLHttp.open "GET", sURL, false
oXMLHttp.send()

'if oXMLHttp.status = 200 Then
' Retrieve enter HTML response
'MsgBox oXMLHttp.ResponseText
'else
' Failed
'end if

Set oXMLHttp = nothing

End Sub

Call MyASPJob()