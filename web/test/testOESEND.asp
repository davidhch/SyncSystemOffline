<%
'response.ContentType="text/xml"
strXML = "<SOAP:Body><Request><PaymentsID>5901</PaymentsID><OrderID>6944</OrderID><NewZappOrderID></NewZappOrderID><PaymentAmount>705</PaymentAmount><PaymentDate>30/10/2007</PaymentDate><CreditCardID>329</CreditCardID><PaymentMethodID>4</PaymentMethodID><TransactionCode>DN-694430102007161519</TransactionCode><ProcessingValid>true</ProcessingValid><ProcessingCode>A</ProcessingCode><ProcessingCodeMessage>Transaction authorised by bank.</ProcessingCodeMessage><AuthCode>006996</AuthCode><AuthMessage></AuthMessage><ResponseCode></ResponseCode><ResponseCodeMessage></ResponseCodeMessage><OnlineID></OnlineID><OfflineID>5901</OfflineID><TableID>3</TableID><Locked>1</Locked><Executed>0</Executed></Request></SOAP:Body></SOAP:Envelope>"

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

Const SoapServerURL = "http://www02.newzapp.co.uk/SOAP/AddPaymentSOAPOO.asp"

objxmlhttp.open "POST", SoapServerURL, False
objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
objxmlhttp.setRequestHeader "MessageType", "CALL"
objxmlhttp.setRequestHeader "Content-Type", "text/xml"
objxmlhttp.send(strXML)

Response.write objxmlhttp.Status & "<br>"
Response.write objxmlhttp.responseText & "<br>"

Set objxmlhttp = Nothing
Set objxmldom = Nothing

Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP") 

%>
