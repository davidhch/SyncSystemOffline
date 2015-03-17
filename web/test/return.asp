
<!--#include file="CSoap.asp"-->
<%
	Dim objSoap
 	Set objSoap = New CSoap
 	objSoap.Init
	objSoap.NewReturnSoap
nSRID 			= objSoap.ReadSingleNode ("SRID")

Select Case nSRID
			

Case 101
			sendReq = ""
			sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			sendReq = sendReq & "<SOAP:Body>"
			sendReq = sendReq & "<Request>"
			
			sendReq = sendReq & "<ReturnData>This was test call coz you sent "&nSRID&"</ReturnData>"
			
			sendReq = sendReq & "</Request>"
			sendReq = sendReq & "</SOAP:Body>"
			sendReq = sendReq & "</SOAP:Envelope>" 
Case 102
			sendReq = ""
			sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			sendReq = sendReq & "<SOAP:Body>"
			sendReq = sendReq & "<Request>"
			
			sendReq = sendReq & "<ReturnData>You want to update the expiry date you sent "&nSRID&"</ReturnData>"
			
			sendReq = sendReq & "</Request>"
			sendReq = sendReq & "</SOAP:Body>"
			sendReq = sendReq & "</SOAP:Envelope>" 
Case 103
			sendReq = ""
			sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			sendReq = sendReq & "<SOAP:Body>"
			sendReq = sendReq & "<Request>"
			
			sendReq = sendReq & "<ReturnData>You are adding a pack to this account you sent "&nSRID&"</ReturnData>"
			
			sendReq = sendReq & "</Request>"
			sendReq = sendReq & "</SOAP:Body>"
			sendReq = sendReq & "</SOAP:Envelope>" 
Case Else
			sendReq = ""
			sendReq = sendReq & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
			sendReq = sendReq & "<SOAP:Body>"
			sendReq = sendReq & "<Request>"
			
			sendReq = sendReq & "<ReturnData>Your soap is bogus, you sent "&nSRID&"</ReturnData>"
			
			sendReq = sendReq & "</Request>"
			sendReq = sendReq & "</SOAP:Body>"
			sendReq = sendReq & "</SOAP:Envelope>" 
End Select

response.write sendReq


%>