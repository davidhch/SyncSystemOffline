<%
'##################################
'
'	Class designed to read n nodes from an SOAP request
'
' ERRORS

'	101 == login failure
' 	201 == No Request

	Public strEnvelopePath
	Public strResponse
	Public objxmldom
	Public objxmlhttp
	Public objNodeList
	Public strReturnNodes
	Public bVerified
	
	Public strResult
	Public bResult
	Public nResult
	
	Public nRequestType
	Public bTest
	
	strEnvelopePath = "SOAP:Envelope/SOAP:Body/Request/" ' - we have chosen this by default
	strReturnNodes	= ""
	bVerified = False
	
	strResult 		=	"Nothing"	' Nothing done 
	bResult			=	False		' False by default
	nResult			=	0			' zero indicates nothing done
	nRequestType 	= 	0			' Nothing requested
	bTest			=	0			' Set to 1 to test in soap xml
	
Class CSoap

	Public Sub Init
		Call CreateObjects
		bVerified = True
		Call RequestXML
	End Sub
	
	Public Sub CreateObjects
		Set objxmldom = server.CreateObject("Microsoft.XMLDOM")
		Set objxmlhttp = server.CreateObject("Microsoft.XMLHTTP")	
	End Sub
	
	Public Sub RequestXML
		objxmldom.load( Request )
		Call NewReturnSoap
		'Call CheckLogin
	End Sub
	
	Public Sub SetNodeList (strNodeName) 
		Set objNodeList = objxmldom.getElementsByTagName(strNodeName)
	End Sub
	
	Public Sub AddNode (strNodeName, strNodeValue)
		strReturnNodes = strReturnNodes & "<"&strNodeName&">"&strNodeValue&"</"&strNodeName&">"
	End Sub
	
	Public Sub SetBool (strNodeValue)
		Call AddNode ("resultBool", strNodeValue)
	End Sub
	
	Public Sub SetText (strNodeValue)
		Call AddNode ("resultText", strNodeValue)
	End Sub
	
	Public Sub SetNumb (strNodeValue)
		Call AddNode ("resultNumb", strNodeValue)
	End Sub
	
	Public Sub SetUniq (strNodeValue)
		Call AddNode ("resultUniq", strNodeValue)
	End Sub
	
	Public Sub SetTest()
		bTest = 1 				' Sets bTest = 1 --  all Library Function Return only the response!
	End Sub
	
	Public Sub CleanUp
		Set objxmldom = Nothing
		Set objxmlhttp = Nothing
	End Sub
	
	'#################################
	
	Public Function GetNodeList(strNodeName)
		Call SetNodeList (strNodeName) 
		For Each Elem In objNodeList 
   			'response.write(Elem.tagName & "<br>") 
		Next
	End Function 
	
	Public Function ReadSingleNode (strNodeName)
		If CheckNodeExists (strNodeName) = True Then
			ReadSingleNode = objxmldom.selectSingleNode(strEnvelopePath & strNodeName).Text
		End If
	End Function 
	
	Public Function CheckNodeExists (strNodeName)
		If(TypeName(objxmldom.selectSingleNode(strEnvelopePath & strNodeName)) = "Nothing") then 
			CheckNodeExists = False
		Else
			CheckNodeExists = True
		End If
	End Function
	
	Public Function ProcessSoapRequest()
		Call CheckLogin
		If bVerified = True Then		
			nRequestType = objSoap.ReadSingleNode("RequestType")
			bTest = objSoap.ReadSingleNode("Test")
			
			Select Case cLng(nRequestType)
				Case 0
					' TEST
					strReturnNodes = ""
					Call SetResults ("THIS IS A TEST",True,100) 
				Case 1
					Call addSubscriber(bTest)
				Case 2
					Call addSubscriberToGroup(bTest)
				Case 3
					Call unsuscribeSubscriber(bTest)
				Case Else
					' Do Nothing
					strReturnNodes = ""
					Call SetResults ("RequestType was empty",False,201) 
			End Select
		End If
		
		ProcessSoapRequest = GetReturnSoap ()
		
	End Function
	
	Sub SetResults(str,b,n)
		Call SetTextResult(str)
		Call SetBooleanResult(b)
		Call SetNumericResult(n)
	End Sub
	
	Sub SetTextResult(str)
		strResult = str
	End Sub
	
	Sub SetNumericResult(n)
		nResult = n
	End Sub
	
	Sub SetBooleanResult(b)
		bResult = b
	End Sub
	
	Public Function GetReturnSoap ()

		strResponse = ""
		If bTest = 1 Then
		strResponse = strResponse & "<!-- TEST MODE -->"
		End If
		strResponse = strResponse & "<SOAP:Envelope xmlns:SOAP=""urn:schemas-xmlsoap-org:soap.v1"">"
		strResponse = strResponse & "<SOAP:Body>"
		strResponse = strResponse & "<Result>"
		
		strResponse = strResponse & "<ID>" & uniqueReqNo() & "</ID>"
		strResponse = strResponse & "<Date>" & now() & "</Date>"
		strResponse = strResponse & "<Text>" & strResult & "</Text>"
		strResponse = strResponse & "<Bool>" & bResult& "</Bool>"
		strResponse = strResponse & "<Num>"& nResult &"</Num>"
		
		strResponse = strResponse & "<Data>"& strReturnNodes &"</Data>"
		
		strResponse = strResponse & "</Result>"
		strResponse = strResponse & "</SOAP:Body>"
		strResponse = strResponse & "</SOAP:Envelope>"

		GetReturnSoap = strResponse		
		
		If bDebug = True Then
			response.ContentType = "text/xml"
			response.write "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>" & vbcrlf
			response.write GetReturnSoap
		End If

	End Function
	
	
	Public Function NewReturnSoap ()
		strReturnNodes = ""
	End Function
	

	
	Public Function uniqueReqNo()
		nNow = now()
		strSecond = second(nNow)
		strMinute = minute(nNow)
		strHour = hour(nNow)
		strDay = day(nNow)
		strMonth = month(nNow)
		strYear = year(nNow) - 138
		uniqueReqNo = session.SessionID & strYear & strSecond & strDay & strHour & strMinute & strMonth
	End Function
	
	Public Function CheckLogin
		' THIS FUNCTION VERIFIES LOG IN DETAILS
		' FIRST CHECK THE REQUIRED NODES ARE IN THE XML
		bVerified = False
		
		If CheckString(objSoap.ReadSingleNode("CID")) AND CheckString(objSoap.ReadSingleNode("UserName")) AND CheckString(objSoap.ReadSingleNode("Password")) Then
			CheckLogin = lib_LoginCheck(objSoap.ReadSingleNode("CID"), objSoap.ReadSingleNode("UserName"),objSoap.ReadSingleNode("Password"))
			bVerified = CheckLogin
		Else
			' FAILED DO NOT HAVE CORRECT NODES
			CheckLogin = False
			Call SetResults ("Login Failed",False,101) 
		End If	
	End Function
		
	Public Function CheckString(strToCheck)
		If Len(strToCheck)>0 Then
			CheckString = True
		Else
			CheckString =False
		End If
	End Function
	
	'***************************************************************
	'	SEND SOAP
	'
	Public Function SendSoap(strXML, SoapServerURL)
		Call CreateObjects
		
		objxmlhttp.open "POST", SoapServerURL, False
		objxmlhttp.setRequestHeader "Man", POST & " " & SoapServerURL & " HTTP/1.1"
		objxmlhttp.setRequestHeader "MessageType", "CALL"
		objxmlhttp.setRequestHeader "Content-Type", "text/xml"
		
		objxmlhttp.send(strXML)
		SendSoap = objxmlhttp.responseText
		
		Call CleanUp
	End Function
	
	'***************************************************************
	'	FUNCTIONS LIBRARY
	'
	Public Function addSubscriber(bTest)
		Call AddNode ("FirstName", "Darren")
		Call AddNode ("LastName", "Curnow")
		Call SetResults ("Request: Add Subscriber",True,100) 
	End Function
	
End Class
%>