<%
response.ContentType="text/html"

'response.write "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
response.write "<ui>"
Dim objConn, objRS, strSQL
Dim x, curValue
Dim dbConn

Dim nID
Dim nParentID 		
Dim strTYPE			
Dim nSectionID		
Dim strSection 		
Dim strScreenType 	
Dim strName 		
Dim strAction		
Dim strEvent		
Dim bMultiples		
Dim strDisabled		
Dim strFeedback		
Dim strFlex			
Dim strLite			
Dim strPro			
Dim strEnt	

Dim strDotNoted	
Dim arrDotNoted		

dbConn = "DRIVER={Microsoft Excel Driver (*.xls)}; IMEX=1;  Excel 8.0; DBQ=" & Server.MapPath("currentholders.xls") & "; "
'strSQL = "SELECT * FROM [Sheet1$] WHERE [SECTION] IS NOT NULL AND [SECTION ID]=2 ORDER BY [SECTION ID],[SECTION], [ID] "'WHERE [SECTION ID]=1"

strSQL = "SELECT * FROM [Sheet1$] WHERE [SECTION] ='SUBSCRIBERS. mainframe. new subscriber' AND [SECTION ID]=2 ORDER BY [SECTION ID],[SECTION], [ID] "

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open dbConn
Set objRS=objConn.Execute(strSQL)

nPreviousSection = 0
nSection = 0
strPreviousSection = ""

nPreviousArr1 = 0
nPreviousArr2 = 0
nPreviousArr3 = 0
nPreviousArr4 = 0
nPreviousArr5 = 0

strPreviousArr1 = ""
strPreviousArr2 = ""
strPreviousArr3 = ""
strPreviousArr4 = ""
strPreviousArr5 = ""

strPreviousSectionPos = 0

Do Until objRS.EOF
		
		nID 			= objRS("ID")
		nParentID 		= objRS("ParentID")
		strType			= objRS("TYPE")
		nSectionID		= objRS("SECTION ID")
		strSection 		= objRS("SECTION")
		strScreenType 	= objRS("SCREEN TYPE")
		strName 		= objRS("HOLDER (current name)")
		strAction		= objRS("CURRENT ACTION")
		strEvent		= objRS("On type")
		bMultiples		= objRS("Multiple Instances")
		strDisabled		= objRS("Disable/Enable")
		strFeedback		= objRS("feedback")
		strFlex			= objRS("FLEX")
		strLite			= objRS("LITE")
		strPro			= objRS("PRO")
		strEnt			= objRS("ENT")
		
		If Len(strEvent)>0 Then
		Else
			strEvent = "none"
		End if
		
		If Len(strType)>0 Then
		Else
			strType = "BUTTON"
		End if
		
		If Len(bMultiples)>0 Then
			bMultiples = "multiple"
		Else
			bMultiples = ""
		End if
		
		If Len(strFeedback)>0 Then
			'strFeedback = "true"
		Else
			strFeedback = "none"
		End if
		
		If Len(strDisabled)>0 Then
		Else
			strDisabled = "dissabled"
		End if
		
		If Len(strFlex) >0 Then
			strFlex = "0"
		Else
			strFlex = "1"
		End If
		
		If Len(strLite) >0 Then
			strLite = "0"
		Else
			strLite = "1"
		End If
		
		If Len(strPro) >0 Then
			strPro = "0"
		Else
			strPro = "1"
		End If
		
		If Len(strEnt) >0 Then
			strEnt = "0"
		Else
			strEnt = "1"
		End If
		
		strDotNoted		= cStr(strSection)
		
		strDotNoted	= replace(strDotNoted," ","")
		strDotNoted = lcase(strDotNoted)

		' NOW SPLIT BY THE DOTS AND REVIEW THE RESULT
		arrDotNoted	= Split(replace(strDotNoted,".",","),",")
		
		' number of items in the list
		
		'strSection
		
		If strPreviousSection<>arrDotNoted(0) Then
			
			nSection = nSection + 1
			strPreviousSection = arrDotNoted(0)
			
			''Response.write "<h1 style=""font-size:48px;"">"&nSection&": " & uCase(arrDotNoted(0)) & "</h1>"
			response.write lcase(node_0_close)
			'BUILD THE TOP LEVEL NODE
			'********************************
			node_0_open 	= "<" & uCase(arrDotNoted(0)) & ">"
			node_0_close 	= "</" & uCase(arrDotNoted(0)) & ">"
			
			node_1_open 	= ""
			node_1_close 	= ""
			
			node_2_open 	= ""
			node_2_close 	= ""
			
			node_3_open 	= ""
			node_3_close 	= ""
			
			node_4_open 	= ""
			node_4_close 	= ""
			
			node_5_open 	= ""
			node_5_close 	= ""
			
			node_6_open 	= ""
			node_6_close 	= ""
			'*******************************
			
			nPreviousArr1 = 0
			nPreviousArr2 = 0
			nPreviousArr3 = 0
			nPreviousArr4 = 0
			nPreviousArr5 = 0
			
			response.write lcase(node_0_open)
		End If
		
		
		
		If uBound(arrDotNoted) => 1 Then
		
			strSectionPos = nSection
			
			If strPreviousArr1<>arrDotNoted(1) Then
				'Response.write "<br>H2 SECTION PAGE FOR " & arrDotNoted(1)
				strPreviousArr1 = arrDotNoted(1)
				nPreviousArr1 = nPreviousArr1 + 1
				
				nPreviousArr2 = 0
				nPreviousArr3 = 0
				nPreviousArr4 = 0
				nPreviousArr5 = 0
				
				''Response.write "<h2 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & ": " & arrDotNoted(0)&"."&arrDotNoted(1) & "</h2>"
				
				response.write lcase(node_1_close)
				'BUILD THE 2nd LEVEL NODE
				'********************************
				node_1_open 	= "<" & uCase(arrDotNoted(1)) & ">"
				node_1_close 	= "</" & uCase(arrDotNoted(1)) & ">"
				
				node_2_open 	= ""
				node_2_close 	= ""
				
				node_3_open 	= ""
				node_3_close 	= ""
				
				node_4_open 	= ""
				node_4_close 	= ""
				
				node_5_open 	= ""
				node_5_close 	= ""
				
				node_6_open 	= ""
				node_6_close 	= ""
				'*******************************
				response.write lcase(node_1_open)
				
			End If
			
			strSectionPos = strSectionPos & "." & nPreviousArr1
		End If
		
		If uBound(arrDotNoted) => 2 Then
			
			strSectionPos = nSection
			
			If strPreviousArr1<>arrDotNoted(1) Then
				strPreviousArr1 = arrDotNoted(1)
				nPreviousArr1 = nPreviousArr1 + 1
			End If
			
			If strPreviousArr2<>arrDotNoted(2) Then
				
				strPreviousArr2 = arrDotNoted(2)
				nPreviousArr2 = nPreviousArr2 + 1
				
				nPreviousArr3 = 0
				nPreviousArr4 = 0
				nPreviousArr5 = 0
				
				''Response.write "<h3 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2&": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2) & "</h3>"
				
				response.write lcase(node_2_close)
				
				'BUILD THE 3rdLEVEL NODE
				'********************************
				node_2_open 	= "<" & uCase(arrDotNoted(2)) & ">"
				node_2_close 	= "</" & uCase(arrDotNoted(2)) & ">"
				
				node_3_open 	= ""
				node_3_close 	= ""
				
				node_4_open 	= ""
				node_4_close 	= ""
				
				node_5_open 	= ""
				node_5_close 	= ""
				
				node_6_open 	= ""
				node_6_close 	= ""
				'*******************************
				
				response.write lcase(node_2_open)
			End If
			
			strSectionPos = strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2
		End If
		
		If uBound(arrDotNoted) => 3 Then
			strSectionPos = nSection
			
			If strPreviousArr1<>arrDotNoted(1) Then
				strPreviousArr1 = arrDotNoted(1)
				nPreviousArr1 = nPreviousArr1 + 1
			End If
			
			If strPreviousArr2<>arrDotNoted(2) Then
				strPreviousArr2 = arrDotNoted(2)
				nPreviousArr2 = nPreviousArr2 + 1
			End If
			
			If strPreviousArr3<>arrDotNoted(3) Then
				'Response.write "<br>H4 SECTION PAGE FOR " & arrDotNoted(3)
				strPreviousArr3 = arrDotNoted(3)
				nPreviousArr3 = nPreviousArr3 + 1
				
				nPreviousArr4 = 0
				nPreviousArr5 = 0
				''Response.write "<h4 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2& "." & nPreviousArr3 &": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2)&"."&arrDotNoted(3) & "</h4>"
				
				response.write lcase(node_3_close)
				'BUILD THE 4th LEVEL NODE
				'********************************
				node_3_open 	= "<" & uCase(arrDotNoted(3)) & ">"
				node_3_close 	= "</" & uCase(arrDotNoted(3)) & ">"
				
				node_4_open 	= ""
				node_4_close 	= ""
				
				node_5_open 	= ""
				node_5_close 	= ""
				
				node_6_open 	= ""
				node_6_close 	= ""
				'*******************************
				response.write lcase(node_3_open)
			End If
			
			strSectionPos = strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2 & "." & nPreviousArr3
		End If
		
		If uBound(arrDotNoted) => 4 Then
			strSectionPos = nSection
			
			If strPreviousArr1<>arrDotNoted(1) Then
				strPreviousArr1 = arrDotNoted(1)
				nPreviousArr1 = nPreviousArr1 + 1
			End If
			
			If strPreviousArr2<>arrDotNoted(2) Then
				strPreviousArr2 = arrDotNoted(2)
				nPreviousArr2 = nPreviousArr2 + 1
			End If
			
			If strPreviousArr3<>arrDotNoted(3) Then
				strPreviousArr3 = arrDotNoted(3)
				nPreviousArr3 = nPreviousArr3 + 1
			End If
			
			If strPreviousArr4<>arrDotNoted(4) Then
				'Response.write "<br>H5 SECTION PAGE FOR " & arrDotNoted(4)
				
				
				
				strPreviousArr4 = arrDotNoted(4)
				nPreviousArr4 = nPreviousArr4 + 1
				''Response.write "<h5 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2& "." & nPreviousArr3 & "." & nPreviousArr4 & ": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2)&"."&arrDotNoted(3)&"."&arrDotNoted(4) & "</h5>"
				
				response.write lcase(node_4_close)
				
				'BUILD THE 5th LEVEL NODE
				'********************************
				node_4_open 	= "<" & uCase(arrDotNoted(4)) & ">"
				node_4_close 	= "</" & uCase(arrDotNoted(4)) & ">"
				
				node_5_open 	= ""
				node_5_close 	= ""
				
				node_6_open 	= ""
				node_6_close 	= ""
				'*******************************
				
				response.write lcase(node_4_open)
				
			End If
			
			strSectionPos = strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2 & "." & nPreviousArr3 & "." & nPreviousArr4
		End If
		
		If strSectionPos<>strPreviousSectionPos Then
			' new item set to 1
			nItem = 1
		Else
			' increment the item number
			nItem = nItem + 1
		End If
		strPreviousSectionPos = strSectionPos
		
		''Response.Write "<h6 style=""font-size:24px;"">"&strSectionPos&"-"&nItem&": " & lcase(strName) &"</h6>"
		nXMLID = strSectionPos&"-"&nItem
		nXMLID = replace(nXMLID,".","-")
		
		nUIID = strDotNoted&"."&lcase(replace(strName," ",""))
		nUIID = replace(nUIID,".","-")
		
		strXMLNode = replace(strName," ","-")
		
		Response.Write "<"&lcase(strXMLNode)&">"
		
			'Response.Write "<dbid>"&nXMLID&"</dbid>" ' this will be DB ID once in the database
			if 1=2 then
			Response.Write "<holderName>"&strDotNoted&"."&lcase(replace(strName," ",""))&"</holderName>"
			Response.Write "<section>"&strDotNoted&"</section>"
			Response.Write "<parent>"&getParent(nParentID)&"</parent>"
			Response.Write "<screentype>"&lcase(strScreenType)&"</screentype>"
			end if
			Response.Write "<id>"&nUIID&"</id>"
			Response.Write "<type>"&lcase(strType)&"</type>"
			Response.Write "<disabled>"&strDisabled&"</disabled>"
			Response.Write "<multiple>"&bMultiples&"</multiple>"
			Response.Write "<title>"&strName&"</title>"
			Response.Write "<class></class>"
			
			Response.Write "<licence>"
				Response.Write "<flex>"&strFlex&"</flex>"
				Response.Write "<lite>"&strLite&"</lite>"
				Response.Write "<pro>"&strPro&"</pro>"
				Response.Write "<ent>"&strEnt&"</ent>"
			Response.Write "</licence>"
			
			
			
			if 1=2 then
			Response.Write "<jsActions>"
				Response.Write "<onblur></onblur>"
				Response.Write "<onchange></onchange>"
				Response.Write "<onclick></onclick>"
				Response.Write "<ondblclick></ondblclick>"
				Response.Write "<onerror></onerror>"
				Response.Write "<onfocus></onfocus>"
				Response.Write "<onkeydown></onkeydown>"
				Response.Write "<onkeypress></onkeypress>"
				Response.Write "<onkeyup></onkeyup>"
				Response.Write "<onload></onload>"
				Response.Write "<onmousedown></onmousedown>"
				Response.Write "<onmousemove></onmousemove>"
				Response.Write "<onmouseout></onmouseout>"
				Response.Write "<onmouseover></onmouseover>"
				Response.Write "<onmouseup></onmouseup>"
				Response.Write "<onreset></onreset>"
				Response.Write "<onresize></onresize>"
				Response.Write "<onselect></onselect>"
				Response.Write "<onsubmit></onsubmit>"
				Response.Write "<onunload></onunload>"
				Response.Write "<onbeforeunload></onbeforeunload>"
			Response.Write "</jsActions>"
			end if
		Response.Write "</"&lcase(strXMLNode)&">"
		
	objRS.MoveNext
Loop

objRS.Close
objConn.Close

Set objRS=Nothing
Set objConn=Nothing


Function getParent(str)
	' access parent here
	bError = 0
	If str="p0" Then
		'bError = 1
	End If


	If Len(str) > 0 Then	
	
	str = replace(str,"p","")
	arrSplitNo = split(str,"$$$$")
	strWHERE = ""
	
	for loopctr = 0 to ubound(arrSplitNo)
	
		If loopctr > 0 Then
			strWHERE = strWHERE & "OR "	
		End If
		
		If Len(arrSplitNo(loopctr))>0 Then
		strWHERE = strWHERE & "[ID] = " & arrSplitNo(loopctr)
		Else
			bError = 1
		End if
		
	next
	
		If bError = 0 Then
			dbConn2 = "DRIVER={Microsoft Excel Driver (*.xls)}; IMEX=1;  Excel 8.0; DBQ=" & Server.MapPath("Copy of currentHolders.xls") & "; "
			strSQL2 = "SELECT * FROM [Sheet1$] WHERE "&strWHERE&""'WHERE [SECTION ID]=1"
			
			Set objConn2 = Server.CreateObject("ADODB.Connection")
			objConn2.Open dbConn2
			Set objRS2=objConn2.Execute(strSQL2)
			
			n=0
			Do Until objRS2.EOF
					
					strSectionParent = objRS2("SECTION") & objRS2("HOLDER (current name)")
					strSectionParent = replace(strSectionParent," ","")
					strSectionParent = lcase(strSectionParent)
					
					If n=0 then
						strSection2 		= strSectionParent
					Else
						strSection2 		= strSection2 &","&strSectionParent
					End If
					
					n=n+1
					objRS2.MoveNext
			Loop
			
			objRS2.Close
			objConn2.Close
			
			Set objRS2=Nothing
			Set objConn2=Nothing
			
			If len(strSection2) > 0 then
				getParent = strSection2
			Else
				getParent = "None"
			End if
		End if
	End if
	
	If len(len(strSection2)) > 0 then
				
	Else
		getParent = "none"
	End if
End Function 


response.write "</ui>"
%>