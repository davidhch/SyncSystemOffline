<html>
<head>
<style>
td{border-bottom:1px solid #000000; font-family:arial; font-size:10px;}
</style>
</head>
<body>
<%
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
strSQL = "SELECT * FROM [Sheet1$] WHERE [SECTION] IS NOT NULL AND [SECTION ID] = 8 ORDER BY [SECTION ID],[SECTION], [ID]"'WHERE [SECTION ID]=1"

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
nCountSection = 0
Do Until objRS.EOF

		'objRS.Fields(x).Name
		'objRS.Fields(x).Value
		nCountSection = nCountSection + 1
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
			bMultiples = "YES"
		Else
			bMultiples = "NO"
		End if
		
		If Len(strFeedback)>0 Then
		Else
			strFeedback = "none"
		End if
		
		If Len(strDisabled)>0 Then
		Else
			strDisabled = "NO"
		End if
		
		If Len(strFlex) >0 Then
		Else
			strFlex = "YES"
		End If
		
		If Len(strLite) >0 Then
		Else
			strLite = "YES"
		End If
		
		If Len(strPro) >0 Then
		Else
			strPro = "YES"
		End If
		
		If Len(strEnt) >0 Then
		Else
			strEnt = "YES"
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
			
			Response.write "<h1 style=""font-size:48px;"">"&nSection&": " & uCase(arrDotNoted(0)) & "</h1>"
			Response.Write "<br style='page-break-before:always'>"
			nPreviousArr1 = 0
			nPreviousArr2 = 0
			nPreviousArr3 = 0
			nPreviousArr4 = 0
			nPreviousArr5 = 0
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
				
				Response.write "<h2 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & ": " & arrDotNoted(0)&"."&arrDotNoted(1) & "</h2>"
				Response.Write "<br style='page-break-before:always'>"
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
				
				Response.write "<h3 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2&": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2) & "</h3>"
				Response.Write "<br style='page-break-before:always'>"
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
				Response.write "<h4 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2& "." & nPreviousArr3 &": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2)&"."&arrDotNoted(3) & "</h4>"
				Response.Write "<br style='page-break-before:always'>"
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
				Response.write "<h5 style=""font-size:32px;"">"&strSectionPos & "." & nPreviousArr1 & "." & nPreviousArr2& "." & nPreviousArr3 & "." & nPreviousArr4 & ": " & arrDotNoted(0)&"."&arrDotNoted(1)&"."&arrDotNoted(2)&"."&arrDotNoted(3)&"."&arrDotNoted(4) & "</h5>"
				Response.Write "<br style='page-break-before:always'>"
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
		
		'Response.Write "<h6>"&strSectionPos&"."&nItem&": "&strDotNoted&"."&lcase(replace(strName," ",""))&" ---- " & strName &"</h6>"
		Response.Write "<h6 style=""font-size:24px;"">"&strSectionPos&"-"&nItem&": " & lcase(strName) &"</h6>"
		
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"">"
		
		xxxx = "&nbsp;"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top"" style=""width:200px;""><strong>HOLDER NAME</strong></td>"
			Response.Write "<td valign=""top"">"&strDotNoted&"."&lcase(replace(strName," ",""))&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>TYPE</strong></td>"
			Response.Write "<td valign=""top"">"&lCase(strType)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>SECTION</strong></td>"
			Response.Write "<td valign=""top"">"&strDotNoted&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>PARENT</strong></td>"
			Response.Write "<td valign=""top"">"&getParent(nParentID)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>SCREEN TYPE</strong></td>"
			Response.Write "<td valign=""top"">"&lcase(strScreenType)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>CURRENT ACTION</strong></td>"
			Response.Write "<td valign=""top"">"&strAction&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>EVENT TYPE</strong></td>"
			Response.Write "<td valign=""top"">"&lcase(strEvent)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>FEEDBACK</strong></td>"
			Response.Write "<td valign=""top"">"&strFeedback&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>MULTIPLE INSTANCES</strong></td>"
			Response.Write "<td valign=""top"">"&bMultiples&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>DISABLED</strong></td>"
			Response.Write "<td valign=""top"">"&strDisabled&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>FLEX</strong></td>"
			Response.Write "<td valign=""top"">"&strFlex&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>LITE</strong></td>"
			Response.Write "<td valign=""top"">"&strLite&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>PRO</strong></td>"
			Response.Write "<td valign=""top"">"&strPro&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>ENT</strong></td>"
			Response.Write "<td valign=""top"">"&strEnt&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>NAME</strong></td>"
			Response.Write "<td valign=""top"">"&lcase(strName)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top""><strong>ORIGINAL NAME</strong></td>"
			Response.Write "<td valign=""top"">"&lcase(strName)&"</td>"
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td valign=""top"" style=""height:300px; border:0px;""><strong>NOTES</strong></td>"
			Response.Write "<td valign=""top"" style=""height:300px; border:0px;"">"&xxxx&"</td>"
		Response.Write "</tr>"
		
		Response.Write "</table>" 
		
		Response.Write "<br style='page-break-before:always'>"
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
						strSection2 		= strSection2 &"<br>"&strSectionParent
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

response.write "************************* " & nCountSection
%>
</body>
</html>