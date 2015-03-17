<html>
<body>
<%
Dim objConn, objRS, strSQL
Dim x, curValue

strProvider = "Microsoft.Jet.OLEDB.4.0"
strExtendedPropeties = "Extended Properties=""Excel 8.0; IMEX=1;"""
strDataSource = "Data Source="&Server.MapPath("currentholders.xls")&"; "
strConnectionString = strDataSource & strExtendedPropeties

sql = "SELECT * FROM [Sheet1$]; "

Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Provider = strProvider
	objConn.ConnectionString = strConnectionString
	objConn.CursorLocation = 3
	objConn.Open
	
Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.Open sql, objConn, 3, 2
		
Set objRs = objConn.OpenSchema(20)
			nRecs=0
			maxLoops = 0
			Do Until objRs.EOF or maxLoops = 10

				tblName = objRs("TABLE_NAME")
				tblName = replace(tblName,"$","")
				tblName = replace(tblName,"'","")
				
				if instr(tblName,"MSys") then
				Else
					nRecs = nRecs + 1
					
					response.write "<br>" & tblName
					'strSheetOptions = strSheetOptions & "<option value=""" & tblName & """>" & tblName & "</option>"
					'strTBLName = tblName
				End If
			maxLoops = maxLoops + 1  
			objRs.MoveNext
			Loop

	Set objRs = nothing

Set objConn=Nothing
%>
</body>
</html>