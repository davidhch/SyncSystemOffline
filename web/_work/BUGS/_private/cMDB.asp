<%
Class CDBConn

		Private objConn
		Public m_strConn
		
		Private Sub Class_Initialize()	
			setDBPath "/_work/bugs/_private/bugs.mdb"
		End Sub
		
		Private Sub Class_Terminate()
			
		End Sub
		
		Public Sub openDBConnection()
			set objConn = Server.CreateObject("ADODB.Connection")
			objConn.Provider="Microsoft.Jet.OLEDB.4.0"
			objConn.Open server.mappath(m_strConn)
		End Sub
		
		Public Sub setDBPath(strPath)
			m_strConn = strPath
		End Sub
		
		Public Sub closeDBConnection()
			objConn.Close
			set objConn = Nothing
		End Sub
		
		Public Sub executeSQL(strSQL)
			objConn.Execute strSQL,recaffected
		End Sub
	
		Public Sub writeDBPath()
			response.write m_strConn
		End Sub
	
End Class	
%>