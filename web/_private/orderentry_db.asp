<%
	bLive = False
	bDev = True
	
	If bLive Then	
		strOEDB = "Provider=SQLOLEDB;PWD=5nowPatrol;UID=OrderEntry_IIS;Database=OrderEntry;Server=DESTINETDC02"
	Else
		strOEDB = "Provider=SQLOLEDB;PWD=5nowPatrol;UID=OrderEntry_IIS;Database=OrderEntry;Server=DESTINETDC02"
	End If

	'If bDev Then	
	'	strOEDB = "Provider=SQLOLEDB;PWD=newzappsoap;UID=OrderEntry-Dev;Database=OrderEntry-Development;Server=DESTINETDC02"
	'Else
	'	strOEDB = "Provider=SQLOLEDB;PWD=newzappsoap;UID=OrderEntry-Dev;Database=OrderEntry-Development;Server=DESTINETDC02"
	'End If
%>
