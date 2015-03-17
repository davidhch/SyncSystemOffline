<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%

	For i = 0 to 5000
		%>
		if($("mySelect<%response.write i%>").nzElementExists()){
		alert("mySelect<%response.write i%> does exist");
		$("mySelect<%response.write i%>").on('change',function(){$("mySelect<%response.write i%>").nzPopulateSelect2("mySelect<%response.write i%>","onchange");});
		$("mySelect<%response.write i%>").on('focus',function(){$("mySelect<%response.write i%>").nzFocusSelectList("mySelect<%response.write i%>","focus");});
		$("mySelect<%response.write i%>").nzLoadSelect2Data("mySelect<%response.write i%>","load");
	}
		<%
	Next

%>