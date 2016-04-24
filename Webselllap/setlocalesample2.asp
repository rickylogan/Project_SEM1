<%@ LANGUAGE=VBScript LCID=1033 CODEPAGE=65001 ENABLESESSIONSTATE=False%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<%
Sub DisplayCurValue(stDescription, lcid, fRtl)
	Dim lcidOld

	lcidOld = GetLocale()
On Error Resume Next
	Call SetLocale(lcid)
	if Err.Number <> 0 Then
		Response.Write "<br>" & stDescription & ":  <font color=red>Sorry, the server does not support this locale</font>" & vbCrLf
	Else
		If fRtl then response.write "<span class=ctlRTL>&rlm;"
		Response.Write "<br>" & stDescription & ": <b>" & FormatCurrency(1234) & "</b>" & vbCrLf
		If fRtl then response.write "</span>"
	End If
	Call SetLocale(lcidOld)
end Sub
%>
</style>
</head>
<body>
<% 
	Call DisplayCurValue("Vietnamese (1066)", 1066, False)
	Call DisplayCurValue("", 1078, False)
	Call DisplayCurValue("Faeroese (1080)", 1080, False)
%>
</body>
</html>