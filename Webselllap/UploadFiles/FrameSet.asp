<!-- MENU-LOCATION=NONE -->
<%@EnableSessionState=False%>
<html>
<head>
<title>Upload in Progress</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<frameset rows="60,*" frameborder="NO" border="0" framespacing="0">
  <frame src="Image.asp" name="topUploadFrame" scrolling="NO" noresize >
  <frame src="ProgressBar.asp?ProgressPage=<%=request.queryString("ProgressPage")%>&UploadID=<%=request.queryString("UploadID")%>" name="mainUploadFrame">
</frameset>
<noframes><body>

</body></noframes>
</html>
