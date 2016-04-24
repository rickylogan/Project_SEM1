<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
	<%
		Dim dem
		dem = Session("dem")
		Dim q
		q = Session("giohang")
		Dim n
		n = Request.QueryString
		Response.Write(q(n,1))
				
		dem = dem
		Session("dem") = dem
		
		Response.Redirect("HienThi.asp")		
	%>
</body>
</html>
