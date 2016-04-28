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
		Dim tongtien
		tongtien = Session("tongtien")
		Dim n
		n = Request.QueryString
		
		tongtien = tongtien - (q(n, 3) * q(n, 4))
		
		for k = 0 to n - 1
			q(k, 0) = q(k, 0)
			q(k, 1) = q(k, 1)
			q(k, 2) = q(k, 2)
			q(k, 3) = q(k, 3)
			q(k, 4) = q(k, 4)			
		next
		
		for k = n to dem - 1
			q(k, 0) = q(k + 1, 0)
			q(k, 1) = q(k + 1, 1)
			q(k, 2) = q(k + 1, 2)
			q(k, 3) = q(k + 1, 3)
			q(k, 4) = q(k + 1, 4)	
		next 
		
		
		
		dem = dem - 1
		Session("dem") = dem
		Session("giohang") = q
		Session("tongtien") = tongtien
		
		Response.Redirect("HienThi.asp")		
	%>
</body>
</html>
