<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
	<%
		'Khai bao gio hang
		Dim p(100, 5)
		Dim dem
		dem = 0
		Dim tongtien
		tongtien = 0
		Dim masp
		Dim tensp
		Dim gia
		Dim hinhanh
		
		'Khoi tao gio hang
		masp = Request.Form("MaSPDatHang")
		tensp = Request.Form("TenSP")
		gia = Request.Form("Gia")
		hinhanh = Request.Form("HinhAnh")		
		
		if(Session("dem") = 0) then			
			p(dem, 0) = masp
			p(dem, 1) = tensp
			p(dem, 2) = hinhanh
			p(dem, 3) = 1
			p(dem, 4) = gia
			tongtien = p(dem, 4) * p(dem, 3)
			dem = dem + 1
			Session("giohang") = p
			Session("dem") = dem
			Session("tongtien") = tongtien
		else
			Dim q
			q = Session("giohang")
			dem = Session("dem")			
			Dim flag
			flag = 0
			for i = 0 to dem - 1
				if(masp = q(i, 0)) then
					flag = 1
					q(i, 3) = q(i, 3) + 1
					dem = dem - 1
				end if
			next
			if(flag = 0) then
				q(dem, 0) = masp
				q(dem, 1) = tensp
				q(dem, 2) = hinhanh
				q(dem, 3) = 1
				q(dem, 4) = gia
			end if
			for i = 0 to dem
				tongtien = tongtien + (q(i, 3) * q(i, 4))
			next
			dem = dem + 1
			Session("giohang") = q
			Session("dem") = dem
			Session("tongtien") = tongtien
		end if 
		Response.Redirect("HienThi.asp")
	%>
</body>
</html>
