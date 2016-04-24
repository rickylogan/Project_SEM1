<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Connection.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

</head>
<body>
	<%
	Dim ngaylap
	ngaylap = date	
	Dim maddh
	Dim dem
	Dim kq
	'Tao ket noi CSDL
	Dim editCmd
	set editCmd = Server.CreateObject("ADODB.Command")
	editCmd.ActiveConnection = MM_Connection_STRING
	
	if(Session("TKKH") = "") then		
		Response.Write("KH Moi")
	'Tao hoa don		
		Dim idkh
		editCmd.commandText = "select max(IDKH) as 'idkh' from KHMoi"
		Dim kqId
		set kqId = editCmd.execute
		idkh = kqId.Fields.Item("idkh").value
				
		editCmd.CommandText = "insert into DonDatHangKHMoi values('"&idkh&"',"&Session("tongtien")&",'"&ngaylap&"')"
		editCmd.Prepared = true
		editCmd.execute
		
		'Lay maddh vua them		
		editCmd.commandText = "select max(MaDDH) as 'maddh' from DonDatHangKHMoi"
		set kq = editCmd.execute
		maddh = kq.Fields.Item("maddh").value	
	
		'Them chi tiet hoa don
		dem = Session("dem")
		q = Session("giohang")
		for i = 0 to dem - 1			
			editCmd.CommandText = "insert into CTDDHKHMoi values("&maddh&", "&q(i,0)&","&q(i,3)&", "&q(i,4)&")"
			editCmd.execute
		next
		Session("dem") = 0
		Session("giohang") = 0
		Session("tongtien") = 0
	else
		Response.Write("KH Đã có TK")
		'Tao hoa don	
		editCmd.CommandText = "insert into DonDatHang values('"&Session("name")&"',"&Session("tongtien")&",'"&ngaylap&"')"
		editCmd.Prepared = true
		editCmd.execute
		
		'Lay maddh vua them
		editCmd.commandText = "select max(MaDDH) as 'maddh' from DonDatHang"
		
		set kq = editCmd.execute
		maddh = kq.Fields.Item("maddh").value		
		'Them chi tiet hoa don
		dem = Session("dem")
		q = Session("giohang")
		for i = 0 to dem - 1
			editCmd.CommandText = "insert into CTDDH values("&maddh&", "&q(i,0)&","&q(i,3)&", "&q(i,4)&")"
			editCmd.execute
		next
		Session("dem") = 0
		Session("giohang") = 0
		Session("tongtien") = 0
	end if
	editCmd.Activeconnection.Close()
	Response.Redirect("index.asp")	
	%>
    
</body>
</html>
