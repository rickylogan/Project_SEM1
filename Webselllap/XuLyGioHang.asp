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
	'Tao hoa don		
		Dim ngaylap
		ngaylap = date
	'Tao ket noi CSDL
		Dim editCmd
		set editCmd = Server.CreateObject("ADODB.Command")
		editCmd.ActiveConnection = MM_Connection_STRING
		editCmd.CommandText = "insert into DonDatHang values('"&Session("name")&"',"&Session("tongtien")&",'"&ngaylap&"')"
		editCmd.Prepared = true
		editCmd.execute
		
	'Lay maddh vua them
		Dim maddh
		editCmd.commandText = "select max(MaDDH) as 'maddh' from DonDatHang"
		Dim kq
		set kq = editCmd.execute
		maddh = kq.Fields.Item("maddh").value		
	'Them chi tiet hoa don
		Dim dem
		dem = Session("dem")
		Dim q
		q = Session("giohang")
		for i = 0 to dem - 1
			
			editCmd.CommandText = "insert into CTDDH values("&maddh&", "&q(i,0)&","&q(i,3)&", "&q(i,4)&")"
			editCmd.execute
		next
		Session("dem") = 0
		Session("giohang") = 0
		Session("tongtien") = 0
		editCmd.Activeconnection.Close()
		Response.Redirect("index.asp")	
	%>
    
</body>
</html>
