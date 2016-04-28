<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
	'Tao ket noi CSDL
	Dim MaDDH
	Dim SoLuong
	Dim MaSP
		MaDDH = Request.Form("MaDDH")
		SoLuong = Request.Form("SoLuong")
		MaSP = Request.Form("MaSP")
	Dim editCmd
	set editCmd = Server.CreateObject("ADODB.Command")
	editCmd.ActiveConnection = MM_Connection_STRING
	editCmd.commandText = "DELETE FROM dbo.DonDatHang WHERE MaDDH=" & MaDDH & "DELETE FROM dbo.CTDDH WHERE MaDDH=" & MaDDH &  "UPDATE SanPham SET SoLuong = SoLuong + " & SoLuong & " WHERE MaSP = " & MaSP
	editCmd.Prepared = true
	editCmd.execute
	Response.Redirect("DDH.asp")
%>