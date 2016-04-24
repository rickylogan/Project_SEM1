<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim DDH__MMColParam
DDH__MMColParam = ""
If (Request.Form("MaDDH") <> "") Then 
  DDH__MMColParam = Request.Form("MaDDH")
  Session("MaDDH") = DDH__MMColParam
End If
%>
<%
Dim DDH
Dim DDH_cmd
Dim DDH_numRows

Set DDH_cmd = Server.CreateObject ("ADODB.Command")
DDH_cmd.ActiveConnection = MM_Connection_STRING
DDH_cmd.CommandText = "SELECT MaDDH FROM dbo.DonDatHang WHERE MaDDH = ?" 
DDH_cmd.Prepared = true
DDH_cmd.Parameters.Append DDH_cmd.CreateParameter("param1", 5, 1, -1, DDH__MMColParam) ' adDouble

Set DDH = DDH_cmd.Execute
DDH_numRows = 0
%>
<%
	Password = Request.Form("txtPassword")
	set conn = server.CreateObject ("ADODB.Connection")		
	conn.Open "DRIVER={SQL Server};SERVER=localhost;UID=sa;PWD=123456;DATABASE=CUA_HANG_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")		
	rs.Open "SELECT * FROM Admin where TaiKhoanAD='ADMINISTRATOR'", conn, 1 
	if rs("MatKhau") = Password then
		Session("confirm") = "confirmed"
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("EditThanhToan.asp")
	else
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("DDH.asp")
	end if	

%>
<%
DDH.Close()
Set DDH = Nothing
%>
