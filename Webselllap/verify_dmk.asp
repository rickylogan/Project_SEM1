<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Username = session("TKKH")
	Password = Request.Form("txtOP")
	set conn = server.CreateObject ("ADODB.Connection")		
	conn.Open "DRIVER={SQL Server};SERVER=localhost;UID=sa;PWD=123456;DATABASE=CUA_HANG_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")		
	rs.Open "SELECT * FROM KhachHang where TKKH='"& Username &"'", conn, 1
	if rs("MatKhau") = Password then
	rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
	Session("ConfirmMK")="Confirmed"
	Response.Redirect("SuaTTCN.asp?TTCN=ChangeNP")
	else
	Session("noti")="*Sai mật khẩu"
	rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
	Response.Redirect("SuaTTCN.asp?TTCN=ChangeOP")
	end if
%>