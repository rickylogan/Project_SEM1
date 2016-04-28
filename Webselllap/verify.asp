<%
	Username = Request.Form("txtUsername")	
	Password = Request.Form("txtPassword")
	set conn = server.CreateObject ("ADODB.Connection")		
	conn.Open "DRIVER={SQL Server};SERVER=localhost;UID=sa;PWD=123456;DATABASE=CUA_HANG_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")		
	rs.Open "SELECT TKKH, MatKhau, TenKH FROM KhachHang where TKKH='"& Username &"'", conn, 1 
	
	If rs.recordcount = 0 then
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		Response.Redirect("login.asp?login=namefailed")
	end if
	
	'If entered password is right, close connection and open mainpage
	if rs("MatKhau") = Password then
		Session("TKKH") = rs("TKKH")
		Session("name") = rs("TenKH")
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("index.asp")
	'If entered password is wrong, close connection 
	'and return to login with QueryString
	else
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("login.asp?login=passfailed")
	end if	

%>
