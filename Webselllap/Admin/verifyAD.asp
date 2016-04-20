<%
	'Save the entered username and password
	Username = Request.Form("txtUsername")	
	Password = Request.Form("txtPassword")
	TenThat  = Request.Form("txtTenThat")
	
	'Build connection with database
	set conn = server.CreateObject ("ADODB.Connection")		
	conn.Open "DRIVER={SQL Server};SERVER=localhost;UID=sa;PWD=123456;DATABASE=CUA_HANG_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")		
	'Open record with entered username
	rs.Open "SELECT * FROM Admin where TaiKhoanAD='"& Username &"'", conn, 1 
	
	'If there is no record with the entered username, close connection
	'and go back to login with QueryString
	If rs.recordcount = 0 then
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		Response.Redirect("loginAD.asp?loginAD=namefailed")
	end if
	
	'If entered password is right, close connection and open mainpage
	if rs("MatKhau") = Password then
		Session("name") = rs("TenTK")
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("Products.asp")
	'If entered password is wrong, close connection 
	'and return to login with QueryString
	else
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("loginAD.asp?loginAD=passfailed")
	end if	

%>
