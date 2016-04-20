<%
	'Save entered username and password
	Username = Request.Form("txtUsername")
	Password = Request.Form("txtPassword")
		
	'Check if username and password are entered
	if Username = "" then Response.redirect("loginAD.asp?loginAD=createnew")
	if Password = "" then Response.Redirect("loginAD.asp?loginAD=createpassfailed")
	
	'Build connection
	set conn = server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={SQL Server};SERVER=ADMINISTRATOR;UID=sa;PWD=123456;DATABASE=WEBSITE_BAN_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")
	'Open all records
	rs.Open "SELECT * FROM Admin", conn, 3, 3 
	
	'Check if username doesn't already exist
	do while not rs.EOF
		if rs("TaiKhoan")=Username then
			set rs=nothing
			set conn=nothing
			Response.Redirect("loginAD.asp?loginAD=createnamefailed")
		end if
		rs.MoveNext
	loop
		
	'Add a record
	rs.AddNew 
	'Put username and password in record
	rs("TaiKhoan")=Username
	rs("MatKhau")=Password
	'Save record
	rs.Update 
	
	set rs=nothing
	set conn=nothing

	Response.Redirect("loginAD.asp?loginAD=creatednew")
%>
