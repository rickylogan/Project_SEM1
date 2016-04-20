<%
	'Save entered username and password
	Fullname = Request.Form("txtFullname")
	Username = Request.Form("txtUsername")
	Password = Request.Form("txtPassword")
	Email = Request.Form("txtEmail")
	Address = Request.Form("txtAddress")
    Phone = Request.Form("txtPhone")
		
	'Check if username and password are entered
	if Username = "" then Response.redirect("login.asp?login=createnew")
	if Password = "" then Response.Redirect("login.asp?login=createpassfailed")
	
	'Build connection
	set conn = server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={SQL Server};SERVER=localhost;UID=sa;PWD=123456;DATABASE=CUA_HANG_MAY_TINH;"
	set rs = server.CreateObject ("ADODB.Recordset")
	'Open all records
	rs.Open "SELECT * FROM KhachHang", conn, 3, 3 
	
	'Check if username doesn't already exist
	do while not rs.EOF
		if rs("TKKH")=Username then
			set rs=nothing
			set conn=nothing
			Response.Redirect("login.asp?login=createnamefailed")
		end if
		if rs("Email")=Email then
			set rs=nothing
			set conn=nothing
			Response.Redirect("login.asp?login=createemailfailed")
		end if
		rs.MoveNext
	loop
		
	'Add a record
	rs.AddNew 
	'Put username and password in record
	rs("TenKH")=Fullname
	rs("TKKH")=Username
	rs("matkhau")=Password
	rs("Email")=Email
	rs("DiaChi")=Address
    rs("SDT")=Phone
	'Save record
	rs.Update 
	
	set rs=nothing
	set conn=nothing

	Response.Redirect("login.asp?login=creatednew")
%>
