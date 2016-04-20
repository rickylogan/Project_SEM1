<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<head>
    <title>Cửa hàng máy tính | Đăng nhập :: Groupfour</title>
    <link rel="shortcut icon" href="images/icon.png">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="css/login.css" rel="stylesheet" type="text/css" media="all" />
    <link href='http://fonts.googleapis.com/css?family=Roboto:400' rel='stylesheet' type='text/css'>

<style>HTML,BODY{cursor: url("images/monkeyani.cur"), url("images/monkey-ani.gif"), auto;}</style>
</head>
<body>
<div class="pages-top">
    <div class="logo">
        <a href="index.asp"><img src="images/logo.png" alt=""/></a>
    </div>
</div>
<div class="cntr">
    <%
        Content = ""							
        QStr = Request.QueryString("login")		
    
        if ucase(left(QStr,6))="CREATE" then 
            Title = "Trang đăng ký"
        else
            Title = "Trang đăng nhập"
        end if
        if QStr="passfailed" then
			Content = Content & "<div class=box>"				
            Content = Content & "<p class=noti>Sai mật khẩu</P><A href=Javascript:history.go(-1) class=colorlink>Quay lại</A>"
			Content = Content & "</div>"
        elseif QStr="createpassfailed" then
			Content = Content & "<div class=box>"		
            Content = Content & "<div><p class=noti>Đăng ký thất bại!</p><A href=Javascript:history.go(-1) class=colorlink>Quay lại đăng ký</A></br><A HREF=login.asp class=colorlink>Hủy đăng ký</A></div>"
			Content = Content & "</div>"
        elseif QStr="namefailed" then
		Content = Content & "<div class=box>"
            Content = Content & "<p class=noti>Tên tài khoản không hợp lệ!</P><A HREF=login.asp?login=createnew class=colorlink>Nhấp vào đây để đăng ký</A><br><br><A HREF=login.asp class=colorlink>Quay lại đăng nhập</A>	"
			Content = Content & "</div>"
        elseif QStr="createnamefailed" then
			Content = Content & "<div class=box>"	
            Content = Content & "<p class=noti>Tên tài khoản đã được dùng!!!</P><A HREF=Javascript:history.go(-1) class=colorlink>Quay lại</A></br><A HREF=login.asp class=colorlink>Hủy đăng ký</A>"
			elseif QStr="createemailfailed" then
			Content = Content & "<div class=box>"	
            Content = Content & "<p class=noti>Email đã được sử dụng!!!</P><A HREF=Javascript:history.go(-1) class=colorlink>Quay lại</A></br><A HREF=login.asp class=colorlink>Hủy đăng ký</A>"
			Content = Content & "</div>"
        elseif QStr="creatednew" then
			Content = Content & "<div class=box>"
            Content = Content & "<p class=noti>Tạo tài khoản thành công!</p><A HREF=login.asp class=colorlink>Đăng nhập</A>"
			Content = Content & "</div>"
        elseif QStr="createnew" then
            Content = Content & "<form name=frmCreate method=POST action=create.asp>"
			Content = Content & "<input type=text name=txtFullname placeholder='Họ và Tên' required>"
            Content = Content & "<input type=text name=txtUsername placeholder='Tên đăng nhập' required maxlength=16>"
            Content = Content & "<input type=password name=txtPassword placeholder='Mật khẩu' required>"
            Content = Content & "<input type=email name=txtEmail placeholder='Email' required>"
			Content = Content & "<input type=text name=txtAddress placeholder='Địa chỉ' required>"
            Content = Content & "<input type=tell name=txtPhone placeholder='Số điện thoại' required>"
            Content = Content & "<button type=submit name=cmdSubmit value='Đăng ký'>Đăng ký</button>"
			Content = Content & "</br></br><A href=login.asp class=colorlink>Quay lại đăng nhập</A>"
            Content = Content & "</form>"
        else
            Content = Content & "<form name=frmMain method=POST action=verify.asp>"
            Content = Content & "<input type=text name=txtUsername placeholder='Tên đăng nhập' required>"
            Content = Content & "<input type=password name=txtPassword placeholder='Mật khẩu' required>"
            Content = Content & "<button type=submit name=cmdSubmit>Đăng nhập</button>"
            Content = Content & "</form>"
            Content = Content & "</br></br><A HREF=login.asp?login=createnew class=colorlink>Nhấp vào đây để đăng ký</A>"
        end if
    
    %>
    <div align="center">
        <%
        Response.Write("<p class=title><td align=center><b>" & Title & "</p>")
        Response.Write(Content)
        %>
    </div>
</div>
</body>
</html>
