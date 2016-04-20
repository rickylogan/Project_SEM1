<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<head>
    <title>Cửa hàng máy tính | Đăng nhập :: Groupfour</title>
    <link rel="shortcut icon" href="../images/icon.png">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="css/login.css" rel="stylesheet" type="text/css" media="all" />
    <link href='http://fonts.googleapis.com/css?family=Roboto:400' rel='stylesheet' type='text/css'>

</head>
<body>
<div class="cntr">
    <%
        Content = ""							
        QStr = Request.QueryString("loginAD")
        if QStr="passfailed" then
			Content = Content & "<div class=box>"				
            Content = Content & "<p class=noti>Sai mật khẩu</P><A href=Javascript:history.go(-1) class=colorlink>Quay lại</A>"
			Content = Content & "</div>"
        elseif QStr="namefailed" then
		Content = Content & "<div class=box>"
            Content = Content & "<p class=noti>Tên tài khoản không hợp lệ!</P><br><br><A HREF=loginAD.asp class=colorlink>Quay lại đăng nhập</A>	"
			Content = Content & "</div>"
		
        else
            Content = Content & "<form name=frmMain method=POST action=verifyAD.asp>"
            Content = Content & "<input type=text name=txtUsername placeholder='Tên đăng nhập' required>"
            Content = Content & "<input type=password name=txtPassword placeholder='Mật khẩu' required>"
            Content = Content & "<button type=submit name=cmdSubmit>Đăng nhập</button>"
            Content = Content & "</form>"
        end if
    
    %>
    <div align="center">
    	<p class=title align=center><b>Trang đăng nhập</br>Quản trị viên</b></p>
        <%
       		Response.Write(Content)
        %>
    </div>
</div>
</body>
</html>
