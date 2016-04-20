<!DOCTYPE HTML>
<html>
<head>
<title>Quản trị viên :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
    <meta name="description" content="">
    <meta name="viewport" content="width=device-width">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href='http://fonts.googleapis.com/css?family=Roboto:400' rel='stylesheet' type='text/css'>
    <link href="http://fonts.googleapis.com/css?family=Open+Sans:300italic,400italic,600italic,700italic,800italic,400,300,600,700,800" rel="stylesheet">
    <link rel="stylesheet" href="../css/bootstrap.min.css">
    <link rel="stylesheet" href="../css/normalize.min.css">
    <link rel="stylesheet" href="../css/animate.css">
    <link rel="stylesheet" href="../css/templatemo_misc.css">
    <link rel="stylesheet" href="../css/templatemo_style.css">
 <script src="../js/jquery.min.js"></script>
</head>
<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("name") = "" then
	Response.write("<a rel=nofollow href=loginAD.asp?loginAD=createnew class=colorlink2><span><ins>Đăng ký</ins></span></a>|<a rel=nofollow href=loginAD.asp class=colorlink2><span><ins>Đăng Nhập</ins></span></a>")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=loginAD.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
	</div>
</div>
    <div id="top">
        <div class="site-header">
            <div class="container">
                <div class="row">
                    <div class="col-md-4 col-sm-6 col-xs-6">
                        <div id="templatemo_logo">
                            <h1><a href="Products.asp">Admin</a></h1>
                        </div>
                    </div>
                </div>
            </div> <!-- /.container -->
        </div> <!-- /.site-header -->
    </div> <!-- /#front -->
    <div class="site-slider"></div>
<div class="clear"></div>
    <script src="js/vendor/jquery-1.10.1.min.js"></script>
    <script src="js/plugins.js"></script>
    <script src="js/main.js"></script>
<div class="footer-bar">
    <span class="article-wrapper">
        <span class="article-label">Trang quản lý</span>
        <span class="article-link"><a href="#top">Lên top</a></span>
    </span>
</div>
</body>
</html>