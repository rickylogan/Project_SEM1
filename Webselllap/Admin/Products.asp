<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim SanPham
Dim SanPham_cmd
Dim SanPham_numRows

Set SanPham_cmd = Server.CreateObject ("ADODB.Command")
SanPham_cmd.ActiveConnection = MM_Connection_STRING
SanPham_cmd.CommandText = "SELECT * FROM dbo.SanPham WHERE tinhtrang=1 ORDER BY MaSP desc" 
SanPham_cmd.Prepared = true

Set SanPham = SanPham_cmd.Execute
SanPham_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
SanPham_numRows = SanPham_numRows + Repeat1__numRows
%>
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
    <link rel="stylesheet" href="css/bootstrap.min.css">
    <link rel="stylesheet" href="css/normalize.min.css">
    <link rel="stylesheet" href="css/templatemo_misc.css">
    <link rel="stylesheet" href="css/templatemo_style.css">
<script src="../js/jquery.min.js"></script>
</head>
<body>
        <div class="site-header">
        <div id="templatemo_logo" class="row col-md-4 col-sm-6 col-xs-6">
                            <h1><a href="Products.asp">Admin</a></h1>
            </div>
            <div class="container">
<div class="gocphaimanhinhTV">
<%
if Session("name") = "" then
	Response.Redirect("loginAD.asp")
else
	Response.write("Xin chào, <b class=tentk>" & Session("name") & "</b><b class=to> |</b>" & "<a href=logoutAD.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
</div>
<div id="top"><p class=title align=center>QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
	<div align="center" class="item">
    	<a href="AddSp.asp" class="colorlink">
        <button type="submit" name=cmdSubmit>Thêm sản phẩm mới</button>
        </a>
    </div>
<div class="product-item">
  <table width="85	%" border="0" cellspacing="0" cellpadding="0" align="center">
    <% 
While ((Repeat1__numRows <> 0) AND (NOT SanPham.EOF)) 
%>
  <tr>
    <td width="10%"><p><img src="<%=(SanPham.Fields.Item("HinhAnh").Value)%>" alt="" name="" width="225" height="150"></p></td>
    <td width="35%"><p><%=(SanPham.Fields.Item("TenSP").Value)%></p>
      <p><%=(SanPham.Fields.Item("MaLoai").Value)%></p>
      <p><%=(SanPham.Fields.Item("Gia").Value)%> VNĐ</p>
      <p>Hiện có <%=(SanPham.Fields.Item("SoLuong").Value)%> sản phẩm</p></td>
    <td width="20%"><form action="Editsp.asp" method="post" name="form1" id="form1">
      <input name="MaSP" type="hidden" id="MaSP" value="<%=(SanPham.Fields.Item("MaSP").Value)%>">
      <input name="NSX" type="hidden" id="NSX" value="<%=(SanPham.Fields.Item("MaNSX").Value)%>">
      <input name="Loai" type="hidden" id="Loai" value="<%=(SanPham.Fields.Item("MaLoai").Value)%>">
      <button type="submit" name="button" id="button" value="CẬP NHẬT">CẬP NHẬT</button>
    </form></td>
    <td width="20%"><form action="Removesp.asp" method="post" name="form1" id="form1">
      <input name="MaSp" type="hidden" id="MaSp" value="<%=(SanPham.Fields.Item("MaSP").Value)%>">
      <button type="submit" name="button2" id="button2" value="XÓA">XÓA</button>
    </form></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  SanPham.MoveNext()
Wend
%>
  </table>

</div>
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
<%
SanPham.Close()
Set SanPham = Nothing
%>