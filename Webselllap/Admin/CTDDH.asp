<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim CTDDH__MMColParam
CTDDH__MMColParam = "1"
If (Request.Form("MaDDH") <> "") Then 
  CTDDH__MMColParam = Request.Form("MaDDH")
End If
%>
<%
Dim CTDDH
Dim CTDDH_cmd
Dim CTDDH_numRows

Set CTDDH_cmd = Server.CreateObject ("ADODB.Command")
CTDDH_cmd.ActiveConnection = MM_Connection_STRING
CTDDH_cmd.CommandText = "SELECT a.MaDDH, a.MaSP, a.SoLuong, a.TongTien, a.NgayLap, a.TenSP, b.TinhTrang FROM dbo.CTDDH a, dbo.DonDatHang b WHERE a.MaDDH = ? and a.MaDDH = b.MaDDH" 
CTDDH_cmd.Prepared = true
CTDDH_cmd.Parameters.Append CTDDH_cmd.CreateParameter("param1", 5, 1, -1, CTDDH__MMColParam) ' adDouble

Set CTDDH = CTDDH_cmd.Execute
CTDDH_numRows = 0
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Quản trị viên :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
    <meta name="description" content="">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" href="css/bootstrap.min.css">
    <link rel="stylesheet" href="css/normalize.min.css">
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
if Session("TKA") = "" then
	Response.Redirect("loginAD.asp")
else
	Response.write("Xin chào, <b class=tentk>" & Session("TenAD") & "</b><b class=to> |</b>" & "<a href=logoutAD.asp class=colorlink2><ins>Thoát</ins></a>")
	
end if
%>
</div>
<div id="top"><p class=title align=center>QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
<table style="margin-top:30px" width="40%" border="2px" Bordercolor="black" cellspacing="0" cellpadding="100px" align="center">
  	<tr>
      <td colspan="2" align="center"><b>CHI TIẾT ĐƠN ĐẶT HÀNG</b></td>
	</tr>
    <tr>
      <td width="20%">      Mã đơn đặt hàng</td>
      <td width="20%">      <%=(CTDDH.Fields.Item("MaDDH").Value)%></td>
  </tr>
    <tr>
      <td>      Mã sản phẩm</td>
      <td>      <%=(CTDDH.Fields.Item("MaSP").Value)%></td>
  </tr>
    <tr>
      <td>      Tên sản phẩm</td>
      <td>      <%=(CTDDH.Fields.Item("TenSP").Value)%></td>
  </tr>
    <tr>
      <td>      Số lượng</td>
      <td>      <%=(CTDDH.Fields.Item("SoLuong").Value)%></td>
      </tr>
    <tr>
      <td>      Tổng tiền</td>
      <td>      <%=(CTDDH.Fields.Item("TongTien").Value)%></td>
      </tr>
    <tr>
      <td>      Ngày lập</td>
      <td>      <%=(CTDDH.Fields.Item("NgayLap").Value)%></td>
    </tr>
  </table>
<div style="margin-top:50px" align="center">
	<a href="DDH.asp">
  <button type="button" name="button" id="button" value="">
  QUAY LẠI<br/>ĐƠN ĐẶT HÀNG</button>
	</a>
</div>
<%
	Dim M_TinhTrang
	Content = ""
	M_TinhTrang = (CTDDH.Fields.Item("TinhTrang").Value)
	if M_TinhTrang = "Đã thanh toán       " then
	Content = Content & "<div><form action='Confirm_EditTT.asp' method=post name=form1 id=form1>"
	Content = Content & "<p><input name=MaDDH type=hidden id=MaDDH value=" & (CTDDH.Fields.Item("MaDDH").Value) & "></p>"
	Content = Content & "<input style=float:right type=password name=txtPassword placeholder='MẬT KHẨU CẬP NHẬT LẠI TÌNH TRẠNG ĐƠN ĐẶT HÀNG' required></button></form></div>"
	else
	Content = Content & ""
	End if
	Response.Write(Content)
%>
<script src="js/vendor/jquery-1.10.1.min.js"></script>
<script src="js/plugins.js"></script>
<script src="js/main.js"></script>
<div class="footer-bar">
    <span class="article-wrapper">
        <span class="article-label">Trang quản lý</span>
        <span class="article-link"><a href="#">Lên <ins>TOP▲</ins></a></span>
    </span>
</div>
</body>
</html>
<%
CTDDH.Close()
Set CTDDH = Nothing
%>
