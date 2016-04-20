<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim CTSPDESKTOP__MMColParam
CTSPDESKTOP__MMColParam = "8"
If (Request.Form("MaSP") <> "") Then 
  CTSPDESKTOP__MMColParam = Request.Form("MaSP")
End If
%>
<%
Dim CTSPDESKTOP
Dim CTSPDESKTOP_cmd
Dim CTSPDESKTOP_numRows

Set CTSPDESKTOP_cmd = Server.CreateObject ("ADODB.Command")
CTSPDESKTOP_cmd.ActiveConnection = MM_Connection_STRING
CTSPDESKTOP_cmd.CommandText = "SELECT * FROM dbo.SanPham WHERE MaSP = ?" 
CTSPDESKTOP_cmd.Prepared = true
CTSPDESKTOP_cmd.Parameters.Append CTSPDESKTOP_cmd.CreateParameter("param1", 5, 1, -1, CTSPDESKTOP__MMColParam) ' adDouble

Set CTSPDESKTOP = CTSPDESKTOP_cmd.Execute
CTSPDESKTOP_numRows = 0
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Cửa hàng máy tính | Desktop :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="../css/myStyle.css" type="text/css" rel="stylesheet" >
<link href='http://fonts.googleapis.com/css?family=Lato:400,300,600,700,800' rel='stylesheet' type='text/css'>
<script src="../js/jquery.min.js"></script>

        <!---------------------------
                  LIGHTBOX
        ---------------------------->
<script type="text/javascript" src="../js/jquery.lightbox.js"></script>
<link rel="stylesheet" type="text/css" href="../css/lightbox.css" media="screen" />
  <script type="text/javascript">
    $(function() {
        $('.gallery a').lightBox();
    });
   </script>

<style>HTML,BODY{cursor: url("../images/monkeyani.cur"), url("../images/monkey-ani.gif"), auto;}</style>
</head>
<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("TKKH") = "" then
	Response.write("<a rel=nofollow href=../login.asp?login=createnew class=colorlink2><span><ins>Đăng ký</ins></span></a>|<a rel=nofollow href=../login.asp class=colorlink2><span><ins>Đăng Nhập</ins></span></a>")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=../logout.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
	</div>
</div>

    <!---------------------------
                SEARCH
    ---------------------------->
    <div class="cntr">
        <div class="cntr-innr">
          <label class="search" for="inpt_search">
                <input id="inpt_search" type="text" />
          </label>
            <p>Sờ vào để tìm thứ bạn cần.</p>
        </div>
    </div>
	<div class="pages-top">
	        <div class="logo">
				<a href="../index.asp"><img src="../images/logo.png" alt=""/></a>
	  </div>
      <div class="h_menu4">
    <!---------------------------
                MENU
    ---------------------------->
		<ul class="nav">
					<li><a href="../index.asp">Trang chủ</a></li>
					<li><a href="../laptop/laptop.asp">Laptop</a>
						<ul class="listmenu">
							<li><a href="../laptop/laptop.asp">DELL</a></li>
							<li><a href="../laptop/laptop.asp">HP</a></li>
							<li><a href="../laptop/laptop.asp">APPLE</a></li>
							<li><a href="../laptop/laptop.asp">ACER</a></li>
							<li><a href="../laptop/laptop.asp">ASUS</a></li>
							<li><a href="../laptop/laptop.asp">LENOVO</a></li>
							<li><a href="../laptop/laptop.asp">VAIO</a></li>
						</ul>
					</li>
					<li><a href="../desktop/desktop.asp">Desktop</a>
						<ul class="listmenu">
							<li><a href="../desktop/desktop.asp">DELL</a></li>
							<li><a href="../desktop/desktop.asp">HP</a></li>
							<li><a href="../desktop/desktop.asp">APPLE</a></li>
							<li><a href="../desktop/desktop.asp">ACER</a></li>
							<li><a href="../desktop/desktop.asp">ASUS</a></li>
							<li><a href="../desktop/desktop.asp">LENOVO</a></li>
						</ul>
					</li>
					<li class="active"><a href="../linhkien/linhkien.asp">Linh kiện</a>
					
					</li>
					<li><a href="../phukien/phukien.asp">Phụ kiện</a>
						<ul class="listmenu">
							<li><a href="../phukien/phukien.asp">Headphones</a></li>
							<li><a href="../phukien/phukien.asp">Earphones</a></li>
							<li><a href="../phukien/phukien.asp">Chuột</a></li>
							<li><a href="../phukien/phukien.asp">Keyboard (bàn phím)</a></li>
							<li><a href="../phukien/phukien.asp">USB</a></li>
						</ul>
					</li>
					<li><a href="../lienhe/lienhe.asp">Liên hệ</a></li>
				</ul>
<script type="text/javascript" src="../js/nav.js"></script>
			</div>
            <!-- END MENU -->
			<div class="clear"></div>
</div>
        <!-- End header main -->
     </div>
<!--gallary-->

	 <div class="main">
	 	<div class="wrap">
	 		<div class="pages">
	 		  <div class="cont1 span_2_of_g1">
		      <p>QUÝ KHÁCH ĐANG XEM SẢN PHẨM </p>
		      <p><%=(CTSPDESKTOP.Fields.Item("TenSP").Value)%></p>
			    <p><img src="<%=(CTSPDESKTOP.Fields.Item("HinhAnh").Value)%>" alt="" width="225" height="150"></p>
			    <table width="356" height="155" border="1">
			      <tr>
			        <td width="168">Tên sản phẩm</td>
			        <td width="124"><%=(CTSPDESKTOP.Fields.Item("TenSP").Value)%></td>
		          </tr>
			      <tr>
			        <td>Thông tin sản phẩm</td>
			        <td><%=(CTSPDESKTOP.Fields.Item("CauHinh").Value)%></td>
		          </tr>
			      <tr>
			        <td>Giá</td>
			        <td><%=(CTSPDESKTOP.Fields.Item("Gia").Value)%> VNĐ</td>
		          </tr>
			      <tr>
			        <td>Số lượng</td>
			        <td><%=(CTSPDESKTOP.Fields.Item("SoLuong").Value)%></td>
		          </tr>
		        </table>
			    <p>&nbsp;</p>
				 
				  <div>
				  <form name="form2" method="post" action="../giohang.asp">
				    <label>
				      <input type="image" name="imageField" id="imageField" src="../Images/giohang.jpg" width="100" height="50">
				    </label>
				    Mua hàng
                    <input name="MaSPDatHang" type="hidden" id="MaSPDatHang" value="<%=(CTSPDESKTOP.Fields.Item("MaSP").Value)%>">
                    <input name="TenSP" type="hidden" id="TenSP" value="<%=(CTSPDESKTOP.Fields.Item("TenSP").Value)%>">
				  <input name="HinhAnh" type="hidden" id="HinhAnh" value="<%=(CTSPDESKTOP.Fields.Item("HinhAnh").Value)%>">
				  <input name="Gia" type="hidden" id="Gia" value="<%=(CTSPDESKTOP.Fields.Item("Gia").Value)%>">
				  </form>
			    </div>
				  <div style="clear:both; text-align:center">
				    
			      </div>
              </div>
	 		  <!-- END gallary-->
        <div class="labout span_1_of_g1">
		  <div class="project-list">
	     	<h4>Loại</h4>
			<ul class="blog-list">
				<li>
                    <form name="frmDell" method="post" action=Dell.asp>
                    <a href="RAM.asp">RAM</a>
                    </form>
                </li>
                <li>
                    <form name="frmApple" method="post" action=Apple.asp>
                    <a href="VGA.asp">Card màn hình</a>
                    </form>
                </li>
                <li>
                    <form name="frmHp" method="post" action=Hp.asp>
                    <a href="M.asp">Mainboard</a>
                    </form>
                </li>
				<li>
                    <form name="frmApple" method="post" action=Apple.asp>
                    <a href="SC.asp">Card âm thanh</a>
                    </form>
                </li>
            </ul>
			<div class="clear"></div>
	      </div>
		   <div class="project-list1">
			<div class="clear"></div>
		   </div>
		   <div class="project-list2">
	     	<h4>Các thẻ chọn</h4>
			<ul>
				<li><a href="#">Web Design</a></li>
				<li><a href="#">Html5</a></li>
				<li><a href="#">Wordpress</a></li>
				<li><a href="#">Logo</a></li>
				<li><a href="#">Web Design</a></li>
				<li><a href="#">Web Design</a></li>
				<li><a href="#">Wordpress</a></li>
				<li><a href="#">Web Design</a></li>
				<li><a href="#">Html5</a></li>
				<li><a href="#">Wordpress</a></li>
				<li><a href="#">Logo</a></li>
				<div class="clear"></div>
			</ul>
		   </div>
		 </div>
		   <div class="clear"></div>	
		  </div>
	   </div>
</div>
	<!---------------------------
                BOTTOM
    ---------------------------->
        <div class="footer">
			<div class="wrap">
				<div class="footer-grid footer-grid1">
					<div class="f-logo">
				     <a href="../index.asp"><img src="../images/logo.png" alt=""></a>
			        </div>
					<p>Nhóm gồm 4 thành viên sáng lập, mỗi thành viên điều rất nhiệt tình trong công việc mình đảm nhận.</p>
				</div>
				<div class="footer-grid footer-grid2">
					<h4>Liên Hệ</h4>
				    <ul>
						<li><i class="pin"> </i><div class="extra-wrap">
							<p>392-394 Hoàng Văn Thụ, P.4<br> TP.HCM</p>
						</div></li>
						<li><i class="phone"> </i><div class="extra-wrap">
							<p>+849 3939 3423</p>
						</div></li>
						<li><i class="mail"> </i><div class="extra-wrap1">
							<p>lopaccp1508@gmail.com</p>
						</div></li>
						<li><i class="earth"> </i><div class="extra-wrap1">
							<p>nhom4@abc.com</p>
						</div></li>
					</ul>
				</div>
				<div class="footer-grid footer-grid3">
					<h4>Tiêu Chí</h4>
					<div class="recent-f">
						<div class="recent-f-icon">
							<span> </span>
						</div>
						<div class="recent-f-info">
							<p>Uy Tín</p>
						</div>
						<div class="clear"> </div>
					</div>
					<div class="recent-f1">
						<div class="recent-f-icon">
							<span> </span>
						</div>
						<div class="recent-f-info">
							<p>Chất Lượng</p>
						</div>
						<div class="clear"> </div>
					</div><br />
                    <div class="recent-f2">
						<div class="recent-f-icon">
							<span> </span>
						</div>
						<div class="recent-f-info">
							<p>Chuyên Nghiệp</p>
						</div>
						<div class="clear"> </div>
					</div>
				</div>
				<div class="footer-grid footer-grid4">
					<h4>Nhận Tin Mới</h4>
					<p>Nhập địa chỉ Email để nhận được những tin tức mới nhất về công nghệ</p>
					<form>
					  <input type="text" value="Địa chỉ Email" onFocus="this.value = '';" onBlur="if (this.value == '') {this.value = 'Địa chỉ Email';}">
						<input type="submit" value="">
					</form>
				</div>
				<div class="clear"> </div>
			</div>
		</div>
		<div class="footer-bottom">
	 		  <div class="wrap">
	     	  	<div class="copy">
				   <p>© 2016 Group Four Computer</p>
			    </div>
			    <div class="social">	
				   <ul>	
					  <li class="facebook"><a href="#"><span> </span></a></li>
					  <li class="linkedin"><a href="#"><span> </span></a></li>
					  <li class="twitter"><a href="#"><span> </span></a></li>		
				   </ul>
			    </div>
			    <div class="clear"></div>
			  </div>
       </div>
       
</body>
</html>
<%
CTSPDESKTOP.Close()
Set CTSPDESKTOP = Nothing
%>
