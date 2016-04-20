<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Connection.asp" -->
<%
Dim SanPham
Dim SanPham_cmd
Dim SanPham_numRows

Set SanPham_cmd = Server.CreateObject ("ADODB.Command")
SanPham_cmd.ActiveConnection = MM_Connection_STRING
SanPham_cmd.CommandText = "SELECT MaSP, TenSP, MaNSX, HinhAnh FROM dbo.SanPham ORDER BY MaSP DESC" 
SanPham_cmd.Prepared = true

Set SanPham = SanPham_cmd.Execute
SanPham_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 3
Repeat1__index = 0
SanPham_numRows = SanPham_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Cửa hàng máy tính | Trang chủ :: Groupfour</title>
<link rel="shortcut icon" href="images/icon.png">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Lato:400,300,600,700,800' rel='stylesheet' type='text/css'>
	<!--jquery lightbox-->
<script type="text/javascript" src="../js/jquery.lightbox.js"></script>
<link rel="stylesheet" type="text/css" href="../css/lightbox.css" media="screen" />

        <!---------------------------
                     CLOCK
        ---------------------------->
    <link href="css/style2.css" rel="stylesheet" />
                <!-- JavaScript Includes -->
<script src="js/jquery.min(v1.10.1).js"></script>
<script src="js/moment.min.js"></script>
<script src="js/script.js"></script>

<script src="js/jquery.min.js"></script>
<!---strat-slider---->
	    <link rel="stylesheet" type="text/css" href="css/slider.css" />
<script type="text/javascript" src="js/modernizr.custom.28468.js"></script>
<script type="text/javascript" src="js/jquery.cslider.js"></script>
<script type="text/javascript">
				$(function() {
				
					$('#da-slider').cslider({
						autoplay	: true,
						bgincrement	: 450
					});
				
				});
			</script>
			
<!---//strat-slider---->
<script src="js/jquery.magnific-popup.js" type="text/javascript"></script>
<link href="css/magnific-popup.css" rel="stylesheet" type="text/css">
<script>
			$(document).ready(function() {
				$('.popup-with-zoom-anim').magnificPopup({
					type: 'inline',
					fixedContentPos: false,
					fixedBgPos: true,
					overflowY: 'auto',
					closeBtnInside: true,
					preloader: false,
					midClick: true,
					removalDelay: 300,
					mainClass: 'my-mfp-zoom-in'
			});
		});
		</script>

<style>HTML,BODY{cursor: url("images/monkeyani.cur"), url("images/monkey-ani.gif"), auto;}</style>
</head>
<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("TKKH") = "" then
	Response.write("<a rel=nofollow href=login.asp?login=createnew class=colorlink2><span><ins>Đăng ký</ins></span></a>|<a rel=nofollow href=login.asp class=colorlink2><span><ins>Đăng Nhập</ins></span></a>")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=logout.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
	</div>
</div>

    <!---------------------------
                SEARCH
    ---------------------------->
    <div class="cntr">
        <div class="cntr-innr">
          <form  action="Search/Search.asp" method="post" id="form1" class="search" for="inpt_search">
                <input name="txtSearch" type="text" id="inpt_search" />
            </form>
            <p>Tìm kiếm</p>
            <p style="margin-left: 1000px"><a href="HienThi.asp" class="colorlink"> Giỏ Hàng </a><%Session("dem")%></p>
      </div>
    </div>
  
<div class="pages-top">
    <div class="logo">
        <a href="index.asp"><img src="images/logo.png" alt=""/></a>
    </div>
             
		     <div class="h_menu4">
    <!---------------------------
                MENU
    ---------------------------->
				<ul class="nav">
					<li class="active"><a href="index.asp">Trang chủ</a></li>
					<li><a href="Laptop/Laptop.asp">Laptop</a>
						<ul>
							<li>
                                <form name="frmDell" method="post" action=laptop/Dell.asp>
                                <a href="Laptop/Dell.asp">DELL</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmHp" method="post" action=laptop/Hp.asp>
                                <a href="Laptop/Hp.asp">HP</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmApple" method="post" action=laptop/Apple.asp>
                                <a href="Laptop/Apple.asp">APPLE</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAsus" method="post" action=laptop/sus.asp>
                                <a href="Laptop/Asus.asp">ASUS</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAcer" method="post" action=laptop/Acer.asp>
                                <a href="Laptop/Acer.asp">ACER</a>
                                </form>
                            </li>
                            <li>
                                <form name="Lenovo" method="post" action=laptop/Lenovo.asp>
                                <a href="Laptop/Lenovo.asp">LENOVO</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmVaio" method="post" action=laptop/Vaio.asp>
                                <a href="Laptop/Vaio.asp">VAIO</a>
                                </form>
                            </li>
						</ul>
					</li>
					<li><a href="Desktop/Desktop.asp">Desktop</a>
						<ul>
							<li>
                                <form name="frmDell" method="post" action=Desktop/Dell.asp>
                                <a href="Desktop/Dell.asp">DELL</a>
                                </form>
                            </li>
							<li>
                                <form name="frmHp" method="post" action=Desktop/Hp.asp>
                                <a href="Desktop/Hp.asp">HP</a>
                                </form>
                            </li>
							<li>
                                <form name="frmApple" method="post" action=Desktop/Apple.asp>
                                <a href="Desktop/Apple.asp">APPLE</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAsus" method="post" action=Desktop/Asus.asp>
                                <a href="Desktop/Asus.asp">ASUS</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAcer" method="post" action=Desktop/Acer.asp>
                                <a href="Desktop/Acer.asp">ACER</a>
                                </form>
                            </li>
                            <li>
                                <form name="Lenovo" method="post" action=Desktop/Lenovo.asp>
                                <a href="Desktop/Lenovo.asp">LENOVO</a>
                         	   </form>
							</li>
         			   </ul>
					</li>
					<li><a href="Linhkien/Linhkien.asp">Linh kiện</a>
						<ul class="listmenu">
                        	<li>
                                <form name="frmRAM" method="post" action=Linhkien/RAM.asp>
                                <a href="Linhkien/RAM.asp">RAM</a>
                                </form>
                            </li>
							<li>
                                <form name="frmVGA" method="post" action=Linhkien/RAM.asp>
                                <a href="Linhkien/VGA.asp">Card màn hình</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmM" method="post" action=Linhkien/RAM.asp>
                                <a href="Linhkien/M.asp">Mainboard</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmSC" method="post" action=Linhkien/RAM.asp>
                                <a href="Linhkien/SC.asp">Card âm thanh</a>
                                </form>
                            </li>
						</ul>
					</li>
					<li><a href="Phukien/Phukien.asp">Phụ kiện</a>
						<ul class="listmenu">
                        	<li>
                                <form name="frmHP" method="post" action=Phukien/HP.asp>
                                <a href="Phukien/HP.asp">Headphones</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmEP" method="post" action=Phukien/EP.asp>
                                <a href="Phukien/EP.asp">Earphones</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmCQ" method="post" action=Phukien/CQ.asp>
                                <a href="Phukien/CQ.asp">Chuột</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmBP" method="post" action=Phukien/BP.asp>
                                <a href="Phukien/BP.asp">Bàn Phím</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmUSB" method="post" action=Phukien/USB.asp>
                                <a href="Phukien/USB.asp">USB</a>
                                </form>
                            </li>
						</ul>
					</li>
					<li><a href="Lienhe/Lienhe.asp">Liên hệ</a></li>
				</ul>
				<script type="text/javascript" src="js/nav.js"></script>
			</div>
            <!--END MENU-->
 
 			<div class="clear"></div>
		</div>
        <!-- end header_main4 -->
        <div class="slider">
				<!---start-da-slider----->
			  <div id="da-slider" class="da-slider">
				  <div class="da-slide">
					<div align="center">
                        <a href="Laptop/Laptop.asp" target="_blank">
                            <img src="images/Slide1.jpg" alt="" />
                        </a>
					</div>
				  </div>
				  <div class="da-slide">
					<div align="center">
                        <a href="Laptop/Laptop.asp" target="_blank">
                            <img src="images/Slide2.jpg" alt="" />
                        </a>
					</div>
				  </div>
				  <div class="da-slide">
					<div align="center">
                        <a href="Laptop/Laptop.asp" target="_blank">
                            <img src="images/Slide3.jpg" alt="" />
                        </a>
					</div>
				  </div>
				  <div class="da-slide">
					<div align="center">
                        <a href="phukien/phukien.asp" target="_blank">
                            <img src="images/Slide4.jpg" alt="" />
                        </a>
					</div>
				  </div>		
				  <nav class="da-arrows">
					<span class="da-arrows-prev"></span>
				 	<span class="da-arrows-next"></span>
				  </nav>
			   </div>
 		       <!---//End-da-slider----->
	      </div>
	 </div>
     
    <!-----------------
		SẢN PHẨM MỚI
    ------------------>
	 <div class="main">
	     <div class="wrap">
	 	   <h2 class="m_1"><a href="Laptop/Laptop.asp" class="colorlink3">SẢN PHẨM MỚI</a></h2>
	 	     <div class="content-top">
               <% 
While ((Repeat1__numRows <> 0) AND (NOT SanPham.EOF)) 
%>
  <div class="col_1_of_4 span_1_of_4">
	<p align="center">
    	<img src="<%=(SanPham.Fields.Item("HinhAnh").Value)%>" alt="" width="225" height="150"/>
	</p>                   
    <div class="desc">
      <h3><%=(SanPham.Fields.Item("TenSP").Value)%></h3>
      <p>
      </p>
      </div>
  </div>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  SanPham.MoveNext()
Wend
%>
<div class="clear"></div>
		     </div>
             
		     <!--- NOTICE
		     <div class="content-middle-bottom">
	 	    	<div class="col_1_of_middle span_1_of_middle">
	 	    	   <h3>Tiêu đề</h3>
	 	    	   <p>Mô tả</p>
	 	    	   <div class="btn"><a href="#">Chi tiết</a></div>
	 	    	</div>
				<div class="col_1_of_middle span_1_of_middle">
					<ul class="progress vertical">
					  <li class="bar bar-success" style="width:100%;height:70%;"> </li>
					</ul>
					<ul class="progress vertical">
					  <li class="bar bar-success" style="width:100%;height:60%;"> </li>
				    </ul>
				    <ul class="progress vertical">
					  <li class="bar bar-success" style="width:100%;height:20%;"> </li>
				    </ul>
				    <ul class="progress vertical">
					  <li class="bar bar-success" style="width:100%;height:40%;"> </li>
				    </ul>
				    <ul class="progress vertical">
					  <li class="bar bar-success" style="width:100%;height:10%;"> </li>
				    </ul>
			    </div>
				<div class="clear"></div>
		     </div> 
		     ----->
             
	<!----------------------
		SẢN PHẨM NỔI BẬT
	----------------------->
		     <div class="m_3"><span class="left_line"></span> Sản Phẩm Nổi Bật<span class="right_line"> </span></div>
		      <div class="content-top">
				 <div class="col_1_of_projects span_1_of_projects"><a href="#">
				   <div class="view view-first">
                    <img src="images/pic3.jpg" alt=""/>
                      <div class="mask">
                        <h2>Chuột quang</h2>
                        <p>Chuột phát ra ánh hào quang.</p>
                         <a class="popup-with-zoom-anim" href="#small-dialog3"> <div class="info">Xem thêm...</div></a>
		                     <div id="small-dialog3" class="mfp-hide">
							   <div class="pop_up2">
							   	  <img src="images/pic3.jpg"/ alt=""/>
							   	  <h3 class="m_4"><a href="#">Chuột quang</a></h3>
				                  <p class="m_5">Chuột phát ra ánh hào quang.</p>
							   </div>
							 </div>
                        </div>
                     </div> 
					<h3 class="m_4"><a href="#">Phụ kiện</a></h3>
				    <p class="m_5">Chuột quang</p>
				  </a> </div>
				 <div class="col_1_of_projects span_1_of_projects"><a href="#">
				   <div class="view view-first">
                    <img src="images/pic6.jpg" alt=""/>
                      <div class="mask">
                        <h2>Headphones</h2>
                        <p>Tai nghe hình tai mèo.</p>
                         <a class="popup-with-zoom-anim" href="#small-dialog4"> <div class="info">Xem thêm...</div></a>
		                     <div id="small-dialog4" class="mfp-hide">
							   <div class="pop_up2">
							   	  <img src="images/pic6.jpg"/ alt=""/>
							   	   <h3 class="m_4"><a href="#">Headphones</a></h3>
				    				<p class="m_5">Tai nghe hình tai mèo.<br>Hàng chính hãng.<br>Tem bảo hành đầy đủ.</p>
							   </div>
							 </div>
                        </div>
                     </div> 
					<h3 class="m_4"><a href="#">Phụ kiện</a></h3>
				    <p class="m_5">Tai nghe có tai mèo</p>
				  </a> </div>
				 <div class="col_1_of_projects span_1_of_projects"><a href="#">
				   <div class="view view-first">
                    <img src="images/pic5.jpg" alt=""/>
                      <div class="mask">
                        <h2>Laptop HP</h2>
                        <p>Một chiếc laptop hãng HP có màu đỏ mới...</p>
                         <a class="popup-with-zoom-anim" href="#small-dialog5"> <div class="info">Xem thêm...</div></a>
		                     <div id="small-dialog5" class="mfp-hide">
							   <div class="pop_up2">
							   	  <img src="images/pic5.jpg"/ alt=""/>
							   	   <h3 class="m_4"><a href="#">Laptop HP màn hình rời</a></h3>
				    			   <p class="m_5">Một chiếc laptop hãng HP có màu đỏ mới cứng.<br>Hàng chính hãng.<br>Tem bảo hành đầy đủ.</p>
							   </div>
							 </div>
                        </div>
                     </div> 
					<h3 class="m_4"><a href="#">Laptop</a></h3>
				    <p class="m_5">Laptop màn hình rời</p>
				  </a> </div>
				 <div class="col_1_of_projects span_1_of_projects"><a href="#">
				   <div class="view view-first">
                    <img src="images/pic4.jpg" alt=""/>
                      <div class="mask">
                        <h2>Laptop</h2>
                        <p>Một chiếc laptop có thể xoay màn hình...</p>
                         <a class="popup-with-zoom-anim" href="#small-dialog6"> <div class="info">Xem thêm...</div></a>
		                     <div id="small-dialog6" class="mfp-hide">
							   <div class="pop_up2">
							   	  <img src="images/pic4.jpg"/ alt=""/>
							   	   <h3 class="m_4"><a href="#">Tiêu đề hiển thị</a></h3>
				    				<p class="m_5">Một chiếc laptop có thể xoay màn hình 360 độ.<br>Hàng chính hãng.<br>Tem bảo hành đầy đủ.</p>
							   </div>
							 </div>
                        </div>
                     </div> 
					<h3 class="m_4"><a href="#">Laptop</a></h3>
				    <p class="m_5">Laptop màn hình xoay</p>
				  </a> </div>
				<div class="clear"></div>
		 </div>
<!---------------------------
             NOTE
---------------------------->
		 <div class="m_3"><span class="left_line"></span>NOTE<span class="right_line"> </span></div>
		      <div class="wmuSlider example">
			    <div class="wmuSliderWrapper">
				   <article style="position: absolute; width: 100%; opacity: 0;"> 
				   	   <div class="banner-text">
						  <p>NOTE 1<br>NOTE 2</p>
					   </div>
					   <h3 class="m_6">Cửa hàng máy tính, <br><span class="bg1">Cty Group Four</span></h3>
				   </article>
				   <article style="position: relative; width: 100%; opacity: 1;"> 
				   	   <div class="banner-text">
						  <p>NOTE 3<br>NOTE 4</p>
					   </div>
					   <h3 class="m_6">Cửa hàng máy tính, <br><span class="bg1">Cty Group Four</span></h3>
				   </article>
				   <article style="position: absolute; width: 100%; opacity: 0;">
				       <div class="banner-text">
						  <p>NOTE 5<br>NOTE 6</p>
					   </div>
					    <h3 class="m_6">Cửa hàng máy tính, <br><span class="bg1">Cty Group Four</span></h3>
				   </article>
				   <article style="position: absolute; width: 100%; opacity: 0;">
				        <div class="banner-text">
						  <p>NOTE 7<br>NOTE 8</p>
					   </div>
					    <h3 class="m_6">Cửa hàng máy tính, <br><span class="bg1">Cty Group Four</span></h3>
				   </article>
				  </div>
                 <ul class="wmuSliderPagination">
                	<li><a href="#" class="">0</a></li>
                	<li><a href="#" class="class="wmuActive"">1</a></li>
                	<li><a href="#" >2</a></li>
                	<li><a href="#">3</a></li>
                </ul>
              </div>
                 <script src="js/jquery.wmuSlider.js"></script> 
				   <script>
       					$('.example').wmuSlider();         
   				   </script> 	
			</div>
            <!--END NOTE-->
        <!---------------------------
                     CLOCK
        ---------------------------->
    <div id="clock" class="light">
        <div class="display">
            <div class="weekdays"></div>
            <div class="ampm"></div>
            <div class="alarm"></div>
            <div class="digits"></div>
        </div>
    </div>
           <!--END CLOCK-->
        
    <!---------------------------
                BOTTOM
    ---------------------------->
        <div class="footer">
			<div class="wrap">
				<div class="footer-grid footer-grid1">
					<div class="f-logo">
				     <a href="index.asp"><img src="images/logo.png" alt=""></a>
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
SanPham.Close()
Set SanPham = Nothing
%>
