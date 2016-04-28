<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim LAPTOP
Dim LAPTOP_cmd
Dim LAPTOP_numRows

Set LAPTOP_cmd = Server.CreateObject ("ADODB.Command")
LAPTOP_cmd.ActiveConnection = MM_Connection_STRING
LAPTOP_cmd.CommandText = "SELECT * FROM dbo.SanPham WHERE Tinhtrang = 1 and MaLoai = 1  and MaNSX = 5" 
LAPTOP_cmd.Prepared = true

Set LAPTOP = LAPTOP_cmd.Execute
LAPTOP_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 9
Repeat1__index = 0
LAPTOP_numRows = LAPTOP_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim LAPTOP_total
Dim LAPTOP_first
Dim LAPTOP_last

' set the record count
LAPTOP_total = LAPTOP.RecordCount

' set the number of rows displayed on this page
If (LAPTOP_numRows < 0) Then
  LAPTOP_numRows = LAPTOP_total
Elseif (LAPTOP_numRows = 0) Then
  LAPTOP_numRows = 1
End If

' set the first and last displayed record
LAPTOP_first = 1
LAPTOP_last  = LAPTOP_first + LAPTOP_numRows - 1

' if we have the correct record count, check the other stats
If (LAPTOP_total <> -1) Then
  If (LAPTOP_first > LAPTOP_total) Then
    LAPTOP_first = LAPTOP_total
  End If
  If (LAPTOP_last > LAPTOP_total) Then
    LAPTOP_last = LAPTOP_total
  End If
  If (LAPTOP_numRows > LAPTOP_total) Then
    LAPTOP_numRows = LAPTOP_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = LAPTOP
MM_rsCount   = LAPTOP_total
MM_size      = LAPTOP_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
LAPTOP_first = MM_offset + 1
LAPTOP_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (LAPTOP_first > MM_rsCount) Then
    LAPTOP_first = MM_rsCount
  End If
  If (LAPTOP_last > MM_rsCount) Then
    LAPTOP_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Cửa hàng máy tính | Laptop :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Lato:400,300,600,700,800' rel='stylesheet' type='text/css'>
<script src="../js/jquery.min.js"></script>

<style>HTML,BODY{cursor: url("../images/monkeyani.cur"), url("../images/monkey-ani.gif"), auto;}</style>
</head>
<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("TKKH") = "" then
	Response.write("<a rel=nofollow href=../login.asp?login=createnew class=colorlink2><span><ins>Đăng ký</ins></span></a>|<a rel=nofollow href=../login.asp class=colorlink2><span><ins>Đăng Nhập</ins></span></a>")
	Response.write("<div style=margin-top:-20px class=cntr>")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=../logout.asp class=colorlink2 <ins>Thoát<ins></a>")
	Response.write("<div><p algin=right class=thongtincanhan><a href=../SuaTTCN.asp rel=nofollow class=colorlink><span><ins>Thông Tin Cá Nhân</ins></span></a></p></div>")
	Response.write("<div style=margin-top:-45px class=cntr>")
end if
%>
    <!---------------------------
                SEARCH
    ---------------------------->
		<div class="cntr-innr">
              <form  action="../Search/Search.asp" method="post" id="form1" class="search" for="inpt_search">
                    <input name="txtSearch" type="text" id="inpt_search" />
                </form>
                <p>Tìm kiếm</p>
          </div>
        </div>
	</div>
</div>
            <!---------------------------
                Giỏ hàng
            ---------------------------->      
        <link rel="stylesheet" type="text/css" href="../css/giohang.css" media="all" />
        <div id="wrapper">
          <div class="cart-tab visible">		
            <a href="../HienThi.asp" title="Xem giỏ hàng của bạn" class="cart-link">
              <span class="contents"><% if Session("dem")="" then Response.Write("0") else Response.Write(Session("dem"))%> sản phẩm</span>
              <span class="amount">
			  <%	if	Session("tongtien")="" or Session("tongtien")="0" then
			  			Response.Write("0 ₫") 
					elseif	x = Session("tongtien") then
								if len(x) mod 3 = 0 then
									Response.Write(right(left(FormatCurrency(x),4*len(x)\3),- 1 + 4*len(x)\3))
								else Response.Write(right(left(FormatCurrency(x),1+ 4*len(x)\3),0+ 4*len(x)\3))
								end if
					end if %>
               </span>
            </a>
            <div class="cart">
              <h2 class="text_giohang">Giỏ hàng</h2>
              <div class="cart-items">
                <ul>
                  <li class="clearfix">
                    <img src="anh.jpg" class="productimg">
                    <h4>Dark Hoodie</h4>
                    <span class="item-price">$11.00</span>
                    <span class="quantity">Số lượng: </span>
                  </li>
                </ul>
              </div><!-- @end .cart-items -->
              <a href="<%if session("TKKH")="" then response.Write("../ThongTinKHMoi.asp") else response.write("../ThongTinKHDangNhap.asp") end if %>" class="checkout">Thanh toán →</a>
            </div><!-- @end .cart -->
          </div>
        </div>
          <!-- End Giỏ hàng -->
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
					<li class="active"><a href="../Laptop/Laptop.asp">Laptop</a>
						
					</li>
					<li><a href="../Desktop/Desktop.asp">Desktop</a>
						<ul>
							<li>
                                <form name="frmDell" method="post" action=Desktop/Dell.asp>
                                <a href="../Desktop/Dell.asp">DELL</a>
                                </form>
                            </li>
							<li>
                                <form name="frmHp" method="post" action=Desktop/Hp.asp>
                                <a href="../Desktop/Hp.asp">HP</a>
                                </form>
                            </li>
							<li>
                                <form name="frmApple" method="post" action=Desktop/Apple.asp>
                                <a href="../Desktop/Apple.asp">APPLE</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAsus" method="post" action=Desktop/Asus.asp>
                                <a href="../Desktop/Asus.asp">ASUS</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAcer" method="post" action=Desktop/Acer.asp>
                                <a href="../Desktop/Acer.asp">ACER</a>
                                </form>
                            </li>
                            <li>
                                <form name="Lenovo" method="post" action=Desktop/Lenovo.asp>
                                <a href="../Desktop/Lenovo.asp">LENOVO</a>
                         	   </form>
							</li>
         			   </ul>
					</li>
					<li><a href="../Linhkien/Linhkien.asp">Linh kiện</a>
						<ul class="listmenu">
                        	<li>
                                <form name="frmRAM" method="post" action=Linhkien/RAM.asp>
                                <a href="../Linhkien/RAM.asp">RAM</a>
                                </form>
                            </li>
							<li>
                                <form name="frmVGA" method="post" action=Linhkien/RAM.asp>
                                <a href="../Linhkien/VGA.asp">Card màn hình</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmM" method="post" action=Linhkien/RAM.asp>
                                <a href="../Linhkien/M.asp">Mainboard</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmSC" method="post" action=Linhkien/RAM.asp>
                                <a href="../Linhkien/SC.asp">Card âm thanh</a>
                                </form>
                            </li>
						</ul>
					</li>
					<li><a href="../Phukien/Phukien.asp">Phụ kiện</a>
						<ul class="listmenu">
                        	<li>
                                <form name="frmHP" method="post" action=Phukien/HP.asp>
                                <a href="../Phukien/HP.asp">Headphones</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmEP" method="post" action=Phukien/EP.asp>
                                <a href="../Phukien/EP.asp">Earphones</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmCQ" method="post" action=Phukien/CQ.asp>
                                <a href="../Phukien/CQ.asp">Chuột</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmBP" method="post" action=Phukien/BP.asp>
                                <a href="../Phukien/BP.asp">Bàn Phím</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmUSB" method="post" action=Phukien/USB.asp>
                                <a href="../Phukien/USB.asp">USB</a>
                                </form>
                            </li>
						</ul>
					</li>
					<li><a href="../Lienhe/Lienhe.asp">Liên hệ</a></li>
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
				  <p>CÁC SẢN PHẨM LAPTOP APPLE MỚI NHẤT</p>
				  <p>&nbsp;</p>
                  <% 
While ((Repeat1__numRows <> 0) AND (NOT LAPTOP.EOF)) 
%>
  <div class="oneItem">
    <p align="center"><a href="ctspLaptop.asp?<%=(LAPTOP.Fields.Item("MaSP").Value)%>"><img src="<%=(LAPTOP.Fields.Item("HinhAnh").Value)%>" alt="" width="225" height="150"></a></p>
    <p><%=(LAPTOP.Fields.Item("TenSP").Value)%></p>
    <p>Giá: <%	x = LAPTOP.Fields.Item("Gia").Value
		  		if len(x) mod 3 = 0 then
		  			Response.Write(right(left(FormatCurrency(x),4*len(x)\3),- 1 + 4*len(x)\3))
		  		else Response.Write(right(left(FormatCurrency(x),1+ 4*len(x)\3),0+ 4*len(x)\3))
				end if
			%><strong><em> <u>VNĐ</u></em></strong></p>
    <p>Hiện còn <strong><%=(LAPTOP.Fields.Item("SoLuong").Value)%></strong> sản phẩm</p>
    <p>&nbsp;</p>    
  </div>
  
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  LAPTOP.MoveNext()
Wend
%>
              </div>
              
  
 </div>
       </div>
               
              
<!-- END gallary-->
        <div class="labout span_1_of_g1">
		  <div class="project-list">
	     	<h4>Loại</h4>
			<ul class="blog-list">
				<li>
                    <form name="frmDell" method="post" action=Dell.asp>
                    <a href="Dell.asp">DELL</a>
                    </form>
                </li>
                <li>
                    <form name="frmHp" method="post" action=Hp.asp>
                    <a href="Hp.asp">HP</a>
                    </form>
                </li>
				<li>
                    <form name="frmApple" method="post" action=Apple.asp>
                    <a href="Apple.asp">APPLE</a>
                    </form>
                </li>
				<li>
                    <form name="frmAsus" method="post" action=Asus.asp>
                    <a href="Asus.asp">ASUS</a>
                    </form>
                </li>
            </ul>
			<ul class="blog-list">
                <li>
                    <form name="frmAcer" method="post" action=Acer.asp>
                    <a href="Acer.asp">ACER</a>
                    </form>
                </li>
                <li>
                    <form name="Lenovo" method="post" action=Lenovo.asp>
                    <a href="Lenovo.asp">LENOVO</a>
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
    <div style="margin-left:300px;">
              <p>&nbsp;<A HREF="<%=MM_moveFirst%>" class="colorlink3">&lt;&lt;Trang đầu</A>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<A HREF="<%=MM_movePrev%>" class="colorlink3">&lt;Trước </A>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<A HREF="<%=MM_moveNext%>" class="colorlink3">Tiếp&gt;</A>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<A HREF="<%=MM_moveLast%>" class="colorlink3">Trang cuối&gt;&gt;</A></p>
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
LAPTOP.Close()
Set LAPTOP = Nothing
%>
