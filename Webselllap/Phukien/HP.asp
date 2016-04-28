<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim PHUKIEN_HEADPHONES
Dim PHUKIEN_HEADPHONES_cmd
Dim PHUKIEN_HEADPHONES_numRows

Set PHUKIEN_HEADPHONES_cmd = Server.CreateObject ("ADODB.Command")
PHUKIEN_HEADPHONES_cmd.ActiveConnection = MM_Connection_STRING
PHUKIEN_HEADPHONES_cmd.CommandText = "SELECT * FROM dbo.SanPham WHERE Tinhtrang = 1 and MaLoai = 4  and TenSP like '%headphone%' or  TenSP like '%tai nghe%'" 
PHUKIEN_HEADPHONES_cmd.Prepared = true

Set PHUKIEN_HEADPHONES = PHUKIEN_HEADPHONES_cmd.Execute
PHUKIEN_HEADPHONES_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 9
Repeat1__index = 0
PHUKIEN_HEADPHONES_numRows = PHUKIEN_HEADPHONES_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim PHUKIEN_HEADPHONES_total
Dim PHUKIEN_HEADPHONES_first
Dim PHUKIEN_HEADPHONES_last

' set the record count
PHUKIEN_HEADPHONES_total = PHUKIEN_HEADPHONES.RecordCount

' set the number of rows displayed on this page
If (PHUKIEN_HEADPHONES_numRows < 0) Then
  PHUKIEN_HEADPHONES_numRows = PHUKIEN_HEADPHONES_total
Elseif (PHUKIEN_HEADPHONES_numRows = 0) Then
  PHUKIEN_HEADPHONES_numRows = 1
End If

' set the first and last displayed record
PHUKIEN_HEADPHONES_first = 1
PHUKIEN_HEADPHONES_last  = PHUKIEN_HEADPHONES_first + PHUKIEN_HEADPHONES_numRows - 1

' if we have the correct record count, check the other stats
If (PHUKIEN_HEADPHONES_total <> -1) Then
  If (PHUKIEN_HEADPHONES_first > PHUKIEN_HEADPHONES_total) Then
    PHUKIEN_HEADPHONES_first = PHUKIEN_HEADPHONES_total
  End If
  If (PHUKIEN_HEADPHONES_last > PHUKIEN_HEADPHONES_total) Then
    PHUKIEN_HEADPHONES_last = PHUKIEN_HEADPHONES_total
  End If
  If (PHUKIEN_HEADPHONES_numRows > PHUKIEN_HEADPHONES_total) Then
    PHUKIEN_HEADPHONES_numRows = PHUKIEN_HEADPHONES_total
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

Set MM_rs    = PHUKIEN_HEADPHONES
MM_rsCount   = PHUKIEN_HEADPHONES_total
MM_size      = PHUKIEN_HEADPHONES_numRows
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
PHUKIEN_HEADPHONES_first = MM_offset + 1
PHUKIEN_HEADPHONES_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (PHUKIEN_HEADPHONES_first > MM_rsCount) Then
    PHUKIEN_HEADPHONES_first = MM_rsCount
  End If
  If (PHUKIEN_HEADPHONES_last > MM_rsCount) Then
    PHUKIEN_HEADPHONES_last = MM_rsCount
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
if Session("name") = "" then
	Response.write("<a rel=nofollow href=../login.asp?login=createnew class=colorlink2><span><ins>Đăng ký</ins></span></a>|<a rel=nofollow href=../login.asp class=colorlink2><span><ins>Đăng Nhập</ins></span></a>")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=../logout.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
	<div class="giohang">
    	<a href="../HienThi.asp"><img width="50" height="50" src="../Images/giohang_index.png" /></a>
        <p class="soluong"> <% Response.Write(Session("dem")) %></p>
    </div>
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
					<li><a href="../Laptop/Laptop.asp">Laptop</a>
						<ul>
							<li>
                                <form name="frmDell" method="post" action=laptop/Dell.asp>
                                <a href="../Laptop/Dell.asp">DELL</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmHp" method="post" action=laptop/Hp.asp>
                                <a href="../Laptop/Hp.asp">HP</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmApple" method="post" action=laptop/Apple.asp>
                                <a href="../Laptop/Apple.asp">APPLE</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAsus" method="post" action=laptop/sus.asp>
                                <a href="../Laptop/Asus.asp">ASUS</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmAcer" method="post" action=laptop/Acer.asp>
                                <a href="../Laptop/Acer.asp">ACER</a>
                                </form>
                            </li>
                            <li>
                                <form name="Lenovo" method="post" action=laptop/Lenovo.asp>
                                <a href="../Laptop/Lenovo.asp">LENOVO</a>
                                </form>
                            </li>
                            <li>
                                <form name="frmVaio" method="post" action=laptop/Vaio.asp>
                                <a href="../Laptop/Vaio.asp">VAIO</a>
                                </form>
                            </li>
						</ul>
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
					<li class="active"><a href="../Phukien/Phukien.asp">Phụ kiện</a>
						
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
				  <p>CÁC SẢN PHẨM PHUKIEN HEADPHONES MỚI NHẤT</p>
				  <p>&nbsp;</p>
                  <% 
While ((Repeat1__numRows <> 0) AND (NOT PHUKIEN_HEADPHONES.EOF)) 
%>
  <div class="oneItem">
    <p><img src="<%=(PHUKIEN_HEADPHONES.Fields.Item("HinhAnh").Value)%>" alt="" width="225" height="150">Sản phẩm: <%=(PHUKIEN_HEADPHONES.Fields.Item("TenSP").Value)%></p>
    <p>Giá: <%=(PHUKIEN_HEADPHONES.Fields.Item("Gia").Value)%></p>
    <p>Hiện còn <%=(PHUKIEN_HEADPHONES.Fields.Item("SoLuong").Value)%> sản phẩm</p>
    <p>&nbsp;</p>
    <form name="form1" method="post" action="ctspPhukien.asp">
      <input name="MaSP" type="hidden" id="MaSP" value="<%=(PHUKIEN_HEADPHONES.Fields.Item("MaSP").Value)%>">
      <label>
        <input type="submit" name="btnchitiet" id="btnchitiet" value="Xem chi tiết sản phẩm...">
      </label>
    </form>
    
  </div>
  
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  PHUKIEN_HEADPHONES.MoveNext()
Wend
%>
             <div style="margin-top:10px; margin-bottom:10px; float:left">
              <p>&nbsp;<A HREF="<%=MM_moveFirst%>">&lt;&lt;Trang đầu </A><A HREF="<%=MM_movePrev%>">&lt;&lt;Trước </A>**** <A HREF="<%=MM_moveNext%>">Tiếp&gt;&gt;</A> <A HREF="<%=MM_moveLast%>">Trang cuối&gt;&gt;</A></p>
              </div>

              </div>
              
  
 </div>
       </div>
               
              
<!-- END gallary-->
        <div class="labout span_1_of_g1">
		  <div class="project-list">
	     	<h4>Loại</h4>
			<ul class="blog-list">
				<li>
					<form name="frmHP" method="post" action=HP.asp>
                    <a href="HP.asp">HEADPHONES</a>
                    </form>
                </li>
                <li>
					<form name="frmEP" method="post" action=EP.asp>
                    <a href="EP.asp">EARPHONES</a>
                    </form>
                </li>
                <li>
					<form name="frmCQ" method="post" action=CQ.asp>
                    <a href="CQ.asp">CHUỘT</a>
                    </form>
                </li><li>
					<form name="frmBP" method="post" action=BP.asp>
                    <a href="BP.asp">BÀN PHÍM</a>
                    </form>
                </li><li>
					<form name="frmUSB" method="post" action=USB.asp>
                    <a href="USB.asp">USB</a>
                    </form>
                </li><li>
					<form name="frmLC" method="post" action=LC.asp>
                    <a href="LC.asp">LÓT CHUỘT</a>
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
PHUKIEN_HEADPHONES.Close()
Set PHUKIEN_HEADPHONES = Nothing
%>
