<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim SanPham
Dim SanPham_cmd
Dim SanPham_numRows

Set SanPham_cmd = Server.CreateObject ("ADODB.Command")
SanPham_cmd.ActiveConnection = MM_Connection_STRING
SanPham_cmd.CommandText = "SELECT a.MaSP, a.TenSP, a.MaNSX, a.MaLoai, a.HinhAnh, a.Gia, a.Tinhtrang, a.SoLuong, b.Loai, c.NSX FROM dbo.SanPham a, dbo.LoaiSP b, dbo.NSX c WHERE a.Tinhtrang=1 and a.MaLoai=b.MaLoai and a.MaNSX=c.MaNSX ORDER BY MaSP DESC" 
SanPham_cmd.Prepared = true

Set SanPham = SanPham_cmd.Execute
SanPham_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Dim Num_page
Dim Page
Page = 0

Repeat1__numRows = 5
Num_page = Repeat1__numRows + 0
Repeat1__index = 0
SanPham_numRows = SanPham_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 5
Repeat2__index = 0
SanPham_numRows = SanPham_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim SanPham_total
Dim SanPham_first
Dim SanPham_last

' set the record count
SanPham_total = SanPham.RecordCount

' set the number of rows displayed on this page
If (SanPham_numRows < 0) Then
  SanPham_numRows = SanPham_total
Elseif (SanPham_numRows = 0) Then
  SanPham_numRows = 1
End If

' set the first and last displayed record
SanPham_first = 1
SanPham_last  = SanPham_first + SanPham_numRows - 1

' if we have the correct record count, check the other stats
If (SanPham_total <> -1) Then
  If (SanPham_first > SanPham_total) Then
    SanPham_first = SanPham_total
  End If
  If (SanPham_last > SanPham_total) Then
    SanPham_last = SanPham_total
  End If
  If (SanPham_numRows > SanPham_total) Then
    SanPham_numRows = SanPham_total
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

Set MM_rs    = SanPham
MM_rsCount   = SanPham_total
MM_size      = SanPham_numRows
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
SanPham_first = MM_offset + 1
SanPham_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (SanPham_first > MM_rsCount) Then
    SanPham_first = MM_rsCount
  End If
  If (SanPham_last > MM_rsCount) Then
    SanPham_last = MM_rsCount
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
Dim MM_numPage

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
        <div id="top" class="site-header">
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
<div><p class="title" align="center">QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
	<div align="center" class="form_menu">
            <a href="AddSp.asp" class="colorlink">
            <button type="submit" name=cmdSubmit>Thêm sản phẩm mới</button>
            </a>
    </div>
	<div align="center" class="form_menu">
            <a href="DDH.asp" class="colorlink">
            <button type="button" name=cmdSubmit>Xem đơn đặt hàng</button>
            </a>
    </div>
<div class="product-item">
  <table width="85%" border="0" cellspacing="0" cellpadding="0" align="center">
    <% 
While ((Repeat1__numRows <> 0) AND (NOT SanPham.EOF)) 
%>
  <tr>
    <td width="15%"><p><img src="<%=(SanPham.Fields.Item("HinhAnh").Value)%>" alt="" name="" width="225" height="150"></p></td>
    <td width="35%"><p>&nbsp;</p>
      <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td width="30%">Tên</td>
          <td width="30%"> <%=(SanPham.Fields.Item("TenSP").Value)%></td>
        </tr>
        <tr>
          <td> Loại</td>
          <td><%=(SanPham.Fields.Item("Loai").Value)%></td>
        </tr>
        <tr>
          <td>Hãng</td>
          <td><%=(SanPham.Fields.Item("NSX").Value)%></td>
        </tr>
        <tr>
          <td>Giá</td>
          <td><%=(SanPham.Fields.Item("Gia").Value)%><em> <u>VNĐ</u></em></td>
        </tr>
        <tr>
          <td>Số lượng</td>
          <td><%=(SanPham.Fields.Item("SoLuong").Value)%> sản phẩm</td>
        </tr>
      </table>
      <p></p>
      <p>&nbsp;</p></td>
    <td width="20%"><form action="Editsp.asp" method="post" name="form1" id="form1">
      <input name="MaSP" type="hidden" id="MaSP" value="<%=(SanPham.Fields.Item("MaSP").Value)%>">
      <input name="NSX" type="hidden" id="NSX" value="<%=(SanPham.Fields.Item("MaNSX").Value)%>">
      <input name="Loai" type="hidden" id="Loai" value="<%=(SanPham.Fields.Item("MaLoai").Value)%>">
      <button type="submit" name="button" id="button" value="CẬP NHẬT">CẬP NHẬT</button>
    </form></td>
    <td width="15%"><form action="Removesp.asp" method="post" name="form1" id="form1">
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
  
        <span class="article-label">
        <A HREF="<%=MM_moveFirst%>">
            <button class="paging_left">
                ◄◄&nbsp;&nbsp;<ins>TRANG ĐẦU</ins>
            </button>
        </A>
        <A HREF="<%=MM_movePrev%>">
            <button class="paging_left">
                ◄&nbsp;&nbsp;&nbsp;<ins>TRƯỚC</ins>
            </button>
        </A>
		<A HREF="<%=MM_moveLast%>">
            <button class="paging_right">
                <ins>TRANG CUỐI</ins>&nbsp;&nbsp;►►
            </button>
        </A>
        <A HREF="<%=MM_moveNext%>">
        <button class="paging_right">
                <ins>SAU</ins>&nbsp;&nbsp;&nbsp;►
            </button>
        </A>
        <% 
While ((Repeat2__numRows <> 0) AND (NOT SanPham.EOF)) 
Page=Page+1
MM_numPage   = MM_urlStr & Page * Num_page
%>

  <A HREF="<%=MM_numPage%>">
    <button class="paging_mid"><%=Page%></button>
    </A>
  <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  SanPham.MoveNext()
Wend
%>
	</span>
  </div>
<script src="js/vendor/jquery-1.10.1.min.js"></script>
<script src="js/plugins.js"></script>
<script src="js/main.js"></script>
<div class="footer-bar">
    <span class="article-wrapper">
	<span >
        <span class="article-link"><a href="#" target="_top">Lên <ins>TOP▲</ins></a></span>
    </span>
</div>
</body>
</html>
<%
SanPham.Close()
Set SanPham = Nothing
%>