<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "UPDATE dbo.SanPham SET Tinhtrang = ? WHERE MaSP = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("Xoa"), Request.Form("Xoa"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "Products.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim SanPham__MMColParam
SanPham__MMColParam = "1"
If (Request.Form("MaSP") <> "") Then 
  SanPham__MMColParam = Request.Form("MaSP")
End If
%>
<%
Dim SanPham
Dim SanPham_cmd
Dim SanPham_numRows

Set SanPham_cmd = Server.CreateObject ("ADODB.Command")
SanPham_cmd.ActiveConnection = MM_Connection_STRING
SanPham_cmd.CommandText = "SELECT a.MaSP, a.TenSP, a.MaNSX, a.MaLoai, a.HinhAnh, a.Gia, a.Tinhtrang, a.SoLuong,a.CauHinh, b.Loai, c.NSX FROM dbo.SanPham a, dbo.LoaiSP b, dbo.NSX c WHERE MaSP = ? and a.Tinhtrang=1 and a.MaLoai=b.MaLoai and a.MaNSX=c.MaNSX ORDER BY MaSP DESC" 
SanPham_cmd.Prepared = true
SanPham_cmd.Parameters.Append SanPham_cmd.CreateParameter("param1", 5, 1, -1, SanPham__MMColParam) ' adDouble

Set SanPham = SanPham_cmd.Execute
SanPham_numRows = 0
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
if Session("name") = "" then
	Response.Redirect("loginAD.asp")
else
	Response.write("Xin chào, <b class=tentk>" & Session("name") & "</b><b class=to> |</b>" & "<a href=logoutAD.asp class=colorlink2><ins>Thoát</ins></a>")
	
end if
%>
</div>
<div id="top"><p class=title align=center>QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
<h1 style="color:rgb(0, 66, 255)" size="300%" align="center">Xóa sản phẩm</h1>
<div class="product-item">
  <table align="center" width="50%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="5%"><img src="<%=(SanPham.Fields.Item("HinhAnh").Value)%>" alt="" width="225" height="150"></td>
      <td width="25%"><p>- <%=(SanPham.Fields.Item("TenSP").Value)%></p>
        <p>- Loại <%=(SanPham.Fields.Item("Loai").Value)%></p>
        <p>- Hãng <%=(SanPham.Fields.Item("NSX").Value)%></p>
        <p>- Cấu hình <%=(SanPham.Fields.Item("CauHinh").Value)%></p>
        <p>- Số lượng <%=(SanPham.Fields.Item("SoLuong").Value)%></p></td>
      <td width="5%"><p>&nbsp;</p>
		<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
          <button type="submit" name="button" id="button" value="XÓA">XÓA</button>
          <input name="Xoa" type="hidden" id="Xoa" value="0">
          <input type="hidden" name="MM_update" value="form1">
          <input type="hidden" name="MM_recordId" value="<%= SanPham.Fields.Item("MaSP").Value %>">
        </form>
        </p>
        <form name="form2" method="post" action="Products.asp">
          <Button type="submit" name="button2" id="button2" value="HỦY">HỦY</button>
        </form>
        <p>&nbsp;</p></td>
    </tr>
  </table>
</div>
<div class="clear"></div>
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
SanPham.Close()
Set SanPham = Nothing
%>
