<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Connection.asp" -->
<%
Dim MM_editAction
Dim M_TinhTrang
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
If (CStr(Request("MM_update")) = "form2") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "UPDATE dbo.DonDatHang SET TinhTrang = ? WHERE MaDDH = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 20, Request.Form("TinhTrang")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "DDH.asp"
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
Dim DDH__MMColParam
DDH__MMColParam = "1"
If (Request.Form("MaDDH") <> "") Then 
  DDH__MMColParam = Request.Form("MaDDH")
End If
%>
<%
Dim DDH
Dim DDH_cmd
Dim DDH_numRows

Set DDH_cmd = Server.CreateObject ("ADODB.Command")
DDH_cmd.ActiveConnection = MM_Connection_STRING
DDH_cmd.CommandText = "SELECT * FROM dbo.DonDatHang WHERE MaDDH = ? ORDER BY MaDDH DESC" 
DDH_cmd.Prepared = true
DDH_cmd.Parameters.Append DDH_cmd.CreateParameter("param1", 5, 1, -1, DDH__MMColParam) ' adDouble

Set DDH = DDH_cmd.Execute
DDH_numRows = 0
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
    <p>
      <script src="js/vendor/jquery-1.10.1.min.js"></script>
      <script src="js/plugins.js"></script>
      <script src="js/main.js"></script>
    </p>
<h1 style="color:rgb(0, 66, 255)" size="300%" align="center">CẬP NHẬT ĐƠN ĐẶT HÀNG</h1>
<% 
	M_TinhTrang=(DDH.Fields.Item("TinhTrang").Value)
%>
<div class="midItem" align="center">
    <table width="100%" style="margin-left:10px" border="2px" Bordercolor="black" cellspacing="0" cellpadding="100px">
      <tr>
        <td colspan="2" align="center"><b>ĐƠN ĐẶT HÀNG</b></td>
      </tr>
      <tr>
        <td width="50%">    Mã đơn đặt hàng</td>
        <td width="50%">    <b><%=(DDH.Fields.Item("MaDDH").Value)%></b></td>
      </tr>
      <tr>
        <td>    Tài khoản khách hàng</td>
        <td>    <%=(DDH.Fields.Item("TKKH").Value)%></td>
      </tr>
      <tr>
        <td>    Ngày đặt</td>
        <td>    <%=(DDH.Fields.Item("NgayDat").Value)%></td>
      </tr>
      <tr>
        <td>    Tổng tiền</td>
        <td>    <%=(DDH.Fields.Item("TongTien").Value)%></td>
      </tr>
      <tr>
        <td>    Tình trạng</td>
        <td>    <%=(DDH.Fields.Item("TinhTrang").Value)%></td>
      </tr>
      <%
	  Content = ""
	  if M_TinhTrang ="Chưa thanh toán     " then
	  Content = Content & ""
	  else
	  Content = Content & "<tr><td colspan=2 align=center><form name=form2 method=POST action=" & MM_editAction & ">"
	  Content = Content & "<input name=TinhTrang type=hidden id=TinhTrang value='Chưa thanh toán'>"
	  Content = Content & "<button type=submit name=button2 id=button2 value=>CHƯA THANH TOÁN</button>"
	  Content = Content & "<input type=hidden name=MM_update value=form2>"
	  Content = Content & "<input type=hidden name=MM_recordId value="& DDH.Fields.Item("MaDDH").Value & ">"
	  Content = Content & "</form></td></tr>"
	  end if
	  Response.Write(Content)
	  %>
  </table>
    <p>&nbsp; </p>
	<a href="DDH.asp">
  <button type="button" name="button" id="button" value="">
  QUAY LẠI<br/>ĐƠN ĐẶT HÀNG</button>
	</a>
</div>
<div class="footer-bar">
    <span class="article-wrapper">
        <span class="article-label">Trang quản lý</span>
        <span class="article-link"><a href="#">Lên <ins>TOP▲</ins></a></span>
    </span>
</div>
</body>
</html>
<%
DDH.Close()
Set DDH = Nothing
%>
