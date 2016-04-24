<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Connection.asp" -->
<%
Dim KhachHang__MMColParam
KhachHang__MMColParam = "1"
If (Session("TKKH") <> "") Then 
  KhachHang__MMColParam = Session("TKKH")
End If
%>
<%
Dim KhachHang
Dim KhachHang_cmd
Dim KhachHang_numRows

Set KhachHang_cmd = Server.CreateObject ("ADODB.Command")
KhachHang_cmd.ActiveConnection = MM_Connection_STRING
KhachHang_cmd.CommandText = "SELECT TKKH, TenKH, DiaChi, Email, SDT FROM dbo.KhachHang WHERE TKKH = ?" 
KhachHang_cmd.Prepared = true
KhachHang_cmd.Parameters.Append KhachHang_cmd.CreateParameter("param1", 200, 1, 50, KhachHang__MMColParam) ' adVarChar

Set KhachHang = KhachHang_cmd.Execute
KhachHang_numRows = 0
%>
<%
Dim DDH
Dim DDH_cmd
Dim DDH_numRows

Set DDH_cmd = Server.CreateObject ("ADODB.Command")
DDH_cmd.ActiveConnection = MM_Connection_STRING
DDH_cmd.CommandText = "SELECT * FROM dbo.DonDatHang" 
DDH_cmd.Prepared = true

Set DDH = DDH_cmd.Execute
DDH_numRows = 0
%>
<%
	if(Session("name") = "") then
		Response.Redirect("login.asp")
	end if
%>
<head>
<title>Thông tin khách hàng | Cửa hàng máy tính  :: Groupfour</title>
<link rel="shortcut icon" href="images/icon.png">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Lato:400,300,600,700,800' rel='stylesheet' type='text/css'>

<style>HTML,BODY{cursor: url("images/monkeyani.cur"), url("images/monkey-ani.gif"), auto;}</style>
</head>
<body>
<div align="center">
    <form id="form1" name="form1" action="XuLyGioHang.asp">
      <table width="654" height="221" border="1">
        <tr>
          <td width="240"><strong>Tên khách hàng</strong></td>
          <td width="398"><%=(KhachHang.Fields.Item("TenKH").Value)%></td>
        </tr>
        <tr>
          <td><strong>Địa chỉ</strong></td>
          <td><%=(KhachHang.Fields.Item("DiaChi").Value)%></td>
        </tr>
        <tr>
          <td><strong>SĐT</strong></td>
          <td><%=(KhachHang.Fields.Item("SDT").Value)%></td>
        </tr>
        <tr>
          <td><strong>Email</strong></td>
          <td><%=(KhachHang.Fields.Item("Email").Value)%></td>
        </tr>
        <tr>
          <td><strong>Tổng số tiền</strong></td>
          <td><label style="color:red;font-weight:bold;">
            <%
			TongTien = left(right(FormatCurrency(Session("tongtien")),13),10)
			Response.Write(TongTien)
			%>
            <em><u> VNĐ</u></em></label></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p>
        <label>
          <button type="submit" name="button" id="button" value="" />Đặt hàng</button>
        </label>
      </p>
    </form>
</div>
</body>
</html>
<%
KhachHang.Close()
Set KhachHang = Nothing
%>
<%
DDH.Close()
Set DDH = Nothing
%>
