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
KhachHang_cmd.CommandText = "SELECT * FROM dbo.KhachHang WHERE TKKH = ?" 
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<form id="form1" name="form1" action="XuLyGioHang.asp">
  <table width="654" height="221" border="1">
    <tr>
      <td width="240">Tên</td>
      <td width="398"><%=(KhachHang.Fields.Item("TenKH").Value)%></td>
    </tr>
    <tr>
      <td>Địa chỉ</td>
      <td><%=(KhachHang.Fields.Item("DiaChi").Value)%></td>
    </tr>
    <tr>
      <td>SĐT</td>
      <td><%=(KhachHang.Fields.Item("SDT").Value)%></td>
    </tr>
    <tr>
      <td>Email</td>
      <td><%=(KhachHang.Fields.Item("Email").Value)%></td>
    </tr>
    <tr>
      <td>Tổng tiền</td>
      <td><label>
        <%	x = Session("tongtien")
		if len(x) mod 3 = 0 then
		  			Response.Write(right(left(FormatCurrency(x),4*len(x)\3),- 1 + 4*len(x)\3)&"<em> <u>VNĐ</u></em>")
		  		else Response.Write(right(left(FormatCurrency(x),1+ 4*len(x)\3),0+ 4*len(x)\3)&"<em> <u>VNĐ</u></em>")
				end if
		%>
      </label></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>
    <label>
      <input type="submit" name="button" id="button" value="Đặt hàng" />
    </label>
  </p>
</form>
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
