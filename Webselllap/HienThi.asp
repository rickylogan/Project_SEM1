<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Connection.asp" -->
<%
Dim KhachHang
Dim KhachHang_cmd
Dim KhachHang_numRows

Set KhachHang_cmd = Server.CreateObject ("ADODB.Command")
KhachHang_cmd.ActiveConnection = MM_Connection_STRING
KhachHang_cmd.CommandText = "SELECT * FROM dbo.KhachHang" 
KhachHang_cmd.Prepared = true

Set KhachHang = KhachHang_cmd.Execute
KhachHang_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="css/myStyle.css" />
<title>Untitled Document</title>

</head>

<body>
<h2>Giỏ hàng của Quý khách có <% Response.Write(Session("dem")) %> sản phẩm</h2>
<%
		Content = ""
		Dim dem
		dem = Session("dem")
		if(dem <> 0) then
		Dim q
		if(Session("dem") <> "") then
			Response.Write("<table border = 1 width = '100%'>")
			dem = Session("dem")
			q = Session("giohang")
			Response.Write("<tr>")
					Response.Write("<td align='center'>Mã SP</td>")
					Response.Write("<td align='center'>Tên SP</td>")					
					Response.Write("<td align='center'>Hình ảnh</td>")'
					Response.Write("<td align='center'>Số lượng</td>")										
					Response.Write("<td align='center'>Giá</td>")
					Response.Write("<td align='center'>Tùy chọn</td>")					
				Response.Write("</tr>")
			for i = 0 to dem - 1
				Response.Write("<tr>")
					Response.Write("<td align='center'>"&q(i,0)&"</td>")
					Response.Write("<td align='center'>"&q(i,1)&"</td>")					
					Response.Write("<td align='center'><img src='"&q(i,2)&"' width=50 height=50/></td>")'
					Response.Write("<td align='center'>"&q(i,3)&"</td>")										
					Response.Write("<td align='right'>"&q(i,4)&"</td>")
					Response.Write("<td align='center'><a href='xoaSPGioHang.asp?"&i&"'><input type='button' value='Xóa'></a></td>")					
				Response.Write("</tr>")
			next	
				Response.Write("<tr>")
				Response.Write("<td colspan=5 align='right'>"&Session("tongtien")&"</td>")
				Response.Write("</tr>")					
			Response.Write("</table>")
			Content = Content & "<input name=btnmuahang type=submit  value='Thanh toán'/>"
		end if
	else
		Response.Write("Quý khách chưa có mặt hàng nào trong giỏ hàng")
	end if
	%>
    <br />
    <a class="link" href="Javascript:history.go(-2)">Tiếp tục mua hàng >>></a>
    <form action="KienTra.asp" method="post">
    <br />
    	<% Response.Write(Content) %>
    	<input name="tkkh" type="hidden" id="tkkh" value="<%=(KhachHang.Fields.Item("TKKH").Value)%>" />
    </form>
</body>
</html>
<%
KhachHang.Close()
Set KhachHang = Nothing
%>
