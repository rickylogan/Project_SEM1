<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
	if(Session("TKKH") = "") then
		Response.Redirect("login.asp?login=createnew")
	else
		Response.Redirect("ThongTinKHDangNhap.asp")
	end if
%>
  <p>
    <label>
      <input type="submit" name="button" id="button" value="Đặt hàng" />
    </label>
    <input name="hiddenField" type="hidden" id="hiddenField" value="<%=(KhachHang.Fields.Item("TKKH").Value)%>" />
  </p>
</form>
	end if
%>
<div align="center">
        <%
        Response.Write("<p class=title><td align=center><b>" & Title & "</p>")
        Response.Write(Content)
        %>
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
