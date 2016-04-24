<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Connection.asp" -->
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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "UPDATE dbo.KhachHang SET MatKhau = ? WHERE TKKH = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("txtPN")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "SuaTTCN.asp"
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
Dim KhachHang__MMColParam
KhachHang__MMColParam = "1"
If (Session("TKKH") <> "") Then 
  KhachHang__MMColParam = Session("TKKH")
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Lato:400,300,600,700,800' rel='stylesheet' type='text/css'>
<link href="css/login.css" rel="stylesheet" type="text/css" media="all" />
	<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.validate.min.js"></script>
	<script type="text/javascript" src="js/messages_vi.js"></script>
    <script type="text/javascript">
		$(document).ready(function(){
			$("#contact").validate({
				errorElement: "span", // Định dạng cho thẻ HTML hiện thông báo lỗi
				rules: {
					reNP: {
						equalTo: "#password" // So sánh trường cpassword với trường có id là password
					},
					min_field: { min: 5 }, //Giá trị tối thiểu
					max_field: { max : 10 }, //Giá trị tối đa
					range_field: { range: [4,10] }, //Giá trị trong khoảng từ 4 - 10
					rangelength_field: { rangelength: [4,10] } //Chiều dài chuỗi trong khoảng từ 4 - 10 ký tự
				}
			});
		});
	</script>
<title>Thông Tin Tài Khoản</title>
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
</head>

<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("TKKH") = "" then
	Response.redirect("index.asp")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=logout.asp class=colorlink2 <ins>Thoát<ins></a>")
	Response.write("<div><p align=right class=thongtincanhan><a href=index.asp rel=nofollow class=colorlink><span><ins>Trở về Trang Chủ</ins></span></a></p></div>")
	Noti = Session("noti")
	Session("noti")=""
end if
%>
	</div>
    <div class="pages-top">
    	<div class="logo">
        <a href="index.asp"><img src="images/logo.png" alt=""/></a>
    	</div>
    	<div class="clear"></div>
	</div>
    <div>
    <table id="tb-accout">
    <style type="text/css">
          #tb-account { width:100%; border-collapse:collapse}
          #danhmuctrai { width:220px; vertical-align:top; padding-right:30px;}
          h2,h3{margin-top:0;}
          #danhmuctrai dt {
            font-weight: bold;
            line-height: 26px;
            background:#D11A1A;
            color: white;
            padding: 0 8px;
            }
          #danhmuctrai dd, dt, dl{ margin:0; padding:0}
          #danhmuctrai dl { margin-bottom:10px}
          #danhmuctrai a {
            color: #00FFFF;
            text-decoration: none;
            line-height: 30px;
            }
      </style>
    	<tbody>
    		<tr>
 				<td id="danhmuctrai">
            		<dl>
                	<dt>Thông Tin Tài Khoản</dt>
                	<dd>
                	<a href="SuaTTCN.asp?TTCN=EditTTCN">Thông Tin Cá Nhân</a>
                	</dd>
                	<dd>
                	<a href="SuaTTCN.asp?TTCN=ChangeOP">Thay Đổi Mật Khẩu</a>
                	</dd>
                	</dl>
                	<dl><dt>Đơn Đặt Hàng Mua</dt>
                	<dd>
                	<a href=?view=accout-order>Danh Sách Đơn Hàng</a>
                	</dd>
                	</dl>
            	</td>
                <td valign="top">
                    

	
		<div>
          <form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
		  <table width="100%" border="1" cellspacing="0" cellpadding="10px">
		    <tr>
		      <td>A</td>
		      <td><label for="txtPN"></label>
		        <input type="password" name="txtPN" id="txtPN" /></td>
		      </tr>
		    <tr>
		      <td>B</td>
		      <td><label for="txtRePN"></label><input type="password" name="txtRePN" id="txtRePN" /></td>
		      </tr>
		    </table>
          <input type="hidden" name="MM_update" value="form1" />
          <input type="hidden" name="MM_recordId" value="<%= KhachHang.Fields.Item("TKKH").Value %>" />
          <input type="submit" name="button" id="button" value="Đổi MK" />
          </form>
        </div>
	
                  <p>Bạn muốn chỉnh sửa click  <a href=thaydoiTT.asp class="colorlink"><ins>tại đây</ins></a></p>
                </td>  
        	</tr>
        </tbody>
     </table>
     </div>

</body>
</html>
<%
KhachHang.Close()
Set KhachHang = Nothing
%>
