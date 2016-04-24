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
If (CStr(Request("MM_update")) = "contact") Then
  If (Not MM_abortEdit) Then
    Dim MM_editCmd
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
	if QStr="ChangeNP" then
		MM_editCmd.ActiveConnection = MM_Connection_STRING
		MM_editCmd.CommandText = "UPDATE dbo.KhachHang SET MatKhau = ? WHERE TKKH = ?" 
		MM_editCmd.Prepared = true
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("txtNP")) ' adLongVarChar
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
	else
		Set MM_editCmd = Server.CreateObject ("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_Connection_STRING
		MM_editCmd.CommandText = "UPDATE dbo.KhachHang SET TenKH = ?, DiaChi = ?, SDT = ? WHERE TKKH = ?" 
		MM_editCmd.Prepared = true
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("txtFullname")) ' adVarWChar
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 100, Request.Form("txtAddress")) ' adVarWChar
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("txtPhone")) ' adLongVarChar
		MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
	end if
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
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
						equalTo: "#txtNP" // So sánh trường cpassword với trường có id là password
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
</head>

<body>
<div class="wrap"> 
    <div class="gocphaimanhinhTV">
<%
if Session("TKKH") = "" then
	Response.redirect("index.asp")
else
	Response.write("Xin chào " & Session("name") & "," & "&nbsp;" & "<a href=logout.asp class=colorlink2><ins>Thoát</ins></a>")
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
                	<a href="SuaTTCN.asp?TTCN=TTCN">Thông Tin Cá Nhân</a>
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
        <%
	Content = ""							
	QStr = Request.QueryString("TTCN")		

	if ucase(left(QStr,6))="CHANGE" then 
		Title = "Đổi mật khẩu"
	else
		Title = "Thông tin cá nhân"
	end if
	if QStr="" then
		Content = Content & "<div><table style='border-collapse:collapse; width:100%' border=1 bordercolor=#FF0000 cellpadding=5>"
		Content = Content & "<tbody style=color:#FFF><tr>"
		Content = Content & "<td width=120px>Họ và Tên </td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("TenKH").Value & "</td>"
		Content = Content & "</tr><tr><td>Địa Chỉ Email</td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("Email").Value & "</td>"
		Content = Content & "</tr><tr><td>Địa Chỉ Nhà </td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("DiaChi").Value & "</td>"
		Content = Content & "</tr><tr><td>Số Điện Thoại </td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("SDT").Value & "</td>"
		Content = Content & "</tr></tbody></table></div>"
	elseif QStr="EditTTCN" then
		Content = Content & "<div><table style='border-collapse:collapse; width:100%' border=1 bordercolor=#FF0000 cellpadding=5><tr>"
		Content = Content & "<form name=contact method=POST action=" & MM_editAction & ">"
		Content = Content & "<td width=120px>Họ và Tên </td><td><input type=text name=txtFullname value='" & KhachHang.Fields.Item("TenKH").Value & "' style=margin:0px required/></td></tr>"
		Content = Content & "<tr><td>Địa Chỉ Email </td><td><input type=Email name=txtEmail readonly=readonly value='" & KhachHang.Fields.Item("Email").Value & "' style=margin:0px required/></td></td></tr>"
		Content = Content & "<tr><td>Địa Chỉ Nhà </td><td><input type=text name=txtAddress value='" & KhachHang.Fields.Item("Diachi").Value & "' style=margin:0px required/></td></td></tr>"
		Content = Content & "<tr><td>Số Điện Thoại</td><td><input type=tell name=txtPhone value=" & KhachHang.Fields.Item("SDT").Value & " style=margin:0px required/></td></td></tr>"
		Content = Content & "</tbody></table><button type=submit name=cmdsubmit value=>Cập Nhật</button>"
		Content = Content & "<input type=hidden name=MM_update value=contact /><input type=hidden name=MM_recordId value=" & KhachHang.Fields.Item("TKKH").Value & " /></form></div>"
	elseif QStr="ChangeOP" then
		Content = Content & "<div><form ACTION=verify_dmk.asp name=form1 METHOD=POST Style=margin:0px>"
		Content = Content & "<table style='border-collapse:collapse; width:100%' border=1 bordercolor=#FF0000 cellpadding=5>"
		Content = Content & "<tr><td>Mật khẩu cũ</td>"
		Content = Content & "<td><input type=password  name=txtOP style=margin:0px required/></td>"
		Content = Content & "</tr><input type=hidden name=MM_update value=form1 /><input type=hidden name=MM_recordId />"
		Content = Content & "</table><button style='margin-left:50px' type=submit name=cmdsubmit value=>Đồng Ý</button></form></div>"
	elseif QStr="ChangeNP" then
		if	Session("ConfirmMK") = "" then
			Response.redirect("SuaTTCN.asp?TTCN=ChangeOP")
		elseif Session("ConfirmMK") = "Confirmed" then
			Session("ConfirmMK") = ""
			Content = Content & "<div><form ACTION=" & MM_editAction & "completed method=POST name=contact id=contact Style=margin:0px>"
			Content = Content & "<table style='border-collapse:collapse; width:100%' border=1 bordercolor=#FF0000 cellpadding=5>"
			Content = Content & "<tr><td>Mật khẩu mới</td>"
			Content = Content & "<td><input minlength=6 id=txtNP name=txtNP type=password required /></td>"
			Content = Content & "<tr><td>Nhập lại mật khẩu</td>"
			Content = Content & "<td><input name=reNP type=password tabindex=3 required /></td></tr>"
			Content = Content & "</table><button style=margin-left:50px type=submit name=cmdsubmit value=>Đồng Ý</button>"
			Content = Content & "<input type=hidden name=MM_update value=contact>"
			Content = Content & "<input type=hidden name=MM_recordId value=" & KhachHang.Fields.Item("TKKH").Value & "></form></div>"
			Session("Noti") = "Completed"
		end if
	elseif QStr="ChangeNPcompleted" then
		if	Session("Noti") = "" then
			Response.redirect("SuaTTCN.asp")
		elseif Session("Noti") = "Completed" then
			Content = Content & "Bạn đã đổi mật khẩu thành công"
			Session("Noti") = ""
		end if
	else
		Content = Content & "<div><table style='border-collapse:collapse; width:100%' border=1 bordercolor=#FF0000 cellpadding=5>"
		Content = Content & "<tbody style=color:#FFF><tr>"
		Content = Content & "<td width=120px>Họ và Tên</td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("TenKH").Value & "</td>"
		Content = Content & "</tr><tr><td>Địa Chỉ Email</td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("Email").Value & "</td>"
		Content = Content & "</tr><tr><td>Địa Chỉ Nhà</td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("DiaChi").Value & "</td>"
		Content = Content & "</tr><tr><td>Số Điện Thoại</td>"
		Content = Content & "<td>" & KhachHang.Fields.Item("SDT").Value & "</td>"
		Content = Content & "</tr></tbody></table></div>"		
	end if
		%>
		<%
				Response.Write("<h2 style='color:#F00; background-color:#00F; font-size:20pt'>" & Title & "</h2>")
				Response.Write("<h2 style='color:#F00; background-color:#00F; text-align:center'>" & Noti & "</h2>")
				Response.Write(Content)
		%>
        </div>
        
                  <p>Bạn muốn chỉnh sửa click  <a href=SuaTTCN.asp?TTCN=EditTTCN class="colorlink"><ins>tại đây</ins></a></p>
                </td>  
        	</tr>
        </tbody>
     </table>
</div>
</div>
</body>
</html>
<%
KhachHang.Close()
Set KhachHang = Nothing
%>
