<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Session("TKKH")=""
Session("dem")=""
Session("name")=""
Session("ConfirmMK")=""
Session("noti")=""
Session("MatKhau")=""
Session("giohang")=""
Session("tongtien")=""
Sub Session_OnEnd
End Sub
Response.Redirect("index.asp")
%>