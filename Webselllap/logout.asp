<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Session("TKKH")=""
Session("dem")="0"
Session("name")=""
Session("ConfirmMK")=""
Session("noti")=""
Session("MatKhau")=""
Session("giohang")=""
Session("tongtien")="0"
Session("NotiTT")=""
Session("NotiNP")=""
Sub Session_OnEnd
End Sub
Response.Redirect("index.asp")
%>