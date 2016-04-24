<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Session("TKA") = ""
Session("TenAD")=""
Sub Session_OnEnd
End Sub
Response.Redirect("loginAD.asp")
%>