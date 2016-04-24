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
If (CStr(Request("MM_insert")) = "KHMoi") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.KHMoi (TenKH, DiaChi, SDT, Email) VALUES (?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 200, Request.Form("txtTenKHMoi")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 100, Request.Form("txtDCKHMoi")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("txtSDTKHMoi")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 100, Request.Form("txtEmailKHMoi")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "XuLyGioHang.asp"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" id="KHMoi" name="KHMoi">
  <table width="532" border="1">
    <tr>
      <td width="156">Tên</td>
      <td width="360"><label>
        <input type="text" name="txtTenKHMoi" id="txtTenKHMoi" />
      </label></td>
    </tr>
    <tr>
      <td>Địa chỉ</td>
      <td><label>
        <input type="text" name="txtDCKHMoi" id="txtDCKHMoi" />
      </label></td>
    </tr>
    <tr>
      <td>SĐT</td>
      <td><label>
        <input type="text" name="txtSDTKHMoi" id="txtSDTKHMoi" />
      </label></td>
    </tr>
    <tr>
      <td>Email</td>
      <td><label>
        <input type="email" name="txtEmailKHMoi" id="txtEmailKHMoi" />
      </label></td>
    </tr>
    <tr>
      <td>Tổng tiền</td>
      <td><%Response.Write(Session("tongtien"))%></td>
    </tr>
    <tr>
      <td><label>
        <input type="submit" name="btnMuaHang" id="btnMuaHang" value="Xác nhận" />
      </label></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="KHMoi" />
</form>

</body>
</html>
