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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.DonDatHang (TenKH, DC, SDT, Email, TKKH, NgayDat, TongTien) VALUES (?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 200, Request.Form("txtTenKH")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 200, Request.Form("txtDC")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 20, Request.Form("txtSDT")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("txtEmail")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 50, Request.Form("TKKH")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 135, 1, -1, MM_IIF(Request.Form("NgayDat"), Request.Form("NgayDat"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("TongTien"), Request.Form("TongTien"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "HienThi.asp"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<form id="form1" name="form1" method="POST" action="XuLyGioHang.asp">
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
        <%
		Response.Write(Session("tongtien"))
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
