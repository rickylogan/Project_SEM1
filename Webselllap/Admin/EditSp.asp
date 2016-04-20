<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
'*********************************
'*                               *
'*  UPDATE RECORD AND UPLOAD     *
'*  http://www.dwzone.it         *
'*                               *
'*********************************
	server.ScriptTimeout = 5400	

	Dim RG_Connection, RG_editColumn, RG_recordId, Form, editAction, editRedirectUrl, RG_Files, RG_formName, UploadType, ParamVal, ParamList, MaxFieldNumber, TmpVal, x, y, Key, ProgressBar, UploadStatus, NumFile
	Dim RG_altVal, RG_columns, RG_Cong, RG_dbValues, RG_dbValuesTmp, RG_delim, RG_editCmd, RG_editQuery, RG_editQueryTmp, RG_emptyVal, RG_Ext, RG_Extensions, RG_fields, RG_FieldValueTmp, RG_FileDel, RG_FileExt, RG_formVal, RG_FS, RG_i, RG_L, RG_Len, RG_Max, RG_Name, RG_New, RG_newName, RG_Num, RG_Path, RG_Rec, RG_ret, RG_Save, RG_tableValues, RG_tableValuesTmp, RG_tst, RG_typeArray, RG_z, TotalFileSize, valueToRedirectSend
	Set Form = New ASPForm
	Dim UploadID
	UploadID = Form.NewUploadID
	ProgressBar = ""
	TotalFileSize = ""
	editRedirectUrl = "Products.asp"
	RG_Connection = MM_Connection_STRING
	RG_editTable = "dbo.SanPham"
	RG_editColumn = "MaSP"
	RG_Files = "/Images;1;;;;0;HinhAnh;1;;2;;0;;;;;;HinhAnh@_@_@0@_@_@ @_@_@@_@_@../"
	RG_formName = "form1"
	UploadType="Update"
	UploadStatus = ""
	valueToRedirectSend = ""
	NumFile = 0

	if len(Request.QueryString("UploadID"))>0 then
		Form.UploadID = Request.QueryString("UploadID")
	end if
	
	if (Request.QueryString <> "") Then
	 	editAction = CStr(Request.ServerVariables("SCRIPT_NAME")) & "?" & Request.QueryString & "&UploadID=" & UploadID
	else
	 	editAction = CStr(Request.ServerVariables("SCRIPT_NAME")) & "?UploadID=" & UploadID	
	End If
	
Const fsCompletted  = 0

If Form.State = fsCompletted Then 
  if Form.State = 0 then

		Set ParamVal = CreateObject("Scripting.Dictionary")
		tmp = split(RG_Files,"@_@_@")
		ParamList = split(tmp(0),"|")
		MaxFieldNumber = ubound(ParamList)
		for x=0 to Ubound(ParamList)
			TmpVal = Split(ParamList(x),";")
			for y=0 to ubound(TmpVal)
				Key = right("00" & cstr(x),3) & cstr(y)
				ParamVal.add Key, TmpVal(y)
			next
		next
		Form.Files.Save
		RG_fieldsStr  = "textTenSP|value|Loai|value|NSX|value|txtCauHinh|value|txtSoLuong|value|txtGia|value|txtGhiChu|value"
  		RG_columnsStr = "TenSP|',none,''|MaLoai|none,none,NULL|MaNSX|none,none,NULL|CauHinh|',none,''|SoLuong|none,none,NULL|Gia|none,none,NULL|GhiChu|',none,''"
		Form.Files.DataBaseUpdate

		response.write(getRedirect())
		response.end
  End If
ElseIf Form.State > 10 then
  response.write ""
End If


function GetFolderName(str):  GetFolderName = Ris : end function

function myGetFileName(str):  myGetFileName = Ris : end function
%>
<%
Dim Loai
Dim Loai_cmd
Dim Loai_numRows

Set Loai_cmd = Server.CreateObject ("ADODB.Command")
Loai_cmd.ActiveConnection = MM_Connection_STRING
Loai_cmd.CommandText = "SELECT * FROM dbo.LoaiSP ORDER BY MaLoai ASC" 
Loai_cmd.Prepared = true

Set Loai = Loai_cmd.Execute
Loai_numRows = 0
%>
<%
Dim NSX
Dim NSX_cmd
Dim NSX_numRows

Set NSX_cmd = Server.CreateObject ("ADODB.Command")
NSX_cmd.ActiveConnection = MM_Connection_STRING
NSX_cmd.CommandText = "SELECT * FROM dbo.NSX ORDER BY MaNSX ASC" 
NSX_cmd.Prepared = true

Set NSX = NSX_cmd.Execute
NSX_numRows = 0
%>
<%
Dim SanPham__MMColParam
SanPham__MMColParam = "1"
If (Request.Form("MaSP") <> "") Then 
  SanPham__MMColParam = Request.Form("MaSP")
End If
%>
<%
Dim SanPham
Dim SanPham_cmd
Dim SanPham_numRows

Set SanPham_cmd = Server.CreateObject ("ADODB.Command")
SanPham_cmd.ActiveConnection = MM_Connection_STRING
SanPham_cmd.CommandText = "SELECT * FROM dbo.SanPham WHERE MaSP = ?" 
SanPham_cmd.Prepared = true
SanPham_cmd.Parameters.Append SanPham_cmd.CreateParameter("param1", 5, 1, -1, SanPham__MMColParam) ' adDouble

Set SanPham = SanPham_cmd.Execute
SanPham_numRows = 0
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Quản trị viên :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
    <meta name="description" content="">
    <meta name="viewport" content="width=device-width">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href='http://fonts.googleapis.com/css?family=Roboto:400' rel='stylesheet' type='text/css'>
    <link href="http://fonts.googleapis.com/css?family=Open+Sans:300italic,400italic,600italic,700italic,800italic,400,300,600,700,800" rel="stylesheet">
    <link rel="stylesheet" href="css/bootstrap.min.css">
    <link rel="stylesheet" href="css/normalize.min.css">
    <link rel="stylesheet" href="css/animate.css">
    <link rel="stylesheet" href="css/templatemo_misc.css">
    <link rel="stylesheet" href="css/templatemo_style.css">
 <script src="../js/jquery.min.js"></script>
</head>
<body>
        <div class="site-header">
        <div id="templatemo_logo" class="row col-md-4 col-sm-6 col-xs-6">
                            <h1><a href="Products.asp">Admin</a></h1>
            </div>
            <div class="container">
<div class="gocphaimanhinhTV">
<%
if Session("name") = "" then
	Response.Redirect("loginAD.asp")
else
	Response.write("Xin chào, <b class=tentk>" & Session("name") & "</b><b class=to> |</b>" & "<a href=logoutAD.asp class=colorlink2 <ins>Thoát<ins></a>")
	
end if
%>
</div>
<div id="top"><p class=title align=center>QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
    <div id="top"></div>
    <h1 style="color:#0CF" size="300%" align="center">Cập Nhật</h1>
<div class="item">
  <form ACTION="<%=editAction%>" onsubmit="return ProgressBar()" method="post" enctype="multipart/form-data" name="form1">
    <table width="85%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td width="25%">Tên sản phẩm</td>
        <td width="60%"><label for="textTenSP"></label>
        <input name="textTenSP" type="text" id="textTenSP" value="<%=(SanPham.Fields.Item("TenSP").Value)%>" /></td>
      </tr>
      <tr>
        <td>Loại</td>
        <td><label for="Loai"></label>
          <select name="Loai" class="list-group-item-info" id="Loai">
            <%
While (NOT Loai.EOF)
%>
            <option value="<%=(Loai.Fields.Item("MaLoai").Value)%>" <%If (Not isNull((Loai.Fields.Item("MaLoai").Value))) Then If (CStr(Loai.Fields.Item("MaLoai").Value) = CStr((Loai.Fields.Item("MaLoai").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(Loai.Fields.Item("Loai").Value)%></option>
            <%
  Loai.MoveNext()
Wend
If (Loai.CursorType > 0) Then
  Loai.MoveFirst
Else
  Loai.Requery
End If
%>
          </select></td>
      </tr>
      <tr>
        <td>NSX</td>
        <td><label for="textfield2"></label>
          <label for="NSX"></label>
          <select name="NSX" class="list-group-item-info" id="NSX">
            <%
While (NOT NSX.EOF)
%>
            <option value="<%=(NSX.Fields.Item("MaNSX").Value)%>" <%If (Not isNull((NSX.Fields.Item("MaNSX").Value))) Then If (CStr(NSX.Fields.Item("MaNSX").Value) = CStr((NSX.Fields.Item("MaNSX").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(NSX.Fields.Item("NSX").Value)%></option>
            <%
  NSX.MoveNext()
Wend
If (NSX.CursorType > 0) Then
  NSX.MoveFirst
Else
  NSX.Requery
End If
%>
          </select></td>
      </tr>
      <tr>
        <td>Cấu hình</td>
        <td><label for="txtCauHinh"></label>
        <textarea  name="txtCauHinh" cols="50" rows="10" id="txtCauHinh" dir="ltr" lang="vi"><%=(SanPham.Fields.Item("CauHinh").Value)%></textarea></td>
      </tr>
      <tr>
        <td>Hình ảnh</td>
        <td><label for="HinhAnh"></label>
        <input name="HinhAnh" type="file" id="fileHinhAnh" /></td>
      </tr>
      <tr>
        <td>Số lượng</td>
        <td><label for="txtSoLuong"></label>
        <input name="txtSoLuong" type="text" id="txtSoLuong" value="<%=(SanPham.Fields.Item("SoLuong").Value)%>" /></td>
      </tr>
      <tr>
        <td>Giá</td>
        <td><label for="txtGia"></label>
        <input name="txtGia" type="text" id="txtGia" value="<%=(SanPham.Fields.Item("Gia").Value)%>" />
        VNĐ</td>
      </tr>
      <tr>
        <td>Ghi chú</td>
        <td><label for="txtGhiChu"></label>
        <input name="txtGhiChu" type="text" id="txtGhiChu" value="<%=(SanPham.Fields.Item("GhiChu").Value)%>" /></td>
      </tr>
      <tr>
      <td>
      </td>
        <td align="center"><button type="submit" name="button" id="button" value="CẬP NHẬT" >CẬP NHẬT</button></td>
      </tr>
    </table>
    <input type="hidden" name="RG_recordId" value="<%= SanPham.Fields.Item("MaSP").Value %>">
  </form>
</div>
<script src="js/vendor/jquery-1.10.1.min.js"></script>
<script src="js/plugins.js"></script>
<script src="js/main.js"></script>
<div class="footer-bar">
    <span class="article-wrapper">
        <span class="article-label">Trang quản lý</span>
        <span class="article-link"><a href="#top">Lên top</a></span>
    </span>
</div>
</body>
</html>
<!--#include file="../UploadFiles/Upload.asp" -->
<!--#include file="../UploadFiles/UploadAdvanced.asp" -->
<%
Loai.Close()
Set Loai = Nothing
%>
<%
NSX.Close()
Set NSX = Nothing
%>
<%
SanPham.Close()
Set SanPham = Nothing
%>
