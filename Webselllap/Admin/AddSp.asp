<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/Connection.asp" -->
<%
	server.ScriptTimeout = 5400

	Dim RG_altVal, RG_columns, RG_Cong, RG_dbValues, RG_dbValuesTmp, RG_delim, RG_editCmd, RG_editQuery, RG_editQueryTmp, RG_emptyVal, RG_Ext, RG_Extensions, RG_fields, RG_FieldValueTmp, RG_FileDel, RG_FileExt, RG_formVal, RG_FS, RG_i, RG_L, RG_Len, RG_Max, RG_Name, RG_New, RG_newName, RG_Num, RG_Path, RG_Rec, RG_ret, RG_Save, RG_tableValues, RG_tableValuesTmp, RG_tst, RG_typeArray, RG_z, UploadStatus, NumFile
	Dim RG_Connection, RG_editColumn, RG_recordId, Form, editAction, editRedirectUrl, RG_Files, RG_formName, UploadType, ParamVal, ParamList, MaxFieldNumber, TmpVal, x, y, Key, ProgressBar
	Dim tmpField_Name(), tmpValue_Name(), tmpField_Size(), tmpValue_Size(), tmpField_Thumb(), tmpValue_Thumb(), QtyRecord, TotalFileSize, valueToRedirectSend
	
	Set Form = New ASPForm
	Dim UploadID
	UploadID = Form.NewUploadID
	ProgressBar = "progress-std.asp"
	TotalFileSize = ""
  	editRedirectUrl = "Products.asp"
	RG_Connection = MM_Connection_STRING
	RG_editTable = "dbo.SanPham"	
	RG_Files = "/Images;1;;;;0;HinhAnh;1;;;;0;;;;;;txtHinhAnh@_@_@1@_@_@ @_@_@@_@_@../"
	RG_formName = "form1"
	UploadType="Insert"
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

		RG_fieldsStr  = "txtTenSp|value|txtTinhtrang|value|txtLoai|value|txtNSX|value|txtCauHinh|value|txtSoLuong|value|txtGia|value|txtGhiChu|value"
  		RG_columnsStr = "TenSP|',none,''|Tinhtrang|none,none,NULL|MaLoai|none,none,NULL|MaNSX|none,none,NULL|CauHinh|',none,''|SoLuong|none,none,NULL|Gia|none,none,NULL|GhiChu|',none,''"
		Form.Files.DataBaseInsert

		response.write(getRedirect())
		response.end
  End If
ElseIf Form.State > 10 then
  response.write "<br><Font Color=red>Some form error.<br>Error code: " & Form.State & "<br>Error List:<br>0  Form was successfully processed. <br>1  Request method is NOT post <br>2  Zero length request (there are no data in a source form) <br>3  Form is in a middle of process. <br>5  Initial form state <br>11  Boundary of multipart/form-data is not specified. <br>12  Unknown source form (Content-type must be multipart/form-data) <br>15  Client was disconnected before upload was completted.<br>16  Unexpected error from Request.BinaryRead method (ASP error).<br></Font><br>"
End If


function GetFolderName(str):  GetFolderName = Ris : end function

function myGetFileName(str):  myGetFileName = Ris : end function
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
Dim LoaiSP
Dim LoaiSP_cmd
Dim LoaiSP_numRows

Set LoaiSP_cmd = Server.CreateObject ("ADODB.Command")
LoaiSP_cmd.ActiveConnection = MM_Connection_STRING
LoaiSP_cmd.CommandText = "SELECT * FROM dbo.LoaiSP ORDER BY MaLoai ASC" 
LoaiSP_cmd.Prepared = true

Set LoaiSP = LoaiSP_cmd.Execute
LoaiSP_numRows = 0
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
    <link rel="stylesheet" href="css/font-awesome.min.css">
    <link rel="stylesheet" href="css/animate.css">
    <link rel="stylesheet" href="css/templatemo_misc.css">
    <link rel="stylesheet" href="css/templatemo_style.css">
    <script src="js/vendor/modernizr-2.6.2.min.js"></script>
<script src="../js/jquery.min.js"></script>
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
<h1 style="color:#0CF" size="300%" align="center">Thêm Sản Phẩm</h1>
    <div class="item">
    <form ACTION="<%=editAction%>" onsubmit="return ProgressBar()" method="post" enctype="multipart/form-data" name="form1" id="form1">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><p>Tên sản phẩm</p></td>
          <td><label for="txtTenSp"></label>
          <input type="text" name="txtTenSp" id="txtTenSp">
          <input name="txtTinhtrang" type="hidden" id="txtTinhtrang" value="1"></td>
        </tr>
        <tr>
          <td><p>Loại</p></td>
          <td><label for="txtLoai2"></label>
            <select name="txtLoai" class="list-group-item-info" id="txtLoai" dir="ltr" lang="vi">
              <%
While (NOT LoaiSP.EOF)
%>
              <option value="<%=(LoaiSP.Fields.Item("MaLoai").Value)%>" <%If (Not isNull((LoaiSP.Fields.Item("MaLoai").Value))) Then If (CStr(LoaiSP.Fields.Item("MaLoai").Value) = CStr((LoaiSP.Fields.Item("MaLoai").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(LoaiSP.Fields.Item("Loai").Value)%></option>
              <%
  LoaiSP.MoveNext()
Wend
If (LoaiSP.CursorType > 0) Then
  LoaiSP.MoveFirst
Else
  LoaiSP.Requery
End If
%>
          </select></td>
        </tr>
        <tr>
          <td><p>NSX</p></td>
          <td><label for="txtNSX"></label>
            <select name="txtNSX" class="list-group-item-info" id="txtNSX" dir="ltr" lang="vi">
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
          <td><p>Cấu hình</p></td>
          <td><label for="txtCauHinh"></label>
            <textarea  name="txtCauHinh" cols="50" rows="10" id="txtCauHinh" dir="ltr" lang="vi"></textarea></td>
        </tr>
        <tr>
          <td><p>Hình ảnh</p></td>
          <td><label for="txtHinhAnh">
            <input type="file" name="txtHinhAnh" id="txtHinhAnh">
          </label></td>
        </tr>
        <tr>
          <td><p>Số lượng </p></td>
          <td>
			<label for="txtSoLuong"></label>
            <input name="txtSoLuong" type="text" id="txtSoLuong">
          </label></td>
        </tr>
        <tr>
          <td><p>Giá</p></td>
			<td>
                <label for="txtGia">
                <input type="text" name="txtGia" id="txtGia">
                </label>
			</td>
        </tr>
        <tr>
		<td><p>Ghi chú</p></td>
		<td>
        	<label for="txtGhiChu">
          	<input type="text" name="txtGhiChu" id="txtGhiChu">
            </label>
		</td>
        </tr>
        <tr>
        <td align="right">
            <button type="submit" name="Submit" id="button" value="Thêm">Thêm</button>
        </td>
        <td align="center">
          <button type="reset" name="button2" id="button2" value="Hủy">Hủy</button>
        </td>
        </tr>
      </table>
    </form>
</div>

    <script src="js/vendor/jquery-1.10.1.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery-1.10.1.min.js"><\/script>')</script>
    <script src="js/jquery.easing-1.3.js"></script>
    <script src="js/bootstrap.js"></script>
    <script src="js/plugins.js"></script>
    <script src="js/main.js"></script>
<div class="footer-bar">
    <span class="article-wrapper">
        <span class="article-label">Trang quản lý</span>
        <span class="article-link"><a href="#">Lên top</a></span>
    </span>
</div>
</body>
</html>
<!--#include file="../UploadFiles/Upload.asp" -->
<!--#include file="../UploadFiles/UploadAdvanced.asp" -->
<%
NSX.Close()
Set NSX = Nothing
%>
<%
LoaiSP.Close()
Set LoaiSP = Nothing
%>
