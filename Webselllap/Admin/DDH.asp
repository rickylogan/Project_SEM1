<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Connection.asp" -->
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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Connection_STRING
    MM_editCmd.CommandText = "UPDATE dbo.DonDatHang SET TinhTrang = ? WHERE MaDDH = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 20, Request.Form("TinhTrang")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "DDH.asp"
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
Dim DDH
Dim DDH_cmd
Dim DDH_numRows

Set DDH_cmd = Server.CreateObject ("ADODB.Command")
DDH_cmd.ActiveConnection = MM_Connection_STRING
DDH_cmd.CommandText = "SELECT * FROM dbo.DonDatHang ORDER BY MaDDH DESC" 
DDH_cmd.Prepared = true

Set DDH = DDH_cmd.Execute
DDH_numRows = 0
%>
<%
Dim Count_DDH
Dim Count_DDH_cmd
Dim Count_DDH_numRows

Set Count_DDH_cmd = Server.CreateObject ("ADODB.Command")
Count_DDH_cmd.ActiveConnection = MM_Connection_STRING
Count_DDH_cmd.CommandText = "SELECT COUNT(MaDDH) FROM dbo.DonDatHang" 
Count_DDH_cmd.Prepared = true

Set Count_DDH = Count_DDH_cmd.Execute
Count_DDH_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 4
Repeat1__index = 0
Dim Num_page
Num_page = Repeat1__numRows
DDH_numRows = DDH_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim DDH_total
Dim DDH_first
Dim DDH_last

' set the record count
DDH_total = DDH.RecordCount

' set the number of rows displayed on this page
If (DDH_numRows < 0) Then
  DDH_numRows = DDH_total
Elseif (DDH_numRows = 0) Then
  DDH_numRows = 1
End If

' set the first and last displayed record
DDH_first = 1
DDH_last  = DDH_first + DDH_numRows - 1

' if we have the correct record count, check the other stats
If (DDH_total <> -1) Then
  If (DDH_first > DDH_total) Then
    DDH_first = DDH_total
  End If
  If (DDH_last > DDH_total) Then
    DDH_last = DDH_total
  End If
  If (DDH_numRows > DDH_total) Then
    DDH_numRows = DDH_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = DDH
MM_rsCount   = DDH_total
MM_size      = DDH_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
DDH_first = MM_offset + 1
DDH_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (DDH_first > MM_rsCount) Then
    DDH_first = MM_rsCount
  End If
  If (DDH_last > MM_rsCount) Then
    DDH_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
Dim Repeat00_numRows
Dim Repeat00_index
Dim Page
Page = 0

Repeat00_numRows = (Count_DDH.Fields.Item("").Value)\Repeat1__numRows + 1
Repeat00_index = 0
Page_numRows = Page_numRows + Repeat00_numRows
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Quản trị viên :: Groupfour</title>
<link rel="shortcut icon" href="../images/icon.png">
    <meta name="description" content="">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" href="css/bootstrap.min.css">
    <link rel="stylesheet" href="css/normalize.min.css">
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
if Session("TenAD") = "" then
	Response.Redirect("loginAD.asp")
else
	Response.write("Xin chào, <b class=tentk>" & Session("TenAD") & "</b><b class=to> |</b>" & "<a href=logoutAD.asp class=colorlink2><ins>Thoát</ins></a>")
	
end if
%>
</div>
<div id="top"><p class=title align=center>QUẢN LÝ SẢN PHẨM</p></div>
                         <!-- /.container -->
        </div> <!-- /.site-header -->
</div> <!-- /#front -->
<div class="site-slider"></div>
<div class="clear"></div>
    <p>
      <script src="js/vendor/jquery-1.10.1.min.js"></script>
      <script src="js/plugins.js"></script>
      <script src="js/main.js"></script>
    </p>
<h1 style="color:rgb(0, 66, 255)" size="300%" align="center">ĐƠN ĐẶT HÀNG</h1>
<%
	  Dim M_TinhTrang
%>
<% 
While ((Repeat1__numRows <> 0) AND (NOT DDH.EOF))
	M_TinhTrang=(DDH.Fields.Item("TinhTrang").Value)
%>

  <div class="oneItem">
    <table width="100%" style="margin-left:10px" border="2px" Bordercolor="black" cellspacing="0" cellpadding="100px" align="center">
      <tr>
        <td colspan="2" align="center"><b>ĐƠN ĐẶT HÀNG</b></td>
      </tr>
      <tr>
        <td width="50%">    Mã đơn đặt hàng</td>
        <td width="50%">    <b><%=(DDH.Fields.Item("MaDDH").Value)%></b></td>
      </tr>
      <tr>
        <td>    Tài khoản khách hàng</td>
        <td>    <%=(DDH.Fields.Item("TKKH").Value)%></td>
      </tr>
      <tr>
        <td>    Ngày đặt</td>
        <td>    <%=(DDH.Fields.Item("NgayDat").Value)%></td>
      </tr>
      <tr>
        <td>    Tổng tiền</td>
        <td>    <%=(DDH.Fields.Item("TongTien").Value)%><em> <u>VNĐ</u></em></td>
      </tr>
      <tr>
        <td>    Tình trạng</td>
        <td>    <%=(DDH.Fields.Item("TinhTrang").Value)%></td>
      </tr>
      <%
	  Content = ""
	  if M_TinhTrang ="Đã thanh toán       " then
	  Content = Content & "</table><div style=margin-top:65px align=right><form action=CTDDH.asp method=post name=form1 id=form1> <p><input name=MaDDH type=hidden id=MaDDH value=" & DDH.Fields.Item("MaDDH").Value & "></p><button type=submit name=button id=button value=>XEM CHI TIẾT</button></form>"
	  else
	  Content = Content & "<tr><td colspan=2 align=center><form name=form1 method=POST action=" & MM_editAction & ">"
	  Content = Content & "<input name=TinhTrang type=hidden id=TinhTrang value='Đã thanh toán'>"
	  Content = Content & "<button type=submit name=button2 id=button2 value=>XÁC NHẬN THANH TOÁN</button>"
	  Content = Content & "<input type=hidden name=MM_update value=form1>"
	  Content = Content & "<input type=hidden name=MM_recordId value="& (DDH.Fields.Item("MaDDH").Value) & ">"
	  Content = Content & "</form></td></tr>"
	  Content = Content & "</table><div align=right><form action=CTDDH.asp method=post name=form2 id=form2> <p><input name=MaDDH type=hidden id=MaDDH value=" & (DDH.Fields.Item("MaDDH").Value) & "></p><button type=submit name=button id=button value=>XEM CHI TIẾT</button></form>"
	  end if
	  Response.Write(Content)
	  %>
    </div>
  </div>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  DDH.MoveNext()
Wend
%>
<span class="pageDDH">
        <A HREF="<%=MM_movePrev%>">
        <button class="paging_left"> ◄&nbsp;&nbsp;&nbsp;<ins>TRƯỚC</ins> </button>
        </A>
<% 
While ((Repeat00_numRows <> 0) AND (NOT DDH.EOF)) 
MM_numPage   = MM_urlStr & Page * Num_page
Page=Page+1
%>

  <A HREF="<%=MM_numPage%>">
    <button class="paging_mid"><%=Page%></button>
</A>
  <% 
  Repeat00_numRows=Repeat00_numRows-1
Wend
%>
        <A HREF="<%=MM_moveNext%>">
        <button class="paging_right"> <ins>SAU</ins>&nbsp;&nbsp;&nbsp;► </button>
        </A>
</span>
  </div>
<script src="js/vendor/jquery-1.10.1.min.js"></script>
<script src="js/plugins.js"></script>
<script src="js/main.js"></script>
<div class="footer-bar">
	<span >
        <span  style="margin-left:20px;"><a href="#" target="_top">Lên <ins>TOP▲</ins></a></span>
    </span>
</div>
</body>
</html>
<%
DDH.Close()
Set DDH = Nothing
%>
<%
Count_DDH.Close()
Set Count_DDH = Nothing
%>
