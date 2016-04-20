<%@EnableSessionState=False%>
<%ThisPage = "Progress"%>
<!--#include file="TempFolder.asp"-->
<%
	Set FS = server.CreateObject("Scripting.FileSystemObject")
	lista = ""
	'TempFolder = FS.GetSpecialFolder(2)
	TempFolder = getTempFolder()
	TempFileName = TempFolder & "\pu" & Request.QueryString("UploadID") & ".~tmp"
	UploadEnd = false
	if FS.FileExists(TempFileName) then
		do while 1
			err.clear
			on error resume next
			set tmpFile = FS.OpenTextFile(TempFileName, 1)
			if err.number=0 then
				lista = tmpFile.readAll
				tmpFile.close
				on error goto 0
				exit do
			end if
			on error goto 0
		loop
		if Lista<>"" then
			valori = split(lista,vbcrlf)
			if trim(valori(0))="DONE" then
				UploadEnd = true
				start = -1
				OnLoad = "onLoad=" & chr(34) & "javascript:setTimeout('closeWin()',150)" & chr(34)
			end if
			start = -1
		else
			start = 0
		end if
	end if
	
	Const RefreshTime = 2'Seconds

	if Request.QueryString("UploadID") = "" then 
		response.write "<script language=javascript>top.close()</script>"
	end if
	Server.ScriptTimeout = 1000
	if not UploadEnd then
		response.cachecontrol = "no-cache"
		response.AddHeader "Pragma","no-cache"
	  	response.addheader "Refresh", RefreshTime
	end if
	
if start then 
	
	BytesRead = valori(2)
	TotalBytes = valori(1)

	if TotalBytes>0 then 
		PercBytesRead = int(100*BytesRead/TotalBytes)
		PercentRest = 100-PercBytesRead
	
		if valori(3)>0 then 
			TransferRate = BytesRead / valori(3)
		end if
		if TransferRate>0 then 
			RestTime = FormatTime((TotalBytes-BytesRead) / TransferRate)
		end if
		TransferRate = FormatSize(1000 * TransferRate)
	else
		RestTime = "?"
		PercBytesRead = 0
		PercentRest = 100
		TransferRate = "?"
	end if

	'Create graphics progress bar.
	'The bar is created with blue (TDsread, completted) / blank (TDsRemain, remaining) TD cells.
	Dim TDsread, TDsRemain
	
	select case request.QueryString("ProgressPage")
	case "Progress-BigBar.asp"
		TDsread = replace(space(0.25*PercBytesRead), " ", "<TD BGColor=blue > </TD>")
		TDsRemain = replace(space(0.25*PercentRest), " ", "<TD > </TD>")
		
		Bar = "<Table cellpadding=0 height=20 cellspacing=0 border=1 width=100" & chr(37) & " style='border:1px inset white' ><tr>" & TDsread & TDsRemain & "</tr></table>"

	case "Progress-BigBar2.asp"
		TDsread = replace(space(0.25*PercBytesRead), " ", "<TD BGColor=blue > </TD>")
		TDsRemain = replace(space(0.25*PercentRest), " ", "<TD > </TD>")
		
		Bar = "<Table cellpadding=0 height=20 cellspacing=2 border=0 width=100" & chr(37) & "  ><tr>" & TDsread & TDsRemain & "</tr></table>"

	case "Progress-LittleBar.asp"
		TDsread = replace(space(0.5*PercBytesRead), " ", "<TD BGColor=blue > </TD>")
	  	TDsRemain = replace(space(0.5*PercentRest), " ", "<TD > </TD>")
		Bar = "<Table cellpadding=0 height=20 cellspacing=0 border=1 width=100" & chr(37) & " style='border:1px inset white' ><tr>" & TDsread & TDsRemain & "</tr></table>"
	
	case "Progress-LittleBar2.asp"
		TDsread = replace(space(0.5*PercBytesRead), " ", "<TD BGColor=blue > </TD>")
  		TDsRemain = replace(space(0.5*PercentRest), " ", "<TD > </TD>")
		Bar = "<Table cellpadding=0 height=20 cellspacing=2 border=0 width=100" & chr(37) & " style='border:1px inset white' ><tr>" & TDsread & TDsRemain & "</tr></table>"
	
	case "progress-Graphics3D.asp"
		Bar = "<Table cellpadding=0 height=20 cellspacing=0 border=0 width=100" & chr(37) & " ><tr><TD align=left><img src=3D_Bar.gif width=" & int(3.35*PercBytesRead) & " height=20></TD></tr></table>"
	
	case else
		TDsread = replace(space(0.5*PercBytesRead), " ", "<TD BGColor=blue> </TD>")
		TDsRemain = replace(space(0.5*PercentRest), " ", "<TD> </TD>")
		Bar = "<Table cellpadding=0 height=20 cellspacing=0 border=0 width=100" & chr(37) & " ><tr>" & TDsread & TDsRemain & "</tr></table>"
	end select

	'Format output values.
	UploadTime = FormatTime(valori(3))
	TotalBytes = FormatSize(valori(1))
	BytesRead = FormatSize(BytesRead)
else
	Bar = "<Table cellpadding=0 height=20 cellspacing=0 border=0 width=100" & chr(37) & " ><tr><TD>&nbsp;</TD></tr></table>"
end if

Function FormatTime(byval ms)
	ms = 0.001 * ms 'get second
	FormatTime = (ms \ 60) & ":" & right("0" & (ms mod 60),2) & "s"
End Function 

'Format bytes to a string
Function FormatSize(byval Number)
  if isnumeric(Number) then
    if Number > &H100000 then'1M
      Number = FormatNumber (Number/&H100000,1) & "MB"
    elseif Number > 1024 then'1k
      Number = FormatNumber (Number/1024,1) & "kB"
    else
      Number = FormatNumber (Number,0) & "B"
    end if
  end if
  FormatSize = Number
End Function

%>
<HTML>
<Head>
<script language="JavaScript">
	parent.document.title = "<%=PercBytesRead%>% completed - upload to <%=Request.ServerVariables("HTTP_HOST")%> in progress";;
	function closeWin(){
		top.close()
	}
	
</script>
<meta http-equiv="Page-Enter" content="revealTrans(Duration=0,Transition=6)">
<%if not UploadEnd then%>
<META HTTP-EQUIV="Refresh" CONTENT="<%=RefreshTime%>">
<%end if%>
<Title>Upload in Progress</Title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
.Testo12BB {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #003399;
}
.Testo12N {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: normal;
	color: #000000;
}
</style>
</Head>

<Body <%=OnLoad%> BGcolor=#CCCCCC scroll="no" LeftMargin=5 TopMargin=5 marginwidth="5" marginheight="5">
<Table width=100% border=0 align="center" cellpadding=3 cellspacing=0 >
  <tr>
    <td width="9%" class="Testo12N">Uploading: </td>
	<td width="91%" class="Testo12N"><%=TotalBytes%> to <%=Request.ServerVariables("HTTP_HOST")%> ...</td>
</tr></Table>

<Table width=100% border=0 align="center" cellpadding=1 cellspacing=0 >
  <tr>
    <td valign="middle" align="left"><%=Bar%></td>
  </tr>
</table>

<Table width="100%" border=0 align="center" cellpadding=3 cellspacing=0>
  <tr> 
    <Td width="77" class="Testo12N" >Progress :</td>
    <Td width="629" class="Testo12N" ><%=BytesRead%> of <%=TotalBytes%> (<%=PercBytesRead%>%) </Td>
  </tr>
  <tr> 
    <Td class="Testo12N">Time :</td>
    <Td class="Testo12N"><%=UploadTime%> (<%=TransferRate%>/s) </Td>
  </tr>
  <tr> 
    <Td class="Testo12N">Time left :</td>
    <Td class="Testo12N"><%=RestTime%> </Td>
  </tr>
</table>
 
</Body>
</HTML>
