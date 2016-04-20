<%

Function StringToBinary(String)
  Dim I, B
  For I=1 to len(String)
    B = B & ChrB(Asc(Mid(String,I,1)))
  Next
  StringToBinary = B
End Function

Function BinaryToString(Binary)
  Dim TempString 

  On Error Resume Next

  TempString = RSBinaryToString(Binary)
  If Len(TempString) <> LenB(Binary) then
    TempString = MBBinaryToString(Binary)
  end if
  BinaryToString = TempString
End Function


Function MBBinaryToString(Binary)

  dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    if cl3>300 then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      if cl2>200 then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  MBBinaryToString = pl1 & pl2 & pl3
End Function

Function MultiByteToBinary(MultiByte)
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = server.CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
	if LMultiByte>0 then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	End If
  MultiByteToBinary = Binary
End Function



Function GetHeadFields(ByVal Head, Name, FileName, Content_Disposition, Content_Type)
  
  Name = (SeparateField(Head, "name=", ";")) 
  If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)

  FileName = (SeparateField(Head, "filename=", ";"))

  If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)

  Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
  Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

Function SeparateField(From, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE, sFrom
  sFrom = LCase(From)
  PosB = InStr(sFrom, sStart)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(From, PosB, PosE - PosB)
  Else
    SeparateField = Empty
  End If
End Function

Function SplitFileName(FullPath)
  Dim Pos, PosF
  PosF = 0
  For Pos = Len(FullPath) To 1 Step -1
    Select Case Mid(FullPath, Pos, 1)
      Case ":", "/", "\": PosF = Pos + 1: Pos = 0
    End Select
  Next
  If PosF = 0 Then PosF = 1
	SplitFileName = PosF
End Function

Function GetPath(FullPath)
  GetPath = left(FullPath, SplitFileName(FullPath)-1)
End Function

Function GetFileName(FullPath)
  GetFileName = Mid(FullPath, SplitFileName(FullPath))
End Function

Function VerifyPath(Path)
	Dim FS
	Set FS = server.CreateObject("Scripting.FileSystemObject")
	if not FS.FolderExists(Path) then
		on error resume next
		FS.CreateFolder Path
		if err.number<>0 then
			response.write "You try to create the folder: " & Path 
			response.write "<br> but You don't have permission to create this folder<br>"
			response.write "<br> Verify the write permission<br>"
			on error goto 0
			response.End()
		end if
	end if
	set FS = nothing
end function

Function RecurseMKDir(ByVal Path)
  Dim FS: Set FS = server.CreateObject("Scripting.FileSystemObject")

  Path = Replace(Path, "/", "\")
  If Right(Path, 1) <> "\" Then Path = Path & "\"   '"
  Dim Pos, n
  Pos = 0: n = 0
  Pos = InStr(Pos + 1, Path, "\")   '"
  Do While Pos > 0
    On Error Resume Next
    FS.CreateFolder Left(Path, Pos - 1)
    If Err = 0 Then n = n + 1
    Pos = InStr(Pos + 1, Path, "\")   '"
  Loop
  RecurseMKDir = n
End Function

Function SaveBinaryData(FileName, ByteArray)
	SaveBinaryData = SaveBinaryDataStream(FileName, ByteArray)
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
  Dim FS : Set FS = server.CreateObject("Scripting.FileSystemObject")
	On error Resume next
  Dim TextStream 
	Set TextStream = FS.CreateTextFile(FileName)
	if Err = &H4c then
		On error Goto 0
		RecurseMKDir GetPath(FileName)
		On error Resume next
		Set TextStream = FS.CreateTextFile(FileName)
	end if

  TextStream.Write BinaryToString(ByteArray) 
  TextStream.Close

	Dim ErrMessage, ErrNumber
	ErrMessage = Err.Description
	ErrNumber = Err

	On Error Goto 0
	if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 

End Function

Function SaveBinaryDataStream(FileName, ByteArray)
	Dim BinaryStream
	Set BinaryStream = server.createobject("ADODB.Stream")
	BinaryStream.Type = 1 
	BinaryStream.Open
	BinaryStream.Write ByteArray
	VerifyPath GetPath(FileName)
	on error resume next
	BinaryStream.SaveToFile FileName, 2 	
	if err.number<>0 then
		response.write "You try to save the file: " & FileName 
		response.write "<br>but You don't have permission to write in the folder<br>"
		response.write "<br>Verify the write permission<br>"
		on error goto 0
		response.End()
	end if
End Function

function RG_RemoveSpecialChar( RG_Str)
	RG_New = ""
	for RG_z = 1 to len(RG_Str)
		RG_tst = asc(mid(RG_Str,RG_z,1))
		if (RG_tst>=97 and RG_tst<=122) or (RG_tst>=65 and RG_tst<=90) or (RG_tst>=48 and RG_tst<=57) or RG_tst=46 or RG_tst=95 or RG_tst=45 then
			RG_New = RG_New & mid(RG_Str,RG_z,1)
		end if
	next
	RG_RemoveSpecialChar = RG_New
end function

function RG_VerifyLen(RG_Len,RG_Max)
	RG_ret = -1
	if RG_Max<>"" and not isnull(RG_Max) then
		if CDbl(RG_Len) <= (Cdbl(RG_Max)*1000) then
			RG_ret = -1
		else
			RG_ret = 0
		end if
	end if
	RG_VerifyLen = RG_ret
end function

function RG_VerifyExt(RG_Name,RG_Ext)
	RG_ret = 0
	if RG_Ext="" or isnull(RG_Ext) then
		RG_ret = -1
	else
		RG_Extensions = split(replace(RG_Ext,".",""),",")
		RG_FileExt = mid(RG_Name,instrRev(RG_Name,".") + 1)
		for RG_L = 0 to Ubound(RG_Extensions)
			if Ucase(trim(CStr(RG_Extensions(RG_L)))) = Ucase(CStr(RG_FileExt)) then
				RG_ret = -1
				exit for
			end if
		next
	end if
	RG_VerifyExt = RG_ret
end function

function GetValById(x,y)
	GetValById = ParamVal(right("00" & cstr(x),3) & cstr(y))
end function

function GetValByName(val,Nome)
	trovato = false
	for x=0 to MaxFieldNumber
		if Ucase(ParamVal(right("00" & cstr(x),3) & IndexFieldName))=Ucase(CStr(Nome)) then
			trovato = true
			exit for
		end if
	next
	if not trovato then
		response.write ("Thông báo bị lỗi!")
		response.End()
	end if
	select case val
	case "UploadFolder"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "0")
	case "FolderType"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "1")
	case "FileName"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "2")
	case "FileNameFormula"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "11")
	case "AllowedExtension"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "3")
	case "FileSize"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "4")
	case "Conflict"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "5")
	case "FieldToSaveName"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "6")
	case "SaveType"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "7")
	case "FieldToSaveSize"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "8")
	case "DelExist"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "9")
	case "FieldExistField"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "10")
	case "MaxWidth"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "12")
	case "MaxHeight"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "13")
	case "DeleteExistFile"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "14")
	case "FieldToSaveThumb"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "15")
	case "ThumbSaveType"
		GetValByName = ParamVal(right("00" & cstr(x),3) & "16")
	end select		
end function

function setValByName(val,Nome,newVal)
	trovato = false
	for x=0 to MaxFieldNumber
		if Ucase(ParamVal(right("00" & cstr(x),3) & IndexFieldName))=Ucase(CStr(Nome)) then
			trovato = true
			exit for
		end if
	next
	if not trovato then
		response.write ("Error in function setValByName - Please report this to revellog@dwzone.it")
		response.End()
	end if
	select case val
	case "UploadFolder"
		ParamVal(right("00" & cstr(x),3) & "0") = newVal
	case "FolderType"
		ParamVal(right("00" & cstr(x),3) & "1") = newVal
	case "FileName"
		ParamVal(right("00" & cstr(x),3) & "2") = newVal
	case "FileNameFormula"
		ParamVal(right("00" & cstr(x),3) & "11") = newVal
	case "AllowedExtension"
		ParamVal(right("00" & cstr(x),3) & "3") = newVal
	case "FileSize"
		ParamVal(right("00" & cstr(x),3) & "4") = newVal
	case "Conflict"
		ParamVal(right("00" & cstr(x),3) & "5") = newVal
	case "FieldToSaveName"
		ParamVal(right("00" & cstr(x),3) & "6") = newVal
	case "SaveType"
		ParamVal(right("00" & cstr(x),3) & "7") = newVal
	case "FieldToSaveSize"
		ParamVal(right("00" & cstr(x),3) & "8") = newVal
	case "DelExist"
		ParamVal(right("00" & cstr(x),3) & "9") = newVal
	case "FieldExistField"
		ParamVal(right("00" & cstr(x),3) & "10") = newVal
	case "MaxWidth"
		ParamVal(right("00" & cstr(x),3) & "12") = newVal
	case "MaxHeight"
		ParamVal(right("00" & cstr(x),3) & "13") = newVal
	case "DeleteExistFile"
		ParamVal(right("00" & cstr(x),3) & "14") = newVal
	end select
end function


function GetValByName_Adv(val,Nome)
	trovato = false
	IndexFieldName_Adv = "5"
	for x=0 to MaxFieldNumberAdv
		if Ucase(ParamVal_AdvMail(right("00" & cstr(x),3) & IndexFieldName_Adv))=Ucase(CStr(Nome)) then
			trovato = true
			exit for
		end if
	next
	if not trovato then
		response.write ("Error in function GetValByName_Adv - Please report this to revellog@dwzone.it")
		response.End()
	end if
	'response.write Nome & " " & x & "<br>"
	select case val
	case "FieldToSaveName"
		GetValByName_Adv = ParamVal_AdvMail(right("00" & cstr(x),3) & "0")
	case "SaveType"
		GetValByName_Adv = ParamVal_AdvMail(right("00" & cstr(x),3) & "1")
	case "FieldToSaveSize"
		GetValByName_Adv = ParamVal_AdvMail(right("00" & cstr(x),3) & "2")
	case "DelExist"
		GetValByName_Adv = ParamVal_AdvMail(right("00" & cstr(x),3) & "3")
	case "FieldExistField"
		GetValByName_Adv = ParamVal_AdvMail(right("00" & cstr(x),3) & "4")
	end select		
end function



function GetValByNameResize(val,Nome)
	trovato = false
	for x=0 to MaxFieldNumber
		if Ucase(ParamVal(right("00" & cstr(x),3) & IndexFieldName))=Ucase(CStr(Nome)) then
			trovato = true
			exit for
		end if
	next
	if not trovato then
		response.write ("Error in function GetValByNameResize - Please report this to revellog@dwzone.it")
		response.End()
	end if
	
	tmp = split(RG_Files,"@_@_@")
	if ubound(tmp) > 2 then
		if trim(tmp(3))<>"" then
			tmp2 = split(tmp(3),"|")
			tmp3 = split(tmp2(x),";")
			select case val
			case "Thumb"
				GetValByNameResize = tmp3(0)
			case "MaxWidth"
				GetValByNameResize = tmp3(1)
			case "MaxHeight"
				GetValByNameResize = tmp3(2)
			case "JpegQuality"
				GetValByNameResize = tmp3(3)
			case "KeepAspect"
				GetValByNameResize = tmp3(4)
			case "ThumbWidth"
				GetValByNameResize = tmp3(5)
			case "ThumbHeight"
				GetValByNameResize = tmp3(6)
			case "ThumbJpegQuality"
				GetValByNameResize = tmp3(7)
			case "ThumbKeepAspect"
				GetValByNameResize = tmp3(8)
			case "Suffix"
				GetValByNameResize = tmp3(9)
			case "Resize"
				GetValByNameResize = tmp3(10)
			case "SuffixPosition"
				GetValByNameResize = tmp3(11)
			end select
		else
			GetValByNameResize = "0"
		end if
	else
		GetValByNameResize = "0"
	end if	
end function


function UploadFormRequest(str)
	UploadFormRequest = Form.Texts.ValueByName(str)
end function

function getRedirect()
	if instr(1,editRedirectUrl,".html",vbtextcompare)>0 or instr(1,editRedirectUrl,".htm",vbtextcompare)>0 then
		newPage = "<script language=javascript>location='" & editRedirectUrl & "'</script>"
	else
		if Request.QueryString<>"" then
			newQryString = ""
			cong = ""
			For Each item In Request.QueryString
				if lcase(item)<>"uploadid" then
					newQryString = cong & item & "=" & Request.QueryString(item)
					cong = "&"
				end if
			Next
			if newQryString<>"" then
				if instr(1,editRedirectUrl,"?",vbtextcompare)>0 then
					separ = "&"
				else
					separ = "?"
				end if
			end if
			editRedirectUrl = editRedirectUrl & separ & newQryString
		end if
		
		newPage = "<html><head><script language=" & chr(34) & "JavaScript" & chr(34) & ">function sendForm(){myForm.submit();}</script></head><body onLoad=" & chr(34) & "sendForm()" & chr(34) & ">"
		newPage = newPage & "<form action=" & editRedirectUrl & " method=" & chr(34) & "get" & chr(34) & " name=" & chr(34) & "myForm" & chr(34) & ">"
		if UploadError = "" then
			newPage = newPage & UploadStatus
			newPage = newPage & "<input name=" & chr(34) & "NumFile" & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & NumFile & chr(34) & ">"
		else
			newPage = newPage & "<input name=" & chr(34) & "UploadedSize" & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & UploadError & chr(34) & ">"
			newPage = newPage & "<input name=" & chr(34) & "MaxFileSize" & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & TotalFileSize & chr(34) & ">"
		end if
		
		if valueToRedirectSend <> "" then
			listToSend = split(valueToRedirectSend,"|")
			for x = 0 to ubound(listToSend)
				newPage = newPage & "<input name=" & chr(34) & listToSend(x) & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & UploadFormRequest(listToSend(x)) & chr(34) & ">"
			next
		end if
		
		newPage = newPage & "</form></body></html>"
		
	end if		
	getRedirect = newPage	
end function

function verFieldTableName(Nome)
	if left(Nome,1)<>"[" then
		Nome = "[" & Nome
	end if
	if right(Nome,1)<>"]" then
		Nome = Nome & "]"
	end if
	Nome = replace(Nome,".","].[")
	verFieldTableName = Nome
end function

function isSingleRecord()
	tmp = split(RG_Files,"@_@_@")
	if tmp(1)<>"" then
		if clng(tmp(1)) = 1 then
			ris = -1
		else
			ris = 0
		end if
	else
		ris = -1
	end if
	isSingleRecord = ris
end function


function getFileSize(FilePath)
	set f = server.CreateObject("Scripting.FileSystemObject")
	set ff = f.getFile(FilePath)
	getFileSize = ff.size
	set f = nothing
end function

sub deleteFile(filePath)
	set RG_FS = server.CreateObject("Scripting.FileSystemObject")
	if RG_FS.FileExists(filePath) then
		set FL = RG_FS.getFile(filePath)
		FL.delete
	end if
	set RG_FS = nothing
end sub

function getProgressPath()
	val = split(Rg_Files,"@_@_@")
	getProgressPath = trim(val(4))
end function


function objWork(obj)
	on error resume next
	Set objTest = Server.CreateObject(obj)
	if err.number<>0 then
		retStr = false
	else
		retStr = true
	end if
	on error goto 0
	err.clear
	objWork = retStr
end function


function getNewFileName(g_FilePath, g_isThumb, g_Suffix, g_imgExt, g_SuffPos)
	pos = instrRev(g_FilePath,"\")
	if pos>0 then
		g_Path = left(g_FilePath,pos)
	else
		g_Path = ""
	end if
	g_Name = mid(g_FilePath,pos+1)
	g_Name = left(g_Name,instrRev(g_Name,".")-1)
	if g_isThumb = "-1" then
		if g_Suffix = "" then
			g_Suffix = "s"
		end if
		if g_SuffPos = "" then
			g_SuffPos = "1"
		end if
		if g_SuffPos = "1" then
			g_Name = g_Name & g_Suffix & "." & g_imgExt
		else
			g_Name = g_Suffix & g_Name & "." & g_imgExt
		end if
	else
		g_Name = g_Name & "." & g_imgExt
	end if	
	getNewFileName = g_Path & g_Name
end function

%>
