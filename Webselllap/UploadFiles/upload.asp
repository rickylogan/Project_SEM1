<!--#include file="TempFolder.asp"-->
<%
Const IndexFieldName = "17"

Const adTypeBinary = 1
Const adTypeText = 2
Const xfsCompleted    = &H0 '0  Form was successfully processed. 
Const xfsNotPost      = &H1 '1  Request method is NOT post 
Const xfsZeroLength   = &H2 '2  Zero length request (there are no data in a source form) 
Const xfsInProgress   = &H3 '3  Form is in a middle of process. 
Const xfsNone         = &H5 '5  Initial form state 
Const xfsError        = &HA '10  
Const xfsNoBoundary   = &HB '11  Boundary of multipart/form-data is not specified. 
Const xfsUnknownType  = &HC '12  Unknown source form (Content-type must be multipart/form-data) 
'Const xfsSizeLimit    = &HD '13  Form size exceeds allowed limit (ScriptUtils.ASPForm.SizeLimit) 
Const xfsTimeOut      = &HE '14  Upload time exceeds allowed limit (ScriptUtils.ASPForm.ReadTimeout) 
Const xfsNoConnected  = &HF '15  Client was disconnected before upload was completted.
Const xfsErrorBinaryRead = &H10 '16  Unexpected error from Request.BinaryRead method (ASP error).
Dim ErrForm, UploadError
ErrForm = ""
UploadError = ""

Class ASPForm
	Private m_ReadTime
	Public ChunkReadSize, BytesRead, TotalBytes, UploadID

	Public TempPath, MaxMemoryStorage, CharSet, FormType, SourceData, ReadTimeout

	public Default Property Get Item(Key)
		Set Item = m_Items.Item(Key)
	End Property

	public Property Get Items
		Read
		Set Items = m_Items
	End Property

	public Property Get Files
		Read
		Set Files = m_Items.Files
	End Property

	public Property Get Texts
		Read
		Set Texts = m_Items.Texts
	End Property
	
	public Property Get NewUploadID
		Randomize
		NewUploadID = clng(rnd * &H7FFFFFFF)
	End Property

	Public Property Get ReadTime
		if isempty(m_ReadTime) then
			if not isempty(StartUploadTime) then ReadTime = Clng((Now() - StartUploadTime) * 86400 * 1000)
		else 
			ReadTime = m_ReadTime
		end if
	End Property

	Public Property Get State
		if m_State = xfsNone Then Read
		State = m_State
	End Property

	Private Function CheckRequestProperties
		set Conn = Server.CreateObject("ADODB.Connection")
		adoVersion = CSng(Replace(Conn.Version,".",","))
		set Conn = Nothing
		if adoVersion < 2.5 then
			Response.Write "<b>You don't have ADO 2.5 installed on the server.</b><br/>"
			Response.Write "The Asp Upload extension needs ADO 2.5 or greater to run properly.<br/>"
			Response.Write "You can download the latest MDAC (ADO is included) from <a href=""www.microsoft.com/data"">www.microsoft.com/data</a><br/>"
			Response.End
		end if
	
	  If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
			ErrForm = "Actual value: " & Request.ServerVariables("REQUEST_METHOD")
			ErrForm = ErrForm & " - Allowed value: POST"
			m_State = xfsNotPost 
			Exit Function
		End If
	
		Dim CT
		CT = Request.ServerVariables("HTTP_Content_Type") 
		if len(CT) = 0 then CT = Request.ServerVariables("CONTENT_TYPE") 
	  If LCase(Left(CT, 19)) <> "multipart/form-data" Then 
			ErrForm = "Actual value: " & Left(CT, 19)
			ErrForm = ErrForm & " - Allowed value: multipart/form-data"
			m_State = xfsUnknownType 
			Exit Function
		End If 

		Dim PosB
		PosB = InStr(LCase(CT), "boundary=")
		If PosB = 0 Then
			m_State = xfsNoBoundary
			Exit Function
		End If 
		If PosB > 0 Then Boundary = Mid(CT, PosB + 9)

		PosB = InStr(LCase(CT), "boundary=") 
		If PosB > 0 then 
			PosB = InStr(Boundary, ",")
			If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)
		end if
		
		On Error Resume next
		TotalBytes = Request.TotalBytes
		If Err<>0 Then
			TotalBytes = CLng(Request.ServerVariables("HTTP_Content_Length"))
			if len(TotalBytes)=0 then TotalBytes = CLng(Request.ServerVariables("CONTENT_LENGTH")) 
		End If
		
		If TotalBytes = 0 then
			m_State = xfsZeroLength 
			Exit Function
		End If

		CheckRequestProperties = True

	End Function

	Public Sub Read()
		if m_State <> xfsNone Then Exit Sub
		If Not CheckRequestProperties Then 
			WriteProgressInfo
			Exit Sub
		End If

		if isempty(bSourceData) then Set bSourceData = server.createobject("ADODB.Stream")
		bSourceData.Open
		bSourceData.Type = 1

		Dim DataPart, PartSize
		BytesRead = 0
		StartUploadTime = Now

		Do While BytesRead < TotalBytes
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			bSourceData.Write DataPart

			WriteProgressInfo

			If Not Response.IsClientConnected Then
				m_State = xfsNoConnected  
				Exit Sub
			End If
		Loop
		m_State = xfsCompleted

		ParseFormData
	End Sub

	Private Sub ParseFormData
		Dim Binary
		bSourceData.Position = 0
		Binary = bSourceData.Read
		m_Items.mpSeparateFields Binary, Boundary
	End Sub

	Public Function getForm(FormID)
		if isempty(ProgressFile.UploadID) Then 
			ProgressFile.UploadID = FormID
		End If

		Dim ProgressData
		
		ProgressData = ProgressFile
		
		if len(ProgressData) > 0 then
			if ProgressData = "DONE" Then 
				ProgressFile.Done
				Err.Raise 1, "getForm", "Upload was done"
			Else
				ProgressData = Split (ProgressData, vbCrLf)
				if ubound(ProgressData) = 3 Then
					m_State = clng(ProgressData(0))
					TotalBytes = clng(ProgressData(1))
					BytesRead = clng(ProgressData(2))
					m_ReadTime = clng(ProgressData(3))
				End If
			End If
		end if
		Set getForm = Me
	End Function

	Private Sub WriteProgressInfo
		If UploadID > 0 Then
			if isempty(ProgressFile.UploadID) Then
				ProgressFile.UploadID = UploadID
			End If

			Dim ProgressData, FileName
			ProgressData = m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
			ProgressFile.Contents = ProgressData
		End If
	End Sub

	Private Sub Class_Initialize()
		ChunkReadSize = &H10000/2 '32 kB

		BytesRead = 0
		m_State = xfsNone
		
		TotalBytes = Request.TotalBytes

		Set ProgressFile = New cProgressFile
		Set m_Items = New cFormFields
	End Sub

	Private Sub Class_Terminate()
		If UploadID > 0 Then 
			ProgressFile.Contents = "DONE" & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
		End If
	End Sub

	Public Boundary
	Private m_Items 
	Private m_State

	Private bSourceData
	Private StartUploadTime , TempFiolder 
	Private ProgressFile
End Class

Class cFormFields
	Dim m_Keys()
	Dim m_Items()
	Dim m_NomeFile()
	Dim m_Count
	
	Public Default Property Get Item(ByVal Key)
		If vartype(Key) = vbInteger or vartype(Key) = vbLong then
			if Key<1 or Key>m_Count Then Err.raise "Index out of bounds"
			Set Item = m_Items(Key-1)
			Exit Property
		end if

		Dim Count
		Count = ItemCount(Key)
		Key = LCase(Key)
		
		If Count > 0 then
			If Count>1 Then
				Dim OutItem, ItemCounter
				Set OutItem = New cFormFields
				ItemCounter = 0
				
				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key then OutItem.Add Key, m_Items(ItemCounter)
				Next
				Set Item = OutItem
			Else 
				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key then exit for
				Next

				if isobject (m_Items(ItemCounter)) then
					Set Item = m_Items(ItemCounter)
				else
					Item = m_Items(ItemCounter)
				end if
			End If
		Else 
			Set Item = New cFormField
		End if
	End Property

	Public Property Get MultiItem(ByVal Key)
		Dim Out: Set Out = New cFormFields
		Dim I, vItem 
		Dim Count
		Count = ItemCount(Key)
		
		if Count = 1 then
			Out.Add Key, Item(Key)
		elseif Count > 1 then
			For Each I In Item(Key).Items
				Out.Add Key, I
			Next
		End If

		Set MultiItem = Out
	End Property

	Public Property Get Value
		Dim I, V
		For Each I in m_Items
			V = V & ", " & I 
		Next
		V = Mid(V, 3)
		Value = V
	End Property


	Public Property Get xA_NewEnum
		Set xA_NewEnum = m_Items
	End Property

	Public Property Get Items()
		Items = m_Items
	End Property

	Public Property Get Keys()
		Keys = m_Keys
	End Property
	
	Public Property Get FileNames()
		FileNames = m_NomeFile
	End Property
	
	Public Property Get ValueByName(str)
		cong = ""
		Ris = ""
		For ItemCounter = 0 To Ubound(m_Keys)
			if ucase(str) = ucase(m_Keys(ItemCounter)) then
				Ris = Ris & cong & m_Items(ItemCounter)
				cong = ","
			end if
		Next
		ValueByName = Ris
	End Property
	
	public Property Get getFilesCount
		Dim cItem, OutItem, ItemCounter, Num
		Set OutItem = New cFormFields 
		ItemCounter = 0
		Num = 0
		if m_Count > 0 then
			For ItemCounter = 0 To Ubound(m_Keys)
				Set cItem = m_Items(ItemCounter)
				if cItem.IsFile then
					Num = Num + 1
				end if
			Next
		End If
		getFilesCount = Num 
	End Property
	
	public Property Get getFileNameById(x)
		Dim cItem, OutItem, ItemCounter, Num, Ris
		Ris = ""
		Set OutItem = New cFormFields 
		ItemCounter = 0
		if m_Count > 0 then
			For ItemCounter = 0 To Ubound(m_Keys)
				Set cItem = m_Items(ItemCounter)
				if cItem.IsFile then
					Num = Num + 1
					if Num = x then
						Ris = RG_RemoveSpecialChar(m_NomeFile(ItemCounter))
						exit for
					end if
				end if
			Next
		End If
		getFileNameById = Ris 
	End Property
	
	public Property Get Files
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields 
		ItemCounter = 0
		if m_Count > 0 then
			For ItemCounter = 0 To Ubound(m_Keys)
				Set cItem = m_Items(ItemCounter)
				if cItem.IsFile then
					OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter), m_NomeFile(ItemCounter)
				end if
			Next
		End If
		Set Files = OutItem 
	End Property

	Public Property Get Texts
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields 
		ItemCounter = 0
		For ItemCounter = 0 To Ubound(m_Keys)
			Set cItem = m_Items(ItemCounter)
			if Not cItem.IsFile then
				OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter),""
			end if
		Next
		Set Texts = OutItem
	End Property

	Public Sub Save()
		Dim Item
		Dim UploadedSize
		UploadedSize = 0
		if TotalFileSize = "" then
			TotalFileSize = 0
		else
			TotalFileSize = clng(TotalFileSize)
		end if
		if TotalFileSize <> 0 then
			For Each Item In m_Items
				If Item.isFile Then
					UploadedSize = UploadedSize + cdbl(Item.Length)
				End If
			Next
		end if
		dim FileSave
		FileSave = true
		if UploadType = "AdvancedMail" and NoSendFileExceed = "1" then
			For Each Item In m_Items
				If Item.isFile Then
					'response.write "-i " & Item.Length & "<br>"
					'response.write "-s " & GetValByName("FileSize",Item.FieldName) & "<br>"
					if not RG_VerifyLen(Item.Length,GetValByName("FileSize",Item.FieldName)) then
						FileSave = false
						UploadMailError = "Files not send - One or more files exceed size limit"
						exit for
					end if
				End If
			Next
		end if
		if UploadType = "AdvancedMail" and NoSendFileExceed = "1" and TotalFileSize <> 0 and FileSave then
			if not RG_VerifyLen(UploadedSize, TotalFileSize) then
				FileSave = false
				UploadMailError = "Files not send - Exceed total size limit"
			end if
		end if
		'response.write "-" & UploadType & "<br>"
		'response.write "-" & NoSendFileExceed & "<br>"
		'response.write "-" & UploadMailError & "<br>"
		'response.write "-" & UploadedSize & "<br>"
		'response.write "-" & TotalFileSize & "<br>"
		
		'response.End()
		if FileSave then
			if TotalFileSize = 0 or UploadedSize <= (cdbl(TotalFileSize)*1000) then
				For Each Item In m_Items
					If Item.isFile Then
						Item.Save
					End If
				Next
			else
				UploadError = round(UploadedSize / 1000)
			end if
		end if
	End Sub
	
	Public Sub DataBaseUpdate()
		
		if RG_editQueryTmp<>"" then
			RG_editQueryTmp = ", " & RG_editQueryTmp
		end if
  
  		RG_fields = Split(RG_fieldsStr, "|")
  		RG_columns = Split(RG_columnsStr, "|")
  
		For RG_i = LBound(RG_fields) To UBound(RG_fields) Step 2
			RG_fields(RG_i+1) = CStr(UploadFormRequest(RG_fields(RG_i)))
		Next
		
		RG_editQuery = "update " & verFieldTableName(RG_editTable) & " set "
		For RG_i = LBound(RG_fields) To UBound(RG_fields) Step 2
			RG_formVal = RG_fields(RG_i+1)
			RG_typeArray = Split(RG_columns(RG_i+1),",")
			RG_delim = RG_typeArray(0)
			If (RG_delim = "none") Then RG_delim = ""
			RG_altVal = RG_typeArray(1)
			If (RG_altVal = "none") Then RG_altVal = ""
			RG_emptyVal = RG_typeArray(2)
			If (RG_emptyVal = "none") Then RG_emptyVal = ""
			If (RG_formVal = "") Then
			  RG_formVal = RG_emptyVal
			Else
			  If (RG_altVal <> "") Then
				RG_formVal = RG_altVal
			  ElseIf (RG_delim = "'") Then
				RG_formVal = "'" & Replace(RG_formVal,"'","''") & "'"
			  Else
				RG_formVal = RG_delim + RG_formVal + RG_delim
			  End If
			End If
			If (RG_i <> LBound(RG_fields)) Then
			  RG_editQuery = RG_editQuery & ","
			End If
			RG_editQuery = RG_editQuery & verFieldTableName(RG_columns(RG_i)) & " = " & RG_formVal
	  	Next
  		RG_editQuery = RG_editQuery & RG_editQueryTmp & " where " & verFieldTableName(RG_editColumn) & " = " & UploadFormRequest("RG_recordId")
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open RG_Connection
		on error resume next
		Conn.execute (RG_editQuery)
		if err.number<>0 then
			FirstError = err.Description
			err.clear
			RG_editQuery = replace(replace(RG_editQuery,"[",""),"]","")
			Conn.execute (RG_editQuery)
			if err.number<>0 then
				err.clear
				RG_editQuery = replace(replace(RG_editQuery,"[","`"),"]","`")
				Conn.execute (RG_editQuery)
				if err.number<>0 then
					response.write "I find an error in the sql:<br>" & replace(replace(RG_editQuery,"`",""),"`","") & "<br>" & "I find this error: " & FirstError
					response.end
				end if
			end if
		end if
		on error goto 0
		
		'DeleteExistingFiles
		for x=0 to MaxFieldNumber
			if ParamVal(right("00" & cstr(x),3) & "14") <> "" then
				if UploadFormRequest(ParamVal(right("00" & cstr(x),3) & "14")) <> "" and CStr(ParamVal(right("00" & cstr(x),3) & "10")) <> "" then
					SQL = "SELECT " & verFieldTableName(CStr(ParamVal(right("00" & cstr(x),3) & "10"))) & " FROM " & verFieldTableName(RG_editTable) & " where " & verFieldTableName(RG_editColumn) & " = " & UploadFormRequest("RG_recordId")
					Set RG_Rec = Server.CreateObject("ADODB.RecordSet")
					RG_Rec.open SQL,Conn
					RG_FileDel = RG_Rec(CStr(ParamVal(right("00" & cstr(x),3) & "10")))
					RG_Rec.close
					SQL = "UPDATE " & verFieldTableName(RG_editTable) & " SET " & verFieldTableName(CStr(ParamVal(right("00" & cstr(x),3) & "10"))) & " = Null where " & verFieldTableName(RG_editColumn) & " = " & UploadFormRequest("RG_recordId")
					Conn.execute(SQL)
					set RG_Rec = nothing
					set RG_FS = server.CreateObject("Scripting.FileSystemObject")
					if CStr(RG_FileDel & "") <> "" then
						if instr(RG_FileDel,":") > 0 then
							if RG_FS.fileExists(RG_FileDel) then
								RG_FS.DeleteFile(RG_FileDel)
							end if
						else
							if RG_FS.fileExists(server.MapPath(RG_FileDel)) then
								RG_FS.DeleteFile(server.MapPath(RG_FileDel))
							end if
						end if						
					end if
					set RG_FS = nothing
				end if
			end if
		next
		
		Conn.close
		Set Conn = nothing
 			
	end sub
	

	
	Public Sub DataBaseInsert()
		
		RG_fields = Split(RG_fieldsStr, "|")
		RG_columns = Split(RG_columnsStr, "|")
		
		For RG_i = LBound(RG_fields) To UBound(RG_fields) Step 2
			RG_fields(RG_i+1) = CStr(UploadFormRequest(RG_fields(RG_i)))
		Next
		
		RG_tableValues = ""
		RG_dbValues = ""
		For RG_i = LBound(RG_fields) To UBound(RG_fields) Step 2
			RG_formVal = RG_fields(RG_i+1)
			RG_typeArray = Split(RG_columns(RG_i+1),",")
			RG_delim = RG_typeArray(0)
			If (RG_delim = "none") Then RG_delim = ""
			RG_altVal = RG_typeArray(1)
			If (RG_altVal = "none") Then RG_altVal = ""
			RG_emptyVal = RG_typeArray(2)
			If (RG_emptyVal = "none") Then RG_emptyVal = ""
			If (RG_formVal = "") Then
		  		RG_formVal = RG_emptyVal
			Else
			  If (RG_altVal <> "") Then
				RG_formVal = RG_altVal
			  ElseIf (RG_delim = "'") Then  
				RG_formVal = "'" & Replace(RG_formVal,"'","''") & "'"
			  Else
				RG_formVal = RG_delim + RG_formVal + RG_delim
			  End If
			End If
			If (RG_i <> LBound(RG_fields)) Then
			  RG_tableValues = RG_tableValues & ","
			  RG_dbValues = RG_dbValues & ","
			End If
			RG_tableValues = RG_tableValues & verFieldTableName(RG_columns(RG_i))
			RG_dbValues = RG_dbValues & RG_formVal
		Next
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open RG_Connection
	
		if isSingleRecord() then
			RG_tableValuesTmp = ""
			RG_dbValuesTmp = ""
			cong = ","
			for J=1 to QtyRecord
				if tmpField_Name(J)<>"" then
					RG_tableValuesTmp = RG_tableValuesTmp & cong & tmpField_Name(J)
					RG_dbValuesTmp = RG_dbValuesTmp & cong & tmpValue_Name(J)
				end if
				if tmpField_Size(J)<>"" then
					RG_tableValuesTmp = RG_tableValuesTmp & cong & tmpField_Size(J)
					RG_dbValuesTmp = RG_dbValuesTmp & cong & tmpValue_Size(J)
				end if
				if tmpField_Thumb(J)<>"" then
					RG_tableValuesTmp = RG_tableValuesTmp & cong & tmpField_Thumb(J)
					RG_dbValuesTmp = RG_dbValuesTmp & cong & tmpValue_Thumb(J)
				end if
			next
			RG_editQuery = "insert into " & verFieldTableName(RG_editTable) & " (" & RG_tableValues & RG_tableValuesTmp & ") values (" & RG_dbValues & RG_dbValuesTmp & ")"
			'response.write RG_editQuery
			'response.End()
			on error resume next
			Conn.execute (RG_editQuery)
			if err.number<>0 then
				FirstError = err.Description
				err.clear
				RG_editQuery = replace(replace(RG_editQuery,"[",""),"]","")
				Conn.execute (RG_editQuery)
				if err.number<>0 then
					err.clear
					RG_editQuery = replace(replace(RG_editQuery,"[","`"),"]","`")
					Conn.execute (RG_editQuery)
					if err.number<>0 then
						response.write "I find an error in the sql:<br>" & replace(replace(RG_editQuery,"`",""),"`","")  & "<br>" & "I find this error: " & FirstError
						response.end
					end if
				end if
			end if
			on error goto 0
		end if
		Conn.close
		Set Conn = nothing			

	End Sub

	Public Property Get ItemCount(ByVal Key)
		Dim cKey, Counter
		Counter = 0
		Key = LCase(Key)
		For Each cKey In m_Keys
			If LCase(cKey) = Key then Counter = Counter + 1
		Next
		ItemCount = Counter
	End Property

	Public Property Get Count()
		Count = m_Count
	End Property

	Public Sub Add(byval Key, Item, FName)
		Key = "" & Key
		ReDim Preserve m_Items(m_Count)
		ReDim Preserve m_Keys(m_Count)
		ReDim Preserve m_NomeFile(m_Count)
		m_Keys(m_Count) = Key
		Set m_Items(m_Count) = Item
		m_NomeFile(m_Count) = FName
		m_Count = m_Count + 1
	End Sub

	Private Sub Class_Initialize()
		Dim vHelp()
		On Error Resume Next
		m_Items = vHelp
		m_Keys = vHelp
		m_NomeFile = vHelp
		m_Count = 0
	End Sub

	Public Sub mpSeparateFields(Binary, ByVal Boundary)
		Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary

		Boundary = "--" & Boundary			
		Boundary = StringToBinary(Boundary)

		PosOpenBoundary = InStrB(Binary, Boundary)
		PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

		Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)

			Dim HeaderContent, bFieldContent

			Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type

			Dim TwoCharsAfterEndBoundary
			PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

			HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
    
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)

			GetHeadFields BinaryToString(HeaderContent), FormFieldName, SourceFileName, Content_Disposition, Content_Type

			Dim Field
			Set Field = New cFormField

			Field.ByteArray = MultiByteToBinary(bFieldContent)

			Field.Name = FormFieldName
			Field.ContentDisposition = Content_Disposition
			if not isempty(SourceFileName) then
				Field.FilePath = SourceFileName
				Field.FileName = GetFileName(SourceFileName)
			End If
			Field.ContentType = Content_Type
	
			Add FormFieldName, Field, GetFileName(SourceFileName)

			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
			isLastBoundary = TwoCharsAfterEndBoundary = "--"
			
			If Not isLastBoundary Then 
				PosOpenBoundary = PosCloseBoundary
				PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
			End If
		Loop
	End Sub
End Class




Class cProgressFile
	Private fs
	'Public TempFolder
	Public m_UploadID
	Public TempFileName

	Public Default Property Get Contents()
		Contents = GetFile(TempFileName)
	End Property

	Public Property Let Contents(inContents)
		WriteFile TempFileName, inContents
	End Property

	Public Sub Done
		'FS.DeleteFile TempFileName
	End Sub

	Public Property Get UploadID()
		UploadID = m_UploadID
	End Property

	Public Property Let UploadID(inUploadID)
		if isempty(FS) then Set fs = server.CreateObject("Scripting.FileSystemObject")
		'TempFolder = fs.GetSpecialFolder(2)

		m_UploadID = inUploadID
		TempFileName = getTempFolder() & "\pu" & m_UploadID & ".~tmp"
		
		Dim DateLastModified
		on error resume next
		DateLastModified = fs.GetFile(TempFileName).DateLastModified
		on error goto 0
		if isempty(DateLastModified) then
		elseif Now-DateLastModified>1 Then
			FS.DeleteFile TempFileName
		end if
	End Property

	Private Function GetFile(Byref FileName)
		Dim InStream
		On Error Resume Next
		Set InStream = fs.OpenTextFile(FileName, 1)
		GetFile = InStream.ReadAll
		On Error Goto 0
	End Function

	Private Function WriteFile(Byref FileName, Byref Contents)
		Dim OutStream
		On Error Resume Next
		Set OutStream = fs.OpenTextFile(FileName, 2, True)
		OutStream.Write Contents
	End Function


	Private Sub Class_Initialize()
	End Sub
End Class




Class cFormField

	Public ContentDisposition, ContentType, FileName, FilePath, Name
	Public ByteArray

	Public CharSet, HexString, InProgress, SourceLength, RAWHeader, Index, ContentTransferEncoding
 
	Public Default Property Get String()
		String = BinaryToString(ByteArray)
	End Property 

	Public Property Get IsFile()
		if len(FileName) = 0 then
			IsFile = false
		else
			IsFile = true
		end if
	End Property

	Public Property Get FieldName()
		FieldName = Name
	End Property
	
	Public Property Get Length()
		Length = LenB(ByteArray)
	End Property

	Public Property Get Value()
		Set Value = Me
	End Property
	
	Public Sub Save()
		if IsFile Then
			StatusFile = "0"
			NewFileName = "0"
			VerLen = RG_VerifyLen(Length,GetValByName("FileSize",Name))
			VerExt = RG_VerifyExt(RG_RemoveSpecialChar(FileName),GetValByName("AllowedExtension",Name))
			if VerLen and VerExt then
				if CStr(GetValByName("FolderType",Name)) = "1"  then
					RG_Path = server.MapPath(CStr(GetValByName("UploadFolder",Name)))
					if right(RG_Path,1)<>"\" then
						RG_Path = RG_Path + "\"
					end if
				elseif CStr(GetValByName("FolderType",Name)) = "2"  then
					RG_Path = CStr(GetValByName("UploadFolder",Name))
					if right(RG_Path,1)<>"\" then
						RG_Path = RG_Path + "\"
					end if
				else
					RG_Path = GetFolderName(CStr(Name))
					if right(RG_Path,1)<>"\" then
						RG_Path = RG_Path + "\"
					end if
				end if
				if CStr(GetValByName("FileName",Name))<>"" and not isnull(CStr(GetValByName("FileName",Name))) or CStr(GetValByName("FileNameFormula",Name))="-1"  then
					if CStr(GetValByName("FileNameFormula",Name))="-1" then
						RG_Name = myGetFileName(Name) & mid(FileName,instrRev(FileName,"."))
					else
						RG_Name = CStr(GetValByName("FileName",Name))
					end if
				else
					RG_Name = RG_RemoveSpecialChar(FileName)
				end if

				if UploadType="Update" then
					setValByName "DeleteExistFile", Name, ""
					if CStr(GetValByName("DelExist",Name)) = "1" and CStr(GetValByName("FieldExistField",Name)) <> "" and not isnull(CStr(GetValByName("FieldExistField",Name))) then 
						SQL = "SELECT " & verFieldTableName(CStr(CStr(GetValByName("FieldExistField",Name)))) & " FROM " & verFieldTableName(RG_editTable) & " where " & verFieldTableName(RG_editColumn) & " = " & UploadFormRequest("RG_recordId")
						Set Conn = Server.CreateObject("AdoDb.Connection")
						Set RG_Rec = Server.CreateObject("ADODB.RecordSet")
						Conn.Open RG_Connection
						RG_Rec.open SQL,Conn
						RG_FileDel = RG_Rec(CStr(GetValByName("FieldExistField",Name)))
						RG_Rec.close
						SQL = "UPDATE " & verFieldTableName(RG_editTable) & " SET " & verFieldTableName(CStr(GetValByName("FieldExistField",Name))) & " = Null where " & verFieldTableName(RG_editColumn) & " = " & UploadFormRequest("RG_recordId")
						Conn.execute(SQL)
						Conn.close
						set Conn = nothing
						set RG_Rec = nothing
						set RG_FS = server.CreateObject("Scripting.FileSystemObject")
						if CStr(RG_FileDel & "") <> "" then
							if instr(RG_FileDel,":") > 0 then
								if RG_FS.fileExists(RG_FileDel) then
									RG_FS.DeleteFile(RG_FileDel)
								end if
							else
								if RG_FS.fileExists(server.MapPath(RG_FileDel)) then
									RG_FS.DeleteFile(server.MapPath(RG_FileDel))
								end if
							end if						
						end if
						set RG_FS = nothing
					end if
				elseif UploadType="AdvancedMail" and isAdvMailUpdate="-1" then
					if CStr(GetValByName_Adv("DelExist",Name)) = "1" and CStr(GetValByName_Adv("FieldExistField",Name)) <> "" and not isnull(CStr(GetValByName_Adv("FieldExistField",Name))) then 
						SQL = "SELECT " & verFieldTableName(CStr(CStr(GetValByName_Adv("FieldExistField",Name)))) & " FROM " & verFieldTableName(getAdvMailPlugInParameter("AdvMail_editTable")) & " where " & verFieldTableName(getAdvMailPlugInParameter("AdvMail_editColumn")) & " = " & getAdvMailPlugInParameter("AdvMail_ColQuote") & UploadFormRequest("RG_recordId") & getAdvMailPlugInParameter("AdvMail_ColQuote")
						Set Conn = Server.CreateObject("AdoDb.Connection")
						Set RG_Rec = Server.CreateObject("ADODB.RecordSet")
						Conn.Open getAdvMailPlugInParameter("AdvMail_editConnection")
						RG_Rec.open SQL,Conn
						RG_FileDel = RG_Rec(CStr(GetValByName_Adv("FieldExistField",Name)))
						RG_Rec.close
						SQL = "UPDATE " & verFieldTableName(getAdvMailPlugInParameter("AdvMail_editTable")) & " SET " & verFieldTableName(CStr(GetValByName_Adv("FieldExistField",Name))) & " = Null where " & verFieldTableName(getAdvMailPlugInParameter("AdvMail_editColumn")) & " = " & getAdvMailPlugInParameter("AdvMail_ColQuote") & UploadFormRequest("RG_recordId") & getAdvMailPlugInParameter("AdvMail_ColQuote")
						Conn.execute(SQL)
						Conn.close
						set Conn = nothing
						set RG_Rec = nothing
						set RG_FS = server.CreateObject("Scripting.FileSystemObject")
						if CStr(RG_FileDel & "") <> "" then
							if instr(RG_FileDel,":") > 0 then
								if RG_FS.fileExists(RG_FileDel) then
									RG_FS.DeleteFile(RG_FileDel)
								end if
							else
								if RG_FS.fileExists(server.MapPath(RG_FileDel)) then
									RG_FS.DeleteFile(server.MapPath(RG_FileDel))
								end if
							end if						
						end if
						set RG_FS = nothing
					end if
				end if
								
				RG_Save = 0
				
				set RG_FS = server.createObject( "Scripting.FileSystemObject")
				if RG_FS.fileExists(RG_Path & RG_Name) then
					select case CStr(GetValByName("Conflict",Name))
					case "0"
						SaveAs RG_Path & RG_Name
						RG_Save = -1
					case "1" 
						RG_Num = -1
						do while 1
							RG_Num = RG_Num + 1
							RG_newName = left(RG_Name,instrRev(RG_Name,".")-1) & RG_Num & "." & mid(RG_Name,instrRev(RG_Name,".")+1)
							if not RG_FS.fileExists(RG_Path & RG_newName) then
								RG_Name = RG_newName
								NewFileName = "-1"
								exit do
							end if
						loop
						SaveAs RG_Path & RG_Name
						RG_Save = -1
					case "2"
						StatusFile = "4"
					end select
				else
					SaveAs RG_Path & RG_Name
					RG_Save = -1
				end if
				
				'Resize
				if RG_Save and GetValByNameResize("Resize",Name) = "-1" then
					select case getComponent()
					case 1
						if resizeAspImage(RG_Path & RG_Name, Name, "0") then
							if lcase(right(RG_Name,4)) = "jpeg" then
								RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
							end if
						end if
					case 2
						if resizeAspJpeg(RG_Path & RG_Name, Name, "0") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
						end if
					case 3
						if resizeAspSmartImage(RG_Path & RG_Name, Name, "0") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
						end if
					case 4
						if resizeAspThumb(RG_Path & RG_Name, Name, "0") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
						end if
					case 5
						resizeImgWriter RG_Path & RG_Name, Name, "0"
					case 6
						if resizeVije(RG_Path & RG_Name, Name, "0") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
						end if
					case 7
						resizeGraphProc RG_Path & RG_Name, Name, "0"
					case 8
						resizeGraphMill RG_Path & RG_Name, Name, "0"
					case 9
						resizeChestySoftCOM RG_Path & RG_Name, Name, "0"
					case 10
						resizeChestySoftOCX RG_Path & RG_Name, Name, "0"
					case 11
						resizeXnView RG_Path & RG_Name, Name, "0"
					case 12
						resizeXnView193 RG_Path & RG_Name, Name, "0"
					case 13
						resizeAspNet RG_Path & RG_Name, Name, "0"
					end select
				end if
				img_Ext = ""
				if RG_Save and GetValByNameResize("Thumb",Name) = "-1" then
					select case getComponent()
					case 1
						if resizeAspImage(RG_Path & RG_Name, Name, "-1") then
							if lcase(right(RG_Name,4)) = "jpeg" then
								RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
							end if
							img_Ext = "jpg"
						end if
					case 2
						if resizeAspJpeg(RG_Path & RG_Name, Name, "-1") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
							img_Ext = "jpg"
						end if
					case 3
						if resizeAspSmartImage(RG_Path & RG_Name, Name, "-1") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
							img_Ext = "jpg"
						end if						
					case 4
						if resizeAspThumb(RG_Path & RG_Name, Name, "-1") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
							imgExt = "jpg"
						end if						
					case 5
						resizeImgWriter RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 6
						if resizeVije(RG_Path & RG_Name, Name, "-1") then
							RG_Name = left(RG_Name,instrRev(RG_Name,".")) & "jpg"
						end if
						img_Ext = "jpg"
					case 7
						resizeGraphProc RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 8
						resizeGraphMill RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 9
						resizeChestySoftCOM RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 10
						resizeChestySoftOCX RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 11
						resizeXnView RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 12
						resizeXnView193 RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					case 13
						resizeAspNet RG_Path & RG_Name, Name, "-1"
						img_Ext = ""
					end select
				end if
				if img_Ext = ""  Then
					img_Ext = lcase(mid(RG_Name,instrRev(RG_Name,".")+1))
				end if
				DW_ThumbName = getNewFileName(RG_Name, "-1", GetValByNameResize("Suffix",Name), img_Ext, GetValByNameResize("SuffixPosition",Name))
								
				set RG_FS = nothing
				
				if UploadType="Update" then
					if RG_Save and GetValByName("FieldToSaveName",Name)<>"" then
						if RG_editQueryTmp<>"" then
							RG_Cong = ", "
						end if
						if CStr(GetValByName("SaveType",Name)) = "1" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveName",Name)) & " = '" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & RG_Name & "'"
						elseif CStr(GetValByName("SaveType",Name)) = "2" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveName",Name)) & " = '" & RG_Path & RG_Name & "'"
						else
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveName",Name)) & " = '" & RG_Name & "'"
						end if
						RG_Cong = ", "
					end if

					if RG_Save and GetValByName("FieldToSaveSize",Name)<>"" then
						if right(GetValByName("FieldToSaveSize",Name),1)="S" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong &  verFieldTableName(left(GetValByName("FieldToSaveSize",Name),len(GetValByName("FieldToSaveSize",Name))-2)) & " = '" & getFileSize(RG_Path & RG_Name) & "'"
						else
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(left(GetValByName("FieldToSaveSize",Name),len(GetValByName("FieldToSaveSize",Name))-2)) & " = " & getFileSize(RG_Path & RG_Name)
						end if
						RG_Cong = ", "
					end if
					
					if RG_Save and GetValByName("FieldToSaveThumb",Name)<>"" then
						if RG_editQueryTmp<>"" then
							RG_Cong = ", "
						end if
						if CStr(GetValByName("ThumbSaveType",Name)) = "1" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveThumb",Name)) & " = '" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & DW_ThumbName & "'"
						elseif CStr(GetValByName("ThumbSaveType",Name)) = "2" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveThumb",Name)) & " = '" & RG_Path & DW_ThumbName & "'"
						else
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName("FieldToSaveThumb",Name)) & " = '" & DW_ThumbName & "'"
						end if
						RG_Cong = ", "
					end if
				elseif UploadType="Insert" then
					ActualIndex = QtyRecord
					if RG_Save and GetValByName("FieldToSaveName",Name)<>"" then
						QtyRecord = QtyRecord + 1
						ReDim Preserve tmpField_Name(QtyRecord)
						ReDim Preserve tmpValue_Name(QtyRecord)
						ReDim Preserve tmpField_Size(QtyRecord)
						ReDim Preserve tmpValue_Size(QtyRecord)
						ReDim Preserve tmpField_Thumb(QtyRecord)
						ReDim Preserve tmpValue_Thumb(QtyRecord)
						tmpField_Name(QtyRecord) = verFieldTableName(GetValByName("FieldToSaveName",Name))
						if CStr(GetValByName("SaveType",Name)) = "1" then
							tmpValue_Name(QtyRecord) = "'" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & RG_Name & "'"
						elseif CStr(GetValByName("SaveType",Name)) = "2" then
							tmpValue_Name(QtyRecord) = "'" & RG_Path & RG_Name & "'"
						else
							tmpValue_Name(QtyRecord) = "'" & RG_Name & "'"
						end if
						
						if GetValByName("FieldToSaveThumb",Name) <> "" then
							tmpField_Thumb(QtyRecord) = verFieldTableName(GetValByName("FieldToSaveThumb",Name))
							if CStr(GetValByName("ThumbSaveType",Name)) = "1" then
								tmpValue_Thumb(QtyRecord) = "'" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & DW_ThumbName & "'"
							elseif CStr(GetValByName("SaveType",Name)) = "2" then
								tmpValue_Thumb(QtyRecord) = "'" & RG_Path & DW_ThumbName & "'"
							else
								tmpValue_Thumb(QtyRecord) = "'" & DW_ThumbName & "'"
							end if
						end if
					end if
					if RG_Save and GetValByName("FieldToSaveSize",Name)<>"" then
						if ActualIndex = QtyRecord then
							QtyRecord = QtyRecord + 1
							ReDim Preserve tmpField_Name(QtyRecord)
							ReDim Preserve tmpValue_Name(QtyRecord)
							ReDim Preserve tmpField_Size(QtyRecord)
							ReDim Preserve tmpValue_Size(QtyRecord)
							ReDim Preserve tmpField_Thumb(QtyRecord)
							ReDim Preserve tmpValue_Thumb(QtyRecord)
						end if
						tmpField_Size(QtyRecord) = verFieldTableName(left(GetValByName("FieldToSaveSize",Name),len(GetValByName("FieldToSaveSize",Name))-2))
						if right(GetValByName("FieldToSaveSize",Name),1)="S" then
							tmpValue_Size(QtyRecord) = "'" & getFileSize(RG_Path & RG_Name) & "'"
						else
							tmpValue_Size(QtyRecord) = getFileSize(RG_Path & RG_Name)
						end if
					end if
				elseif UploadType="AdvancedMail" and isAdvMailInsert="-1" then
					ActualIndex = QtyRecord
					if RG_Save and GetValByName_Adv("FieldToSaveName",Name)<>"" then
						QtyRecord = QtyRecord + 1
						ReDim Preserve tmpField_Name(QtyRecord)
						ReDim Preserve tmpValue_Name(QtyRecord)
						ReDim Preserve tmpField_Size(QtyRecord)
						ReDim Preserve tmpValue_Size(QtyRecord)
						ReDim Preserve tmpField_Thumb(QtyRecord)
						ReDim Preserve tmpValue_Thumb(QtyRecord)
						tmpField_Name(QtyRecord) = verFieldTableName(GetValByName_Adv("FieldToSaveName",Name))
						if CStr(GetValByName_Adv("SaveType",Name)) = "1" then
							tmpValue_Name(QtyRecord) = "'" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & RG_Name & "'"
						elseif CStr(GetValByName_Adv("SaveType",Name)) = "2" then
							tmpValue_Name(QtyRecord) = "'" & RG_Path & RG_Name & "'"
						else
							tmpValue_Name(QtyRecord) = "'" & RG_Name & "'"
						end if
					end if
					if RG_Save and GetValByName_Adv("FieldToSaveSize",Name)<>"" then
						if ActualIndex = QtyRecord then
							QtyRecord = QtyRecord + 1
							ReDim Preserve tmpField_Name(QtyRecord)
							ReDim Preserve tmpValue_Name(QtyRecord)
							ReDim Preserve tmpField_Size(QtyRecord)
							ReDim Preserve tmpValue_Size(QtyRecord)
							ReDim Preserve tmpField_Thumb(QtyRecord)
						ReDim Preserve tmpValue_Thumb(QtyRecord)
						end if
						tmpField_Size(QtyRecord) = verFieldTableName(left(GetValByName_Adv("FieldToSaveSize",Name),len(GetValByName_Adv("FieldToSaveSize",Name))-2))
						if right(GetValByName_Adv("FieldToSaveSize",Name),1)="S" then
							tmpValue_Size(QtyRecord) = "'" & getFileSize(RG_Path & RG_Name) & "'"
						else
							tmpValue_Size(QtyRecord) = getFileSize(RG_Path & RG_Name)
						end if
					end if
				elseif UploadType="AdvancedMail" and isAdvMailUpdate="-1" then
					if RG_Save and GetValByName_Adv("FieldToSaveName",Name)<>"" then
						if RG_editQueryTmp<>"" then
							RG_Cong = ", "
						end if
						if CStr(GetValByName_Adv("SaveType",Name)) = "1" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName_Adv("FieldToSaveName",Name)) & " = '" & replace(replace(replace(RG_Path,"/","\"),server.MapPath("/"),"",1,1,1),"\","/") & RG_Name & "'"
						elseif CStr(GetValByName_Adv("SaveType",Name)) = "2" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName_Adv("FieldToSaveName",Name)) & " = '" & RG_Path & RG_Name & "'"
						else
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(GetValByName_Adv("FieldToSaveName",Name)) & " = '" & RG_Name & "'"
						end if
						RG_Cong = ", "
					end if

					if RG_Save and GetValByName_Adv("FieldToSaveSize",Name)<>"" then
						if right(GetValByName_Adv("FieldToSaveSize",Name),1)="S" then
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong &  verFieldTableName(left(GetValByName_Adv("FieldToSaveSize",Name),len(GetValByName_Adv("FieldToSaveSize",Name))-2)) & " = '" & getFileSize(RG_Path & RG_Name) & "'"
						else
							RG_editQueryTmp = RG_editQueryTmp & RG_Cong & verFieldTableName(left(GetValByName_Adv("FieldToSaveSize",Name),len(GetValByName_Adv("FieldToSaveSize",Name))-2)) & " = " & getFileSize(RG_Path & RG_Name)
						end if
						RG_Cong = ", "
					end if
				end if
			else
				if not VerLen and not VerExt then
					StatusFile = "3"
				elseif VerLen and not VerExt then
					StatusFile = "2"
				elseif not VerLen and VerExt then
					StatusFile = "1"
				end if
			end if
			
			if RG_Name = "" then
				RG_Name = FileName
			end if
			if StatusFile<>"0" then
				FileMatch = "0"
			end if
			
			NumFile = NumFile + 1
			UploadStatus = UploadStatus & "<input name=" & chr(34) & "FileName_" & NumFile & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & RG_Name & chr(34) & ">"
			UploadStatus = UploadStatus & "<input name=" & chr(34) & "FileSize_" & NumFile & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & Length & chr(34) & ">"
			UploadStatus = UploadStatus & "<input name=" & chr(34) & "FileStatus_" & NumFile & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & StatusFile & chr(34) & ">"
			UploadStatus = UploadStatus & "<input name=" & chr(34) & "NewFileName_" & NumFile & chr(34) & " type=" & chr(34) & "hidden" & chr(34) & " value=" & chr(34) & NewFileName & chr(34) & ">"	
		End If
	End Sub

	Public Sub SaveAs(newFilePath)
		if len(ByteArray)>0 then SaveBinaryData newFilePath, ByteArray
	End Sub
	
End Class



'****************************
'                           *
'  JavaScript Functions     *
'                           *
'****************************
if ThisPage <> "Progress" then
%>

<script language="JavaScript">
var GoUpload
var imgUpload = new Image();
function ProgressBar(){
	GoUpload = true
	var FileFind = false
	var ListVal = "<%=trim(replace(left(RG_Files,instr(RG_Files,"@_@_@")-1),"\","/"))%>"
	if(ListVal!=""){
		var tmp = ListVal.split("@_@_@")
		FieldListVal = tmp[0].split("|")
		for(i=0;i<FieldListVal.length;i++){
			FieldVal = FieldListVal[i].split(";")
			//var FileName = "file:///"+<%=RG_formName%>.elements[FieldVal[<%=IndexFieldName%>]].value;
			var FileName = "file:///"+Dwz_findObj(FieldVal[<%=IndexFieldName%>]).value
			if(FileName!="file:///"){
				var FileFind = true
				FileObj(FileName,FieldVal[12],FieldVal[13],FieldVal[4],FieldVal[3]);
				FileExt = FileName.substring(FileName.lastIndexOf(".")).toLowerCase()
				if((FileExt==".gif") || (FileExt==".jpg") || (FileExt==".jpeg") || (FileExt==".png") || (FileExt==".bmp")){ 
					t=0; while(imgUpload.fileSize==-1){t+=1; if(t>1000000){break}}
				}
				VerifyFile()
			}
		}
	}
	if(!GoUpload){
		alert("Upload Aborted")
		return false
	}
		
	<%if ProgressBar<>"" then%>
	if(FileFind){
		var ProgressURL
		var PosL = (screen.width/2)-150
		var PosT = (screen.height/2)-110
		ProgressURL = '<%=getProgressPath()%>UploadFiles/FrameSet.asp?ProgressPage=<%=ProgressBar%>&UploadID=<%=UploadID%>'
		var v = window.open(ProgressURL,'progressBarPage','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=350,height=210,top='+PosT+',left='+PosL)
	}
	<%end if%>
	return true;
}

function FileObj(Path,maxW,maxH,maxSize,extension){
	imgUpload.src=Path;
	imgUpload.maxW=maxW;
	imgUpload.maxH=maxH;
	imgUpload.maxSize=maxSize;
	imgUpload.extension=extension;
}

function VerifyFile(){
	var errorMsg="";
	//Verify extension
	if(imgUpload.extension!=""){
		imgUpload.extension = imgUpload.extension.toLowerCase()
		ext = imgUpload.extension.split(",")
		var ris = false
		for(g=0;g<ext.length;g++){
			thisExt = imgUpload.src.substring(imgUpload.src.length-ext[g].length)
			if(thisExt.toLowerCase()==ext[g]){
				ris = true
				break
			}
		}
		if(ris==false){
			errorMsg+="File extension allowed for this file is: "+imgUpload.extension+" Actual extension is "+imgUpload.src.substring(imgUpload.src.lastIndexOf("."))+"\n";
		}
	}
	//Verisy Size
	thisExt = imgUpload.src.substring(imgUpload.src.lastIndexOf(".")).toLowerCase()
	if(imgUpload.maxSize!=""&&((thisExt==".gif") || (thisExt==".bmp") || ((thisExt==".jpg")) || ((thisExt==".jpeg")) || ((thisExt==".png")))&&parseInt(imgUpload.fileSize/1000)>parseInt(imgUpload.maxSize)){	
		errorMsg+="Maximum size allowed for this file is: "+imgUpload.maxSize+" Kb - Actual size is "+parseInt(imgUpload.fileSize/1000)+" Kb\n";}
	//Verify With
	if(imgUpload.maxW!=""&&((thisExt==".gif") || (thisExt==".bmp") || ((thisExt==".jpg")) || ((thisExt==".jpeg")) || ((thisExt==".png")))&&imgUpload.width>imgUpload.maxW){
		errorMsg+="Maximum width allowed for this image is: "+imgUpload.maxW+" pixels - Actual width is "+imgUpload.width+" pixels\n";
	}
	//verify height
	if(imgUpload.maxH!=""&&((thisExt==".gif") || (thisExt==".bmp") || ((thisExt==".jpg")) || ((thisExt==".jpeg")) || ((thisExt==".png")))&&imgUpload.height>imgUpload.maxH){
		errorMsg+="Maximum height allowed for this image is: "+imgUpload.maxH+" pixels - Actual height is "+imgUpload.height+" pixels\n";
	}		
	if(errorMsg!=""){alert(imgUpload.src.replace("file:///","")+"\n\n"+errorMsg);}
	if(errorMsg!=""){GoUpload=false}
}


// Example: obj = findObj("image1");
function Dwz_findObj(theObj, theDoc)
{
  var p, i, foundObj;
  
  if(!theDoc) theDoc = document;
  if( (p = theObj.indexOf("?")) > 0 && parent.frames.length)
  {
    theDoc = parent.frames[theObj.substring(p+1)].document;
    theObj = theObj.substring(0,p);
  }
  if(!(foundObj = theDoc[theObj]) && theDoc.all) foundObj = theDoc.all[theObj];
  for (i=0; !foundObj && i < theDoc.forms.length; i++) 
    foundObj = theDoc.forms[i][theObj];
  for(i=0; !foundObj && theDoc.layers && i < theDoc.layers.length; i++) 
    foundObj = findObj(theObj,theDoc.layers[i].document);
  if(!foundObj && document.getElementById) foundObj = document.getElementById(theObj);
  
  return foundObj;
}

</script> 
<%end if%>