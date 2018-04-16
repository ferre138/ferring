<%

' ----------------------------------------------
'  ASPMaker 6 common variables (begin)
'
' used by system generated functions

Dim sSqlWrk, sWhereWrk, rswrk, ari, arwrk, armultiwrk, rowswrk, rowcntwrk, selwrk, jswrk

' used by header.asp, export checking
Dim sExport, sExportFile

'
'  ASPMaker 6 common variables (end)
' ----------------------------------------------
'
' --------------------------------------
'  ASPMaker 6 Email class (begin)
'
Class cEmail

	' Class properties
	Dim Sender ' Sender
	Dim Recipient ' Recipient
	Dim Cc ' Cc
	Dim Bcc ' Bcc
	Dim Subject ' Subject
	Dim Format ' Format
	Dim Content ' Content

	' Method to load email from template
	Public Sub Load(fn)
		Dim sWrk, sHeader, arrHeader
		Dim sName, sValue
		Dim i, j
		sWrk = ew_LoadTxt(fn) ' Load text file content
		sWrk = Replace(sWrk, vbCrLf, vbLf) ' Convert to Lf
		sWrk = Replace(sWrk, vbCr, vbLf) ' Convert to Lf
		If sWrk <> "" Then

			' Locate Header & Mail Content
			i = InStr(sWrk, vbLf&vbLf)
			If i > 0 Then
				sHeader = Mid(sWrk, 1, i)
				Content = Mid(sWrk, i+2)
				arrHeader = Split(sHeader, vbLf)
				For j = 0 to UBound(arrHeader)
					i = InStr(arrHeader(j), ":")
					If i > 0 Then
						sName = Trim(Mid(arrHeader(j), 1, i-1))
						sValue = Trim(Mid(arrHeader(j), i+1))
						Select Case LCase(sName)
							Case "subject"
								Subject = sValue
							Case "from"
								Sender = sValue
							Case "to"
								Recipient = sValue
							Case "cc"
								Cc = sValue
							Case "bcc"
								Bcc = sValue
							Case "format"
								Format = sValue
						End Select
					End If
				Next
			End If
		End If
	End Sub

	' Method to replace sender
	Public Sub ReplaceSender(ASender)
		Sender = Replace(Sender, "<!--$From-->", ASender)
	End Sub

	' Method to replace recipient
	Public Sub ReplaceRecipient(ARecipient)
		Recipient = Replace(Recipient, "<!--$To-->", ARecipient)
	End Sub

	' Method to add Cc email
	Public Sub AddCc(ACc)
		If ACc <> "" Then
			If Cc <> "" Then Cc = Cc & ";"
			Cc = Cc & ACc
		End If
	End Sub

	' Method to add Bcc email
	Public Sub AddBcc(ABcc)
		If ABcc <> "" Then
			If Bcc <> "" Then Bcc = Bcc & ";"
			Bcc = Bcc & ABcc
		End If
	End Sub

	' Method to replace subject
	Public Sub ReplaceSubject(ASubject)
		Subject = Replace(Subject, "<!--$Subject-->", ASubject)
	End Sub

	' Method to replace content
	Public Sub ReplaceContent(Find, ReplaceWith)
		Content = Replace(Content, Find, ReplaceWith)
	End Sub

	' Method to send email
	Public Sub Send
		Call ew_SendEmail(Sender, Recipient, Cc, Bcc, Subject, Content, Format)
	End Sub
End Class

'
'  ASPMaker 6 Email class (end)
' ------------------------------------
'
' -------------------------------------------------------
'  ASPMaker 6 Pager classes and functions (begin)
'
' Function to create numeric pager
Function ew_NewNumericPager(FromIndex, PageSize, RecordCount, Range)
	Set ew_NewNumericPager = New cNumericPager
	ew_NewNumericPager.FromIndex = CLng(FromIndex)
	ew_NewNumericPager.PageSize = CLng(PageSize)
	ew_NewNumericPager.RecordCount = CLng(RecordCount)
	ew_NewNumericPager.Range = CLng(Range)
	ew_NewNumericPager.Init
End Function

' Function to create next prev pager
Function ew_NewPrevNextPager(FromIndex, PageSize, RecordCount)
	Set ew_NewPrevNextPager = New cPrevNextPager
	ew_NewPrevNextPager.FromIndex = CLng(FromIndex)
	ew_NewPrevNextPager.PageSize = CLng(PageSize)
	ew_NewPrevNextPager.RecordCount = CLng(RecordCount)
	ew_NewPrevNextPager.Init
End Function

' Class for Pager item
Class cPagerItem
	Dim Start, Text, Enabled
End Class

' Class for Numeric pager
Class cNumericPager
	Dim Items()
	Dim Count, FromIndex, ToIndex, RecordCount, PageSize, Range
	Dim FirstButton, PrevButton, NextButton, LastButton, ButtonCount

	' Class Initialize
	Private Sub Class_Initialize()
		Set FirstButton = New cPagerItem
		Set PrevButton = New cPagerItem
		Set NextButton = New cPagerItem
		Set LastButton = New cPagerItem
	End Sub

	' Method to init pager
	Public Sub Init()
		If FromIndex > RecordCount Then FromIndex = RecordCount
		ToIndex = FromIndex + PageSize - 1
		If ToIndex > RecordCount Then ToIndex = RecordCount
		Count = -1
		ReDim Items(0)
		SetupNumericPager()
		Redim Preserve Items(Count)

		' Update button count
		ButtonCount = Count + 1
		If FirstButton.Enabled Then ButtonCount = ButtonCount + 1
		If PrevButton.Enabled Then ButtonCount = ButtonCount + 1
		If NextButton.Enabled Then ButtonCount = ButtonCount + 1
		If LastButton.Enabled Then ButtonCount = ButtonCount + 1
	End Sub

	' Add pager item
	Private Sub AddPagerItem(StartIndex, Text, Enabled)
		Count = Count + 1
		If Count > UBound(Items) Then
			Redim Preserve Items(UBound(Items)+10)
		End If
		Dim Item
		Set Item = New cPagerItem
		Item.Start = StartIndex
		Item.Text = Text
		Item.Enabled = Enabled
		Set Items(Count) = Item
	End Sub

	' Setup pager items
	Private Sub SetupNumericPager()
		Dim Eof, x, y, dx1, dx2, dy1, dy2, ny, HasPrev, TempIndex
		If RecordCount > PageSize Then
			Eof = (RecordCount < (FromIndex + PageSize))
			HasPrev = (FromIndex > 1)

			' First Button
			TempIndex = 1
			FirstButton.Start = TempIndex
			FirstButton.Enabled = (FromIndex > TempIndex)

			' Prev Button
			TempIndex = FromIndex - PageSize
			If TempIndex < 1 Then TempIndex = 1
			PrevButton.Start = TempIndex
			PrevButton.Enabled = HasPrev

			' Page links
			If HasPrev Or Not Eof Then
				x = 1
				y = 1
				dx1 = ((FromIndex-1)\(PageSize*Range))*PageSize*Range + 1
				dy1 = ((FromIndex-1)\(PageSize*Range))*Range + 1
				If (dx1+PageSize*Range-1) > RecordCount Then
					dx2 = (RecordCount\PageSize)*PageSize + 1
					dy2 = (RecordCount\PageSize) + 1
				Else
					dx2 = dx1 + PageSize*Range - 1
					dy2 = dy1 + Range - 1
				End If
				While x <= RecordCount
					If x >= dx1 And x <= dx2 Then
						Call AddPagerItem(x, y, FromIndex<>x)
						x = x + PageSize
						y = y + 1
					ElseIf x >= (dx1-PageSize*Range) And x <= (dx2+PageSize*Range) Then
						If x+Range*PageSize < RecordCount Then
							Call AddPagerItem(x, y & "-" & (y+Range-1), True)
						Else
							ny = (RecordCount-1)\PageSize + 1
							If ny = y Then
								Call AddPagerItem(x, y, True)
							Else
								Call AddPagerItem(x, y & "-" & ny, True)
							End If
						End If
						x = x + Range*PageSize
						y = y + Range
					Else
						x = x + Range*PageSize
						y = y + Range
					End If
				Wend
			End If

			' Next Button
			NextButton.Start = FromIndex + PageSize
			TempIndex = FromIndex + PageSize
			NextButton.Start = TempIndex
			NextButton.Enabled = Not Eof

			' Last Button
			TempIndex = ((RecordCount-1)\PageSize)*PageSize + 1
			LastButton.Start = TempIndex
			LastButton.Enabled = (FromIndex < TempIndex)
		End If
	End Sub

    ' Terminate
	Private Sub Class_Terminate()
		Set FirstButton = Nothing
		Set PrevButton = Nothing
		Set NextButton = Nothing
		Set LastButton = Nothing
		For Each Item In Items
			Set Item = Nothing
		Next
		Erase Items
	End Sub
End Class

' Class for PrevNext pager
Class cPrevNextPager
	Dim FirstButton, PrevButton, NextButton, LastButton
	Dim CurrentPage, PageSize, PageCount, FromIndex, ToIndex, RecordCount

	' Class Initialize
	Private Sub Class_Initialize()
		Set FirstButton = New cPagerItem
		Set PrevButton = New cPagerItem
		Set NextButton = New cPagerItem
		Set LastButton = New cPagerItem
	End Sub

	' Method to init pager
	Public Sub Init()
		Dim TempIndex
		CurrentPage = (FromIndex-1)\PageSize + 1
		PageCount = (RecordCount-1)\PageSize + 1
		If FromIndex > RecordCount Then FromIndex = RecordCount
		ToIndex = FromIndex + PageSize - 1
		If ToIndex > RecordCount Then ToIndex = RecordCount

		' First Button
		TempIndex = 1
		FirstButton.Start = TempIndex
		FirstButton.Enabled = (TempIndex <> FromIndex)

		' Prev Button
		TempIndex = FromIndex - PageSize
		If TempIndex < 1 Then TempIndex = 1
		PrevButton.Start = TempIndex
		PrevButton.Enabled = (TempIndex <> FromIndex)

		' Next Button
		TempIndex = FromIndex + PageSize
		If TempIndex > RecordCount Then TempIndex = FromIndex
		NextButton.Start = TempIndex
		NextButton.Enabled = (TempIndex <> FromIndex)

		' Last Button
		TempIndex = ((RecordCount-1)\PageSize)*PageSize + 1
		LastButton.Start = TempIndex
		LastButton.Enabled = (TempIndex <> FromIndex)
	End Sub

	' Terminate
	Private Sub Class_Terminate()
	Set FirstButton = Nothing
		Set PrevButton = Nothing
		Set NextButton = Nothing
		Set LastButton = Nothing
	End Sub
End Class

'
'  ASPMaker 6 Pager classes and functions (end)
' ------------------------------------------------------
'
' -----------------------------
'  ASPMaker 6 Field class
'
Class cField
	Dim TblVar ' Table var
	Dim FldName ' Field name
	Dim FldVar ' Field var
	Dim FldExpression ' Field expression (used in sql)
	Dim FldType ' Field type

	Public Property Get FldDataType() ' Field data type
		Select Case FldType
			Case 20, 3, 2, 16, 4, 5, 131, 6, 17, 18, 19, 21 ' Numeric
				FldDataType = 1
			Case 7, 133, 135 ' Date
				FldDataType = 2
			Case 134 ' Time
				FldDataType = 7
			Case 201, 203, 129, 130, 200, 202 ' String
				FldDataType = 3
			Case 11 ' Boolean
				FldDataType = 4
			Case 72 ' GUID
				FldDataType = 5
			Case Else
				FldDataType = 6
			End Select
	End Property
	Dim FldDateTimeFormat ' Date time format
	Dim CssStyle ' Css style
	Dim CssClass ' Css class
	Dim ImageAlt ' Image alt
	Dim ImageWidth ' Image width
	Dim ImageHeight ' Image height
	Dim ViewCustomAttributes ' View custom attributes

	' View Attributes
	Public Property Get ViewAttributes()
		Dim sAtt
		sAtt = ""
		If Trim(CssStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(CssStyle) & """" 
		End If
		If Trim(CssClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(CssClass) & """" 
		End If
		If Trim(ImageAlt) <> "" Then
			sAtt = sAtt & " alt=""" & Trim(ImageAlt) & """"
		End If
		If CInt(ImageWidth) > 0 Then
			sAtt = sAtt & " width=""" & CInt(ImageWidth) & """"
		End If
		If CInt(ImageHeight) > 0 Then
			sAtt = sAtt & " height=""" & CInt(ImageHeight) & """"
		End If
		If Trim(ViewCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(ViewCustomAttributes) 
		End If
		ViewAttributes = sAtt
	End Property
	Dim EditCustomAttributes ' Edit custom attributes

	' Edit Attributes
	Public Property Get EditAttributes()
		Dim sAtt
		sAtt = ""
		If Trim(CssStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(CssStyle) & """" 
		End If
		If Trim(CssClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(CssClass) & """" 
		End If
		If Trim(EditCustomAttributes) <> "" Then
			sAtt = sAtt & " " & Trim(EditCustomAttributes) 
		End If
		EditAttributes = sAtt
	End Property
	Dim CellCssClass ' Cell Css class
	Dim CellCssStyle ' Cell Css style

	' Cell Attributes
	Public Property Get CellAttributes()
		Dim sAtt
		sAtt = ""
		If Trim(CellCssStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(CellCssStyle) & """" 
		End If
		If Trim(CellCssClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(CellCssClass) & """" 
		End If
		CellAttributes = sAtt
	End Property

	' Sort Attributes
	Public Property Get Sort()
		Sort = Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar)
	End Property

	Public Property Let Sort(v)
		If Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar) <> v Then
			Session(EW_PROJECT_NAME & "_" & TblVar & "_" & EW_TABLE_SORT & "_" & FldVar) = v
		End If
	End Property

	Public Function ReverseSort()
		If Sort = "ASC" Then
			ReverseSort = "DESC"
		Else
			ReverseSort = "ASC"
		End If
	End Function
	Dim MultiUpdate ' Multi update
	Dim CurrentValue ' Current value
	Dim ViewValue ' View value
	Dim EditValue ' Edit value
	Dim EditValue2 ' Edit value 2 (search)
	Dim HrefValue ' Href value

	' Form value
	Private m_FormValue

	Public Property Get FormValue()
		FormValue = m_FormValue
	End Property

	Public Property Let FormValue(v)
		m_FormValue = v
		CurrentValue = m_FormValue
	End Property

	' QueryString value
	Private m_QueryStringValue

	Public Property Get QueryStringValue()
		QueryStringValue = m_QueryStringValue
	End Property

	Public Property Let QueryStringValue(v)
		m_QueryStringValue = v
		CurrentValue = m_QueryStringValue
	End Property

	' Database Value
	Dim m_DbValue

	Public Property Get DbValue()
		DbValue = m_DbValue
	End Property

	Public Property Let DbValue(v)
		m_DbValue = v
		CurrentValue = m_DbValue
	End Property

	' Set up database value
	Public Sub SetDbValue(value, default)
		Select Case FldType
			Case 2, 3, 16, 17, 18, 19, 20, 21 ' Int
				If IsNumeric(value) Then
					m_DbValue = CLng(value)
				Else
					m_DbValue = default
				End If
			Case 5, 6, 14, 131 ' Double
				If IsNumeric(value) Then
					m_DbValue = CDbl(value)
				Else
					m_DbValue = default
				End If
			Case 4 ' Single
				If IsNumeric(value) Then
					m_DbValue = CSng(value)
				Else
					m_DbValue = default
				End If
			Case 7, 133, 134, 135 ' Date
				If IsDate(value) Then
					m_DbValue = CDate(value)
				Else
					m_DbValue = default
				End If
			Case 201, 203, 129, 130, 200, 202 ' String
				m_DbValue = Trim(value)
				If m_DbValue = "" Then m_DbValue = default
			Case 128, 204, 205 ' Binary
				If IsNull(value) Then
					m_DbValue = default
				Else
					m_DbValue = value
				End If
			Case 72 ' GUID
				Dim RE
				Set RE = New RegExp
				RE.Pattern = "^(\{{1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{1})$"
				If RE.Test(Trim(value)) Then
					m_DbValue = Trim(value)
				Else
					m_DbValue = default
				End If
				Set RE = Nothing
			Case Else
				m_DbValue = value
		End Select
	End Sub

	' Session Value
	Public Property Get SessionValue()
		SessionValue = Session(EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_SessionValue")
	End Property

	Public Property Let SessionValue(v)
		Session(EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_SessionValue") = v
	End Property
	Dim Count ' Count
	Dim Total ' Total

	' AdvancedSearch Object
	Private m_AdvancedSearch

	Public Property Get AdvancedSearch()
		If Not IsObject(m_AdvancedSearch) Then Set m_AdvancedSearch = New cAdvancedSearch
		Set AdvancedSearch = m_AdvancedSearch
	End Property

	' Upload Object
	Private m_Upload

	Public Property Get Upload()
		If Not IsObject(m_Upload) Then
			Set m_Upload = New cUpload
			m_Upload.TblVar = TblVar
			m_Upload.FldVar = FldVar
		End If
		Set Upload = m_Upload
	End Property

	' Show object as string
	Public Function AsString()
		Dim AdvancedSearchAsString, UploadAsString
		If IsObject(m_AdvancedSearch) Then
			AdvancedSearchAsString = m_AdvancedSearch.AsString
		Else
			AdvancedSearchAsString = "{Null}"
		End If
		If IsObject(m_Upload) Then
			UploadAsString = m_Upload.AsString
		Else
			UploadAsString = "{Null}"
		End If
		AsString = "{" & _
			"FldName: " & FldName & ", " & _
			"FldVar: " & FldVar & ", " & _
			"FldExpression: " & FldExpression & ", " & _
			"FldType: " & FldType & ", " & _
			"FldDateTimeFormat: " & FldDateTimeFormat & ", " & _
			"CssStyle: " & CssStyle & ", " & _
			"CssClass: " & CssClass & ", " & _
			"ImageAlt: " & ImageAlt & ", " & _
			"ImageWidth: " & ImageWidth & ", " & _
			"ImageHeight: " & ImageHeight & ", " & _
			"ViewCustomAttributes: " & ViewCustomAttributes & ", " & _
			"EditCustomAttributes: " & EditCustomAttributes & ", " & _
			"CellCssStyle: " & CellCssStyle & ", " & _
			"CellCssClass: " & CellCssClass & ", " & _
			"Sort: " & Sort & ", " & _
			"MultiUpdate: " & MultiUpdate & ", " & _
			"CurrentValue: " & CurrentValue & ", " & _
			"ViewValue: " & ViewValue & ", " & _
			"EditValue: " & ValueToString(EditValue) & ", " & _
			"EditValue2: " & ValueToString(EditValue2) & ", " & _
			"HrefValue: " & HrefValue & ", " & _
			"FormValue: " & m_FormValue & ", " & _
			"QueryStringValue: " & m_QueryStringValue & ", " & _
			"DbValue: " & m_DbValue & ", " & _
			"SessionValue: " & SessionValue & ", " & _
			"Count: " & Count & ", " & _
			"Total: " & Total & ", " & _
			"AdvancedSearch: " & AdvancedSearchAsString & ", " & _
			"Upload: " & UploadAsString & _
			"}"
	End Function

	' Value to string
	Private Function ValueToString(value)
		If IsArray(value) Then
			ValueToString = "[Array]"
		Else
			ValueToString = value
		End If
	End Function

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_AdvancedSearch) Then
			Set m_AdvancedSearch = Nothing
		End If
		If IsObject(m_Upload) Then
			Set m_Upload = Nothing
		End If
	End Sub
End Class

'
'  ASPMaker 6 Field class (end)
' -----------------------------------

%>
<%

' --------------------------------------------------
'  ASPMaker 6 Advanced Search class (begin)
'
Class cAdvancedSearch
	Dim SearchValue ' Search value
	Dim SearchOperator ' Search operator
	Dim SearchCondition ' Search condition
	Dim SearchValue2 ' Search value 2
	Dim SearchOperator2 ' Search operator 2

	' Show object as string
	Public Function AsString()
		AsString = "{" & _
			"SearchValue: " & SearchValue & ", " & _
			"SearchOperator: " & SearchOperator & ", " & _
			"SearchCondition: " & SearchCondition & ", " & _
			"SearchValue2: " & SearchValue2 & ", " & _
			"SearchOperator2: " & SearchOperator2 & _
			"}"
	End Function
End Class

'
'  ASPMaker 6 Advanced Search class (end)
' -------------------------------------------------

%>
<%

' ---------------------------------------
'  ASPMaker 6 Upload class (begin)
'
Class cUpload
	Dim Index ' Index to handle multiple form elements

	' Class initialize
	Private Sub Class_Initialize
		Index = 0
	End Sub
	Dim TblVar ' Table variable
	Dim FldVar ' Field variable

	' Error message
	Private m_Message

	Public Property Get Message()
		Message = m_Message
	End Property
	Dim DbValue ' Value from database

	' Upload value
	Dim m_Value

	Public Property Get Value()
		Value = m_Value
	End Property

	Public Property Let Value(v)
		m_Value = v
	End Property

	' Upload action
	Private m_Action

	Public Property Get Action()
		Action = m_Action
	End Property
	Dim UploadPath ' Upload path

	' Upload file name
	Private m_FileName

	Public Property Get FileName()
		FileName = m_FileName
	End Property

	' Upload file size
	Private m_FileSize

	Public Property Get FileSize()
		FileSize = m_FileSize
	End Property

	' File content type
	Private m_ContentType

	Public Property Get ContentType()
		ContentType = m_ContentType
	End Property

	' Image width
	Private m_ImageWidth

	Public Property Get ImageWidth()
		ImageWidth = m_ImageWidth
	End Property

	' Image height
	Private m_ImageHeight

	Public Property Get ImageHeight()
		ImageHeight = m_ImageHeight
	End Property

	' Save Db value to Session
	Public Sub SaveDbToSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session(sSessionID & "_DbValue") = DbValue
	End Sub

	' Restore Db value from Session
	Public Sub RestoreDbFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		DbValue = Session(sSessionID & "_DbValue")
	End Sub

	' Remove Db value from Session
	Public Sub RemoveDbFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session.Contents.Remove(sSessionID & "_DbValue")
	End Sub

	' Save Upload values to Session
	Public Sub SaveToSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session(sSessionID & "_Action") = m_Action
		Session(sSessionID & "_FileSize") = m_FileSize
		Session(sSessionID & "_FileName") = m_FileName
		Session(sSessionID & "_ContentType") = m_ContentType
		Session(sSessionID & "_ImageWidth") = m_ImageWidth
		Session(sSessionID & "_ImageHeight") = m_ImageHeight
		Session(sSessionID & "_Value") = m_Value
	End Sub

	' Restore Upload values from Session
	Public Sub RestoreFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		m_Action = Session(sSessionID & "_Action")
		m_FileSize = Session(sSessionID & "_FileSize")
		m_FileName = Session(sSessionID & "_FileName")
		m_ContentType = Session(sSessionID & "_ContentType")
		m_ImageWidth = Session(sSessionID & "_ImageWidth")
		m_ImageHeight = Session(sSessionID & "_ImageHeight")
		m_Value = Session(sSessionID & "_Value")
	End Sub

	' Remove Upload values from Session
	Public Sub RemoveFromSession()
		Dim sSessionID
		sSessionID = EW_PROJECT_NAME & "_" & TblVar & "_" & FldVar & "_" & Index
		Session.Contents.Remove(sSessionID & "_Action")
		Session.Contents.Remove(sSessionID & "_FileSize")
		Session.Contents.Remove(sSessionID & "_FileName")
		Session.Contents.Remove(sSessionID & "_ContentType")
		Session.Contents.Remove(sSessionID & "_ImageWidth")
		Session.Contents.Remove(sSessionID & "_ImageHeight")
		Session.Contents.Remove(sSessionID & "_Value")
	End Sub

	' Function to check the file type of the uploaded file
	Private Function UploadAllowedFileExt(FileName)
		If Trim(FileName & "") = "" Then
			UploadAllowedFileExt = True
			Exit Function
		End If
		Dim Ext, Pos, arExt, FileExt
		arExt = Split(EW_UPLOAD_ALLOWED_FILE_EXT & "", ",")
		Ext = ""
		Pos = InStrRev(FileName, ".")
		If Pos > 0 Then	Ext = Mid(FileName, Pos+1)
		UploadAllowedFileExt = False
		For Each FileExt in arExt
	 		If LCase(Trim(FileExt)) = LCase(Ext) Then
				UploadAllowedFileExt = True
				Exit For
			End If
		Next
	End Function

	' Get upload file
	Public Function UploadFile()
		Dim sFldVar, sFldVarAction, sFldVarWidth, sFldVarHeight
		sFldVar = FldVar
		sFldVarAction = "a" & Mid(sFldVar, 2)
		sFldVarWidth = "wd" & Mid(sFldVar, 2)
		sFldVarHeight = "ht" & Mid(sFldVar, 2)

		' Get action
		m_Action = objForm.GetValue(sFldVarAction)

		' Get and check the upload file size
		m_FileSize = objForm.GetUploadFileSize(sFldVar)
		If m_FileSize > 0 And CLng(EW_MAX_FILE_SIZE) > 0 Then
			If m_FileSize > CLng(EW_MAX_FILE_SIZE) Then
				m_Message = Replace("Max. file size (%s bytes) exceeded.", "%s", EW_MAX_FILE_SIZE)
				UploadFile = False
				Exit Function
			End If
		End If

		' Get and check the upload file type
		m_FileName = objForm.GetUploadFileName(sFldVar)
		m_FileName = Replace(m_FileName, " ", "_") ' Replace space with underscore
		If Not UploadAllowedFileExt(m_FileName) Then
			m_Message = "File type is not allowed."
			UploadFile = False
			Exit Function
		End If

		' Get upload file content type
		m_ContentType = objForm.GetUploadFileContentType(sFldVar)

		' Get upload value
		m_Value = objForm.GetUploadFileData(sFldVar)

		' Get image width and height
		m_ImageWidth = objForm.GetUploadImageWidth(sFldVar)
		m_ImageHeight = objForm.GetUploadImageHeight(sFldVar)
		If m_ImageWidth < 0 Or m_ImageHeight < 0 Then
			m_ImageWidth = objForm.GetValue(sFldVarWidth)
			m_ImageHeight = objForm.GetValue(sFldVarHeight)
		End If
		UploadFile = True ' Normal return
	End Function

	' Resize image
	Public Function Resize(width, height, interpolation)
		Dim wrkwidth, wrkheight
		If Not IsNull(m_Value) Then
			wrkwidth = width
			wrkheight = height
			If ew_ResizeBinary(m_Value, wrkwidth, wrkheight, interpolation) Then
				m_ImageWidth = wrkwidth
				m_ImageHeight = wrkheight
				m_FileSize = LenB(m_Value)
			End If
		End If
	End Function

	' Show object as string
	Public Function AsString()
		AsString = "{" & _
			"Index: " & Index & ", " & _
			"Message: " & m_Message & ", " & _
			"Action: " & m_Action & ", " & _
			"UploadPath: " & UploadPath & ", " & _
			"FileName: " & m_FileName & ", " & _
			"FileSize: " & m_FileSize & ", " & _
			"ContentType: " & m_ContentType & ", " & _
			"ImageWidth: " & m_ImageWidth & ", " & _
			"ImageHeight: " & m_ImageHeight & _
			"}"
	End Function
End Class

'
'  ASPMaker 6 Upload class (end)
' -------------------------------------

%>
<%

' ----------------------------------------------------
' ASPMaker 6 Advanced Security class (begin)
'
Class cAdvancedSecurity
	Dim m_ArUserLevel
	Dim m_ArUserLevelPriv
	Dim arUserID

	' Current user name
	Public Property Get CurrentUserName()
		CurrentUserName = Session(EW_SESSION_USER_NAME) & ""
	End Property

	Public Property Let CurrentUserName(v)
		Session(EW_SESSION_USER_NAME) = v
	End Property

	' Current user id
	Public Property Get CurrentUserID()
		CurrentUserID = Session(EW_SESSION_USER_ID) & ""
	End Property

	Public Property Let CurrentUserID(v)
		Session(EW_SESSION_USER_ID) = v
	End Property

	' Current parent user id
	Public Property Get CurrentParentUserID()
		CurrentParentUserID = Session(EW_SESSION_PARENT_USER_ID) & ""
	End Property

	Public Property Let CurrentParentUserID(v)
		Session(EW_SESSION_PARENT_USER_ID) = v
	End Property

	' Current user level id
	Public Property Get CurrentUserLevelID()
		CurrentUserLevelID = Session(EW_SESSION_USER_LEVEL_ID)
	End Property

	Public Property Let CurrentUserLevelID(v)
		Session(EW_SESSION_USER_LEVEL_ID) = v
	End Property

	' Current user level value
	Public Property Get CurrentUserLevel()
		CurrentUserLevel = Session(EW_SESSION_USER_LEVEL)
	End Property

	Public Property Let CurrentUserLevel(v)
		Session(EW_SESSION_USER_LEVEL) = v
	End Property

	' Can add
	Public Property Get CanAdd()
		CanAdd = ((CurrentUserLevel And EW_ALLOW_ADD) = EW_ALLOW_ADD)
	End Property

	' Can delete
	Public Property Get CanDelete()
		CanDelete = ((CurrentUserLevel And EW_ALLOW_DELETE) = EW_ALLOW_DELETE)
	End Property

	' Can edit
	Public Property Get CanEdit()
		CanEdit = ((CurrentUserLevel And EW_ALLOW_EDIT) = EW_ALLOW_EDIT)
	End Property

	' Can view
	Public Property Get CanView()
		CanView = ((CurrentUserLevel And EW_ALLOW_VIEW) = EW_ALLOW_VIEW)
	End Property

	' Can list
	Public Property Get CanList()
		CanList = ((CurrentUserLevel And EW_ALLOW_LIST) = EW_ALLOW_LIST)
	End Property

	' Can report
	Public Property Get CanReport()
		CanReport = ((CurrentUserLevel And EW_ALLOW_REPORT) = EW_ALLOW_REPORT)
	End Property

	' Can search
	Public Property Get CanSearch()
		CanSearch = ((CurrentUserLevel And EW_ALLOW_SEARCH) = EW_ALLOW_SEARCH)
	End Property

	' Can admin
	Public Property Get CanAdmin()
		CanAdmin = ((CurrentUserLevel And EW_ALLOW_ADMIN) = EW_ALLOW_ADMIN)
	End Property

	' Last url
	Public Property Get LastUrl()
		LastUrl = Request.Cookies(EW_PROJECT_NAME)("lasturl")
	End Property

	' Save last url
	Public Sub SaveLastUrl()
		Dim s, q
		s = Request.ServerVariables("SCRIPT_NAME")
		q = Request.ServerVariables("QUERY_STRING")
		If q <> "" Then s = s & "?" & q
		Response.Cookies(EW_PROJECT_NAME)("lasturl") = s
	End Sub

	' Auto login
	Public Function AutoLogin()
		Dim usr, pwd, sFilter
		If Request.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" Then
			usr = Request.Cookies(EW_PROJECT_NAME)("username")
			pwd = Request.Cookies(EW_PROJECT_NAME)("password")
			pwd = TEAdecrypt(ew_Decode(pwd), EW_RANDOM_KEY)
			AutoLogin = ValidateUser(usr, pwd)
		Else
			AutoLogin = False
		End If
	End Function

	' Validate user
	Public Function ValidateUser(usr, pwd)
		Dim rs, sFilter, sSql
		Dim CaseSensitive
		CaseSensitive = False ' Modify case sensitivity here
		ValidateUser = False

		' Check other users
		If Not ValidateUser Then
				sFilter = "([UserName] = '" & ew_AdjustSql(usr) & "')"

				' Set up filter (Sql Where Clause) and get Return Sql
				' Sql constructor in <UseTable> class, <UserTable>info.asp

				Customers.CurrentFilter = sFilter
				sSql = Customers.SQL
				Set rs = conn.Execute(sSql)
				If Not rs.Eof Then
					If CaseSensitive Then
						ValidateUser = (rs("passwrd") = pwd)
					Else
						ValidateUser = (LCase(rs("passwrd")) = LCase(pwd))
					End If
					If ValidateUser Then
						Session(EW_SESSION_STATUS) = "login"
						Session(EW_SESSION_SYS_ADMIN) = 0 ' Non System Administrator
						CurrentUserName = rs("UserName") ' Load user name
						CurrentUserID = rs("CustomerID") ' Load user id
						CurrentParentUserID = rs("CustomerID") ' Load parent user id
					End If
				End If
				rs.Close
				Set rs = Nothing
		End If
	End Function

	' No user level security
	Public Sub SetUpUserLevel()
	End Sub

	' Load current user level
	Public Sub LoadCurrentUserLevel(Table)
		Call LoadUserLevel()
		CurrentUserLevel = CurrentUserLevelPriv(Table)
	End Sub

	' Get current user privilege
	Private Function CurrentUserLevelPriv(TableName)
		If IsLoggedIn() Then
			CurrentUserLevelPriv = GetUserLevelPrivEx(TableName, CurrentUserLevelID)
		Else

			'CurrentUserLevelPriv = GetUserLevelPrivEx(TableName, 0)
			CurrentUserLevelPriv = 0
		End If
	End Function

	' Get user privilege based on table name and user level
	Public Function GetUserLevelPrivEx(TableName, UserLevelID)
		GetUserLevelPrivEx = 0
		If CStr(UserLevelID) = "-1" Then ' System Administrator
			If EW_USER_LEVEL_COMPAT Then
				GetUserLevelPrivEx = 31 ' Use old user level values
			Else
				GetUserLevelPrivEx = 127 ' Use new user level values (separate View/Search)
			End If
		ElseIf UserLevelID >= 0 Then
			If IsArray(m_ArUserLevelPriv) Then
				Dim i
				For i = 0 to UBound(m_ArUserLevelPriv, 2)
					If CStr(m_ArUserLevelPriv(0, i)) = CStr(TableName) And _
						CStr(m_ArUserLevelPriv(1, i)) = CStr(UserLevelID) Then
						GetUserLevelPrivEx = m_ArUserLevelPriv(2, i)
						If IsNull(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
						If Not IsNumeric(GetUserLevelPrivEx) Then GetUserLevelPrivEx = 0
						GetUserLevelPrivEx = CLng(GetUserLevelPrivEx)
						Exit For
					End If
				Next
			End If
		End If
	End Function

	' Get current user level name
	Public Function CurrentUserLevelName()
		CurrentUserLevelName = GetUserLevelName(CurrentUserLevelID)
	End Function

	' Get user level name based on user level
	Public Function GetUserLevelName(UserLevelID)
		GetUserLevelName = ""
		If CStr(UserLevelID) = "-1" Then
			GetUserLevelName = "Administrator"
		ElseIf UserLevelID >= 0 Then
			If IsArray(m_ArUserLevel) Then
				Dim i
				For i = 0 to UBound(m_ArUserLevel, 2)
					If CStr(m_ArUserLevel(0, i)) = CStr(UserLevelID) Then
						GetUserLevelName = m_ArUserLevel(1, i)
						Exit For
					End If
				Next
			End If
		End If
	End Function

	' Sub to display all the User Level settings (for debug only)
	Public Sub ShowUserLevelInfo()
		Dim i
		If IsArray(m_ArUserLevel) Then
			Response.Write "User Levels:<br>"
			Response.Write "UserLevelId, UserLevelName<br>"
			For i = 0 To UBound(m_ArUserLevel, 2)
				Response.Write "&nbsp;&nbsp;" & m_ArUserLevel(0, i) & ", " & _
					m_ArUserLevel(1, i) & "<br>"
			Next
		Else
			Response.Write "No User Level definitions." & "<br>"
		End If
		If IsArray(m_ArUserLevelPriv) Then
			Response.Write "User Level Privs:<br>"
			Response.Write "TableName, UserLevelId, UserLevelPriv<br>"
			For i = 0 To UBound(m_ArUserLevelPriv, 2)
				Response.Write "&nbsp;&nbsp;" & m_ArUserLevelPriv(0, i) & ", " & _
					m_ArUserLevelPriv(1, i) & ", " & m_ArUserLevelPriv(2, i) & "<br>"
			Next
		Else
			Response.Write "No User Level privilege settings." & "<br>"
		End If
		Response.Write "CurrentUserLevel = " & CurrentUserLevel & "<br>"
	End Sub

	' Function to check privilege for List page (for menu items)
	Public Function AllowList(TableName)
		AllowList = CBool(CurrentUserLevelPriv(TableName) And EW_ALLOW_LIST)
	End Function

	' Check if user is logged in
	Public Function IsLoggedIn()
		IsLoggedIn = (Session(EW_SESSION_STATUS) = "login")
	End Function

	' Check if user is system administrator
	Public Function IsSysAdmin()
		IsSysAdmin = (Session(EW_SESSION_SYS_ADMIN) = 1)
	End Function

	' Check if user is administrator
	Function IsAdmin()
		IsAdmin = (CurrentUserLevelID = -1 Or IsSysAdmin)
	End Function

	' Save user level to session
	Public Sub SaveUserLevel()
		Session(EW_SESSION_AR_USER_LEVEL) = m_ArUserLevel
		Session(EW_SESSION_AR_USER_LEVEL_PRIV) = m_ArUserLevelPriv
	End Sub

	' Load user level from session
	Public Sub LoadUserLevel()
		If Not IsArray(Session(EW_SESSION_AR_USER_LEVEL)) Then
			Call SetupUserLevel()
			Call SaveUserLevel()
		Else
			m_ArUserLevel = Session(EW_SESSION_AR_USER_LEVEL)
			m_ArUserLevelPriv = Session(EW_SESSION_AR_USER_LEVEL_PRIV)
		End If
	End Sub

	' Function to get user email
	Public Function CurrentUserEmail()
		CurrentUserEmail = CurrentUserInfo("inv_EmailAddress")
	End Function

	' Function to get user info
	Public Function CurrentUserInfo(fieldname)
		On Error Resume Next
		CurrentUserInfo = Null
		If CurrentUserName = "" Then Exit Function
		Dim rs, sSql, fldtype

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in <UseTable> class, <UserTable>info.asp

		sFilter = "([UserName] = '" & ew_AdjustSql(CurrentUserName) & "')"
		Customers.CurrentFilter = sFilter
		sSql = Customers.SQL
		Set rs = conn.Execute(sSql)
		If Not rs.Eof Then
			CurrentUserInfo = rs(fieldname)
			fldtype = rs.Fields(fieldname).Type
			If fldtype = 18 Or fldtype = 19 Then
				CurrentUserInfo = ew_Conv(CurrentUserInfo, fldtype)
			End If
		End If
		rs.Close
		Set rs = Nothing
	End Function

	' list of allowed user ids for this user
	Function IsValidUserID(userid)
		IsValidUserID = False
		If IsLoggedIn() Then
			Dim sFilter, sSql, rs

			' Load user id list
			If Not IsArray(arUserID) Then
				Redim arUserID(0)
				sFilter = Customers.AddUserIDFilter("", CurrentUserID)
				Customers.CurrentFilter = sFilter
				sSql = Customers.SQL
				Set rs = conn.execute(sSql)
				Do While Not rs.Eof
					Redim Preserve arUserID(UBound(arUserID)+1)
					arUserID(UBound(arUserID)) = rs("CustomerID")
					rs.MoveNext
				Loop
				rs.Close
				Set rs = Nothing
			End If

			' Check user id
			For i = 1 to UBound(arUserID)
				If arUserID(i) & "" = userid & "" Then
					IsValidUserID = True
					Exit Function
				End If
			Next
		End If
	End Function
End Class

'
' ASPMaker 6 Advanced Security class (end)
' --------------------------------------------------

%>
<%

' ----------------------------------------------
' ASPMaker 6 common functions (begin)
'
' Check if valid operator
Function ew_IsValidOpr(Opr, FldType)
	ew_IsValidOpr = (Opr = "=" Or Opr = "<" Or Opr = "<=" Or _
		Opr = ">" Or Opr = ">=" Or Opr = "<>")
	If FldType = EW_DATATYPE_STRING Then
		ew_IsValidOpr = ew_IsValidOpr Or Opr = "LIKE" Or Opr = "NOT LIKE" Or Opr = "STARTS WITH"
	End If
End Function

' Quoted value for field type
Function ew_QuotedValue(Value, FldType) 
	Select Case FldType
	Case EW_DATATYPE_STRING
		ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
	Case EW_DATATYPE_GUID
		If EW_IS_MSACCESS Then
			ew_QuotedValue = "{guid " & ew_AdjustSql(Value) & "}"
		Else
			ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
		End If
	Case EW_DATATYPE_DATE
		If EW_IS_MSACCESS Then
			ew_QuotedValue = "#" & ew_AdjustSql(Value) & "#"
		Else
			ew_QuotedValue = "'" & ew_AdjustSql(Value) & "'"
		End If
	Case Else
		ew_QuotedValue = Value
	End Select
End Function

' Pad zeros before number
Function ew_ZeroPad(m, t)
	ew_ZeroPad = String(t - Len(m), "0") & m
End Function

' IIf function
Function ew_IIf(cond, v1, v2)
	On Error Resume Next
	If CBool(cond) Then
		ew_IIf = v1
	Else
		ew_IIf = v2
	End If
End Function

' Convert different data type value
Function ew_Conv(v, t)
	Select Case t

	' adBigInt/adUnsignedBigInt
	Case 20, 21
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CLng(v)
		End If

	' adSmallInt/adInteger/adTinyInt/adUnsignedTinyInt/adUnsignedSmallInt/adUnsignedInt/adBinary
	Case 2, 3, 16, 17, 18, 19, 128
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CLng(v)
		End If

	' adSingle
	Case 4
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CSng(v)
		End If

	' adDouble/adCurrency/adNumeric
	Case 5, 6, 131
		If IsNull(v) Then
			ew_Conv = Null
		Else
			ew_Conv = CDbl(v)
		End If
	Case Else
		ew_Conv = v
	End Select
End Function

' Function for debug
Sub ew_Trace(aMsg)
	On Error Resume Next
	Dim fso, ts
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(Server.MapPath("debug.txt"), 8, True)
	ts.writeline(aMsg)
	ts.Close
	Set ts = Nothing
	Set fso = Nothing
End Sub

' Function to compare values with special handling for null values
Function ew_CompareValue(v1, v2)
	If IsNull(v1) And IsNull(v2) Then
		ew_CompareValue = True
	ElseIf IsNull(v1) Or IsNull(v2) Then
		ew_CompareValue = False
	Else
		ew_CompareValue = (v1 = v2)
	End If
End Function

' Adjust sql for special characters
Function ew_AdjustSql(str)
	Dim sWrk
	sWrk = Trim(str & "")
	sWrk = Replace(sWrk, "'", "''") ' Adjust for Single Quote
	sWrk = Replace(sWrk, "[", "[[]") ' Adjust for Open Square Bracket
	ew_AdjustSql = sWrk
End Function

' Build sql based on different sql part
Function ew_BuildSql(sSelect, sWhere, sGroupBy, sHaving, sOrderBy, sFilter, sSort)
	Dim sSql, sDbWhere, sDbOrderBy
	sDbWhere = sWhere
	If sDbWhere <> "" Then
		sDbWhere = "(" & sDbWhere & ")"
	End If
	If sFilter <> "" Then
		If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
		sDbWhere = sDbWhere & "(" & sFilter & ")"
	End If	
	sDbOrderBy = sOrderBy
	If sSort <> "" Then
		sDbOrderBy = sSort
	End If
	sSql = sSelect
	If sDbWhere <> "" Then
		sSql = sSql & " WHERE " & sDbWhere
	End If
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If
	If sDbOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sDbOrderBy
	End If
	ew_BuildSql = sSql
End Function

' Note: Object "conn" is required
' Return sql scalar value
Function ew_ExecuteScalar(SQL)
	ew_ExecuteScalar = Null
	If Trim(SQL&"") = "" Then	Exit Function
	Dim rs
	Set rs = conn.Execute(SQL)
	If Not rs.Eof Then ew_ExecuteScalar = rs(0)
	rs.Close
	Set rs = Nothing
End Function

' Clone recordset
Function ew_CloneRs(Rs)
	Dim oStream
	Dim oRsClone

	' Save the recordset to the stream object
	Set oStream = Server.CreateObject("ADODB.Stream")
	Rs.Save oStream

	' Open the stream object into a new recordset
	Set oRsClone = Server.CreateObject("ADODB.Recordset")
	oRsClone.Open oStream, , , 2

	' Return the cloned recordset
	Set ew_CloneRs = oRsClone

	' Release the reference
	Set oRsClone = Nothing
End Function

' Function to Load a Text File
Function ew_LoadTxt(fn)
	Dim fso, fobj

	' Get text file content
	ew_LoadTxt = ""
	If Trim(fn) <> "" Then
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(Server.MapPath(fn)) Then
			Set fobj = fso.OpenTextFile(Server.MapPath(fn))
			ew_LoadTxt = fobj.ReadAll ' Read all Content
			fobj.Close
			Set fobj = Nothing
		End If
		Set fso = Nothing
	End If
End Function

' Write Audit Trail (login/logout)
Sub ew_WriteAuditTrailOnLogInOut(logtype)
	On Error Resume Next
	Dim table, sKey
	table = logtype
	sKey = ""

	' Write Audit Trail
	Dim filePfx, curDate, curTime, id, user, action, field, keyvalue, oldvalue, newvalue
	Dim i
	filePfx = "log"
	curDate = ew_ZeroPad(Year(Date), 4) & "/" & ew_ZeroPad(Month(Date), 2) & "/" & ew_ZeroPad(Day(Date), 2)
	curTime = ew_ZeroPad(Hour(Time), 2) & ":" & ew_ZeroPad(Minute(Time), 2) & ":" & ew_ZeroPad(Second(Time), 2)
	id = Request.ServerVariables("SCRIPT_NAME")
	user = Security.CurrentUserName
	action = logtype
	Call ew_WriteAuditTrail(filePfx, curDate, curTime, id, user, action, table, field, keyvalue, oldvalue, newvalue)
End Sub

' Write Audit Trail (insert/update/delete)
Sub ew_WriteAuditTrail(pfx, curDate, curTime, id, user, action, table, field, keyvalue, oldvalue, newvalue)
	On Error Resume Next
	Dim fso, ts, sMsg, sFn, sFolder
	Dim bWriteHeader, sHeader
	Dim userwrk
	userwrk = user
	If userwrk = "" Then userwrk = "-1" ' assume Administrator if no user

	' Write audit trail to log file
	sHeader = "date" & vbTab & _
		"time" & vbTab & _
		"id" & vbTab & _
		"user" & vbTab & _
		"action" & vbTab & _
		"table" & vbTab & _
		"field" & vbTab & _
		"key value" & vbTab & _
		"old value" & vbTab & _
		"new value"
	sMsg = curDate & vbTab & _
		curTime & vbTab & _
		id & vbTab & _
		userwrk & vbTab & _
		action & vbTab & _
		table & vbTab & _
		field & vbTab & _
		keyvalue & vbTab & _
		oldvalue & vbTab & _
		newvalue
	sFolder = EW_AUDIT_TRAIL_PATH
	sFn = pfx & "_" & ew_ZeroPad(Year(Date), 4) & ew_ZeroPad(Month(Date), 2) & ew_ZeroPad(Day(Date), 2) & ".txt"
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	bWriteHeader = Not fso.FileExists(ew_UploadPathEx(True, sFolder) & sFn)
	Set ts = fso.OpenTextFile(ew_UploadPathEx(True, sFolder) & sFn, 8, True)
	If bWriteHeader Then
		ts.writeline(sHeader)
	End If
	ts.writeline(sMsg)
	ts.Close
	Set ts = Nothing
	Set fso = Nothing

	' Sample code to write audit trail to database
	' Dim sAuditSql
	' sAuditSql = "INSERT INTO AuditTrail VALUES (" & _
	' 	"#" & ew_AdjustSql(curDate) & "#, " & _
	'	"#" & ew_AdjustSql(curTime) & "#, " & _
	'	"""" & ew_AdjustSql(id) & """, " & _
	' 	"""" & ew_AdjustSql(user) & """, " & _
	'	"""" & ew_AdjustSql(action) & """, " & _
	'	"""" & ew_AdjustSql(table) & """, " & _
	'	"""" & ew_AdjustSql(field) & """, " & _
	'	"""" & ew_AdjustSql(keyvalue) & """, " & _
	'	"""" & ew_AdjustSql(oldvalue) & """, " & _
	'	"""" & ew_AdjustSql(newvalue) & """)"
	'	' Response.Write sAuditSql ' uncomment to debug
	' conn.execute(sAuditSql)

End Sub

'-------------------------------------------------------------------------------
' Functions for default date format
' ANamedFormat = 0-8, where 0-4 same as VBScript
' 5 = "yyyymmdd"
' 6 = "mmddyyyy"
' 7 = "ddmmyyyy"
' 8 = Short Date + Short Time
' 9 = "yyyymmdd HH:MM:SS"
' 10 = "mmddyyyy HH:MM:SS"
' 11 = "ddmmyyyy HH:MM:SS"
' 12 = "HH:MM:SS"
' Format date time based on format type
Function ew_FormatDateTime(ADate, ANamedFormat)
	If IsDate(ADate) Then
		If ANamedFormat >= 0 And ANamedFormat <= 4 Then
			ew_FormatDateTime = FormatDateTime(ADate, ANamedFormat)
		ElseIf ANamedFormat = 5 Or ANamedFormat = 9 Then
			ew_FormatDateTime = Year(ADate) & EW_DATE_SEPARATOR & Month(ADate) & EW_DATE_SEPARATOR & Day(ADate)
		ElseIf ANamedFormat = 6 Or ANamedFormat = 10 Then
			ew_FormatDateTime = Month(ADate) & EW_DATE_SEPARATOR & Day(ADate) & EW_DATE_SEPARATOR & Year(ADate)
		ElseIf ANamedFormat = 7 Or ANamedFormat = 11 Then
			ew_FormatDateTime = Day(ADate) & EW_DATE_SEPARATOR & Month(ADate) & EW_DATE_SEPARATOR & Year(ADate)
		ElseIf ANamedFormat = 8 Then
			ew_FormatDateTime = FormatDateTime(ADate, 2)
			If Hour(ADate) <> 0 Or Minute(ADate) <> 0 Or Second(ADate) <> 0 Then
				ew_FormatDateTime = ew_FormatDateTime & " " & FormatDateTime(ADate, 4) & ":" & ew_ZeroPad(Second(ADate), 2)
			End If
		ElseIf ANamedFormat = 12 Then
			ew_FormatDateTime = ew_ZeroPad(Hour(ADate), 2) & ":" & ew_ZeroPad(Minute(ADate), 2) & ":" & ew_ZeroPad(Second(ADate), 2)
		Else
			ew_FormatDateTime = ADate
		End If
		If ANamedFormat >= 9 And ANamedFormat <= 11 Then
				ew_FormatDateTime = ew_FormatDateTime & " " & ew_ZeroPad(Hour(ADate), 2) & ":" & ew_ZeroPad(Minute(ADate), 2) & ":" & ew_ZeroPad(Second(ADate), 2)
		End If
	Else
		ew_FormatDateTime = ADate
	End If
End Function

' Unformat date time based on format type
Function ew_UnFormatDateTime(ADate, ANamedFormat)
	Dim arDateTime, arDate
	ADate = Trim(ADate & "")
	While Instr(ADate, "  ") > 0
		ADate = Replace(ADate, "  ", " ")
	Wend
	arDateTime = Split(ADate, " ")
	If UBound(arDateTime) < 0 Then
		ew_UnFormatDateTime = ADate
		Exit Function
	End If
	If ANamedFormat = 0 And IsDate(ADate) Then
		ew_UnFormatDateTime = Year(arDateTime(0)) & "/" & Month(arDateTime(0)) & "/" & Day(arDateTime(0))
		If UBound(arDateTime) > 0 Then
			ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(1)
		End If
	Else
		arDate = Split(arDateTime(0), EW_DATE_SEPARATOR)
		If UBound(arDate) = 2 Then
			ew_UnFormatDateTime = arDateTime(0)
			If ANamedFormat = 6 Or ANamedFormat = 10 Then ' mmddyyyy
				If Len(arDate(0)) <= 2 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 4 Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(0) & "/" & arDate(1)
				End If
			ElseIf (ANamedFormat = 7 Or ANamedFormat = 11) Then ' ddmmyyyy
				If Len(arDate(0)) <= 2 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 4 Then
					ew_UnFormatDateTime = arDate(2) & "/" & arDate(1) & "/" & arDate(0)
				End If
			ElseIf ANamedFormat = 5 Or ANamedFormat = 9 Then ' yyyymmdd
				If Len(arDate(0)) <= 4 And Len(arDate(1)) <= 2 And Len(arDate(2)) <= 2 Then
					ew_UnFormatDateTime = arDate(0) & "/" & arDate(1) & "/" & arDate(2)
				End If
			End If
			If UBound(arDateTime) > 0 Then
				If IsDate(arDateTime(1)) Then ' Is time
					ew_UnFormatDateTime = ew_UnFormatDateTime & " " & arDateTime(1)
				End If
			End If
		Else
			ew_UnFormatDateTime = ADate
		End If
	End If
End Function

' Format currency
Function ew_FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next 
	ew_FormatCurrency = FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatCurrency = Expression
	End If
End Function

' Format number
Function ew_FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next 
	ew_FormatNumber = FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatNumber = Expression
	End If
End Function

' Format percent
Function ew_FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next
	ew_FormatPercent = FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		Err.Clear
		ew_FormatPercent = FormatNumber(Expression*100, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) & "%"
	End If
End Function

' Encode html
Function ew_HtmlEncode(Expression)
	ew_HtmlEncode = Server.HtmlEncode(Expression & "")
End Function

' Generate Value Separator based on current row count
' rowcnt - zero based row count
Function ew_ValueSeparator(rowcnt)
	ew_ValueSeparator = ", "
End Function

' Generate View Option Separator based on current row count (Multi-Select / CheckBox)
' rowcnt - zero based row count
Function ew_ViewOptionSeparator(rowcnt)
	ew_ViewOptionSeparator = ", "

	' Sample code to adjust 2 options per row
	'If ((rowcnt + 1) Mod 2 = 0) Then ' 2 options per row
		'ew_ViewOptionSeparator = ew_ViewOptionSeparator & "<br>"
	'End If

End Function

' Render repeat column table
' rowcnt - zero based row count
Function ew_RepeatColumnTable(totcnt, rowcnt, repeatcnt, rendertype)
	Dim sWrk, i
	sWrk = ""

	' Render control start
	If rendertype = 1 Then
		If rowcnt = 0 Then sWrk = sWrk & "<table class=""aspmakerlist"">"
		If (rowcnt mod repeatcnt = 0) Then sWrk = sWrk & "<tr>"
		sWrk = sWrk & "<td>"

	' Render control end
	ElseIf rendertype = 2 Then
		sWrk = sWrk & "</td>"
		If (rowcnt mod repeatcnt = repeatcnt -1) Then
			sWrk = sWrk & "</tr>"
		ElseIf rowcnt = totcnt Then
			For i = ((rowcnt mod repeatcnt) + 1) to repeatcnt - 1
				sWrk = sWrk & "<td>&nbsp;</td>"
			Next
			sWrk = sWrk & "</tr>"
		End If
		If rowcnt = totcnt Then sWrk = sWrk & "</table>"
	End If
	ew_RepeatColumnTable = sWrk
End Function

' Truncate Memo Field based on specified length, string truncated to nearest space or CrLf
Function ew_TruncateMemo(str, ln)
	Dim i, j, k
	If Len(str) > 0 And Len(str) > ln Then
		k = 1
		Do While k > 0 And k < Len(str)
			i = InStr(k, str, " ", 1)
			j = InStr(k, str, vbCrLf, 1)
			If i < 0 And j < 0 Then ' Not able to truncate
				ew_TruncateMemo = str
				Exit Function
			Else

				' Get nearest space or CrLf
				If i > 0 And j > 0 Then
					If i < j Then
						k = i
					Else
						k = j
					End If
				ElseIf i > 0 Then
					k = i
				ElseIf j > 0 Then
					k = j
				End If

				' Get truncated text
				If k >= ln Then
					ew_TruncateMemo = Mid(str, 1, k-1) & "..."
					Exit Function
				Else
					k = k + 1
				End If
			End If
		Loop
	Else
		ew_TruncateMemo = str
	End If
End Function

' Send notify email
Sub ew_SendNotifyEmail(sFn, sSubject, sTable, sKey, sAction)
	On Error Resume Next

	' Send Email
	If EW_SENDER_EMAIL <> "" And EW_RECIPIENT_EMAIL <> "" Then
		Dim Email
		Set Email = New cEmail
		Email.Load(sFn)
		Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
		Email.ReplaceRecipient(EW_RECIPIENT_EMAIL) ' Replace Recipient
		Email.ReplaceSubject(sSubject) ' Replace Subject
		Email.ReplaceContent "<!--table-->", sTable
		Email.ReplaceContent "<!--key-->", sKey
		Email.ReplaceContent "<!--action-->", sAction
		Email.Send()
		Set Email = Nothing
	End If
End Sub

' Function to Send out Email
' Supports CDO, w3JMail and ASPEmail
Function ew_SendEmail(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat)
	On Error Resume Next
	Dim i, objMail, sServerVersion, sIISVer, EmailComponent, arrEmail, sEmail
	Dim arCDO, arASPEmail, arw3JMail, arEmailComponent
	sServerVersion = Request.ServerVariables("SERVER_SOFTWARE")
	If InStr(sServerVersion, "Microsoft-IIS") > 0 Then
		i = InStr(sServerVersion, "/")
		If i > 0 Then
			sIISVer = Trim(Mid(sServerVersion, i+1))
		End If
	End If
	arw3JMail = Array("w3JMail", "JMail.Message")
	arASPEmail = Array("ASPEmail", "Persits.MailSender")
	If sIISVer < "5.0" Then ' NT using CDONTS
		arCDO = Array("CDO", "CDONTS.NewMail")
	Else ' 2000 / XP / 2003 using CDO
		arCDO = Array("CDO", "CDO.Message")
	End If

	' Change your precedence here
	arEmailComponent = Array(arw3JMail, arASPEmail, arCDO)
	For i = 0 to UBound(arEmailComponent)
		Err.Clear
		Set objMail = Server.CreateObject(arEmailComponent(i)(1))
		If Err.Number = 0 Then
			EmailComponent = arEmailComponent(i)(0)
			Exit For
		End If
	Next
	If Err.Number <> 0 Then
		ew_SendEmail = False
		Exit Function
	End If
	If EmailComponent = "w3JMail" Then

		' Set objMail = Server.CreateObject("JMail.Message")
		objMail.Logging = True
		objMail.Silent = True
		objMail.From = sFrEmail
		arrEmail = Split(Replace(sToEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipient sEmail
			End If
		Next
		arrEmail = Split(Replace(sCcEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipientCC sEmail
			End If
		Next
		arrEmail = Split(Replace(sBccEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddRecipientBCC sEmail
			End If
		Next
		objMail.Subject = sSubject
		If LCase(sFormat) = "html" Then
			objMail.HTMLBody = sMail
		Else
			objMail.Body = sMail
		End If
		If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
			objMail.MailServerUserName = EW_SMTP_SERVER_USERNAME
			objMail.MailServerPassword = EW_SMTP_SERVER_PASSWORD
		End If
		ew_SendEmail = objMail.Send(EW_SMTP_SERVER)
		If Not ew_SendEmail Then
			Err.Raise vbObjectError + 1, EmailComponent, objMail.Log
		End If
		Set objMail = nothing
	ElseIf EmailComponent = "ASPEmail" Then

		' Set objMail = Server.CreateObject("Persits.MailSender")
		objMail.From = sFrEmail
		arrEmail = Split(Replace(sToEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddAddress sEmail
			End If
		Next
		arrEmail = split(Replace(sCcEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddCC sEmail
			End If
		Next
		arrEmail = split(Replace(sBccEmail, ",", ";"), ";")
		For i = 0 to UBound(arrEmail)
			sEmail = Trim(arrEmail(i))
			If sEmail <> "" Then
				objMail.AddBcc sEmail
			End If
		Next
		If LCase(sFormat) = "html" Then
			objMail.IsHTML = True ' html
		Else
			objMail.IsHTML = False ' text
		End If
		objMail.Subject = sSubject
		objMail.Body = sMail
		objMail.Host = EW_SMTP_SERVER
		If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
			objMail.Username = EW_SMTP_SERVER_USERNAME
			objMail.Password = EW_SMTP_SERVER_PASSWORD
		End If
		ew_SendEmail = objMail.Send
		Set objMail = Nothing
	ElseIf EmailComponent = "CDO" Then
		Dim objConfig, sSmtpServer, iSmtpServerPort
		If sIISVer < "5.0" Then ' NT using CDONTS

			' Set objMail = Server.CreateObject("CDONTS.NewMail")
			objMail.From = sFrEmail
			objMail.To = Replace(sToEmail, ",", ";")
			If sCcEmail <> "" Then
				objMail.Cc = Replace(sCcEmail, ",", ";")
			End If
			If sBccEmail <> "" Then
				objMail.Bcc = Replace(sBccEmail, ",", ";")
			End If
			If LCase(sFormat) = "html" Then
				objMail.BodyFormat = 0 ' 0 means HTML format, 1 means text
				objMail.MailFormat = 0 ' 0 means MIME, 1 means text
			End If
			objMail.Subject = sSubject
			objMail.Body = sMail
			objMail.Send
			Set objMail = Nothing
		Else ' 2000 / XP / 2003 using CDO

			' Set up Configuration
			Set objConfig = Server.CreateObject("CDO.Configuration")
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EW_SMTP_SERVER ' cdoSMTPServer
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = EW_SMTP_SERVER_PORT ' cdoSMTPServerPort
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			If EW_SMTP_SERVER_USERNAME <> "" And EW_SMTP_SERVER_PASSWORD <> "" Then
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic (clear text)
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = EW_SMTP_SERVER_USERNAME
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EW_SMTP_SERVER_PASSWORD
			End If
			objConfig.Fields.Update

			' Set up Mail
			Set objMail = Server.CreateObject("CDO.Message")
			objMail.From = sFrEmail
			objMail.To = Replace(sToEmail, ",", ";")
			If sCcEmail <> "" Then
				objMail.Cc = Replace(sCcEmail, ",", ";")
			End If
			If sBccEmail <> "" Then
				objMail.Bcc = Replace(sBccEmail, ",", ";")
			End If
			If LCase(sFormat) = "html" Then
				objMail.HtmlBody = sMail
			Else
				objMail.TextBody = sMail
			End If
			objMail.Subject = sSubject
			If EW_SMTP_SERVER <> "" And LCase(EW_SMTP_SERVER) <> "localhost" Then
				Set objMail.Configuration = objConfig ' Use Configuration
				objMail.Send
			Else
				On Error Resume Next
				objMail.Send ' Send without Configuration
				If Err.Number <> 0 Then
					If Hex(Err.Number) = "80040220" Then ' Requires Configuration
						Set objMail.Configuration = objConfig
						Err.Clear
						On Error GoTo 0
						objMail.Send
					Else
						Dim ErrNo, ErrSrc, ErrDesc
						ErrNo = Err.Number
						ErrSrc = Err.Source
						ErrDesc = Err.Description
						On Error GoTo 0
						Err.Raise ErrNo, ErrSrc, ErrDesc
					End If
				End If
			End If
			Set objMail = Nothing
			Set objConfig = Nothing
		End If
		ew_SendEmail = (Err.Number = 0)
	End If
End Function 

' Return path of the uploaded file
'	Parameter: If PhyPath is true(1), return physical path on the server;
'	           If PhyPath is false(0), return relative URL
Function ew_UploadPathEx(PhyPath, DestPath)
	Dim Pos
	If PhyPath Then
		ew_UploadPathEx = Request.ServerVariables("APPL_PHYSICAL_PATH")
		ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
		ew_UploadPathEx = ew_UploadPathEx & Replace(DestPath, "/", "\")
	Else
		ew_UploadPathEx = Request.ServerVariables("APPL_MD_PATH")
		Pos = InStr(1, ew_UploadPathEx, "Root", 1)
		If Pos > 0 Then	ew_UploadPathEx = Mid(ew_UploadPathEx, Pos+4)
		ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
		ew_UploadPathEx = ew_UploadPathEx & DestPath
	End If
	ew_UploadPathEx = ew_IncludeTrailingDelimiter(ew_UploadPathEx, PhyPath)
End Function

' Change the file name of the uploaded file
Function ew_UploadFileNameEx(Folder, FileName)
	Dim OutFileName

	' By default, ewUniqueFileName() is used to get an unique file name.
	' Amend your logic here

	OutFileName = ew_UniqueFileName(Folder, FileName)

	' Return computed output file name
	ew_UploadFileNameEx = OutFileName
End Function

' Return path of the uploaded file
' returns global upload folder, for backward compatibility only
Function ew_UploadPath(PhyPath)
	ew_UploadPath = ew_UploadPathEx(PhyPath, EW_UPLOAD_DEST_PATH)
End Function

' Change the file name of the uploaded file
' use global upload folder, for backward compatibility only
Function ew_UploadFileName(FileName)
	ew_UploadFileName = ew_UploadFileNameEx(ew_UploadPath(True), FileName)
End Function

' Generate an unique file name (filename(n).ext)
Function ew_UniqueFileName(Folder, FileName)
	If FileName = "" Then FileName = ew_DefaultFileName()
	If FileName = "." Then
		Response.Write "Invalid file name: " & FileName
		Response.End
		Exit Function
	End If
	If Folder = "" Then
		Response.Write "Unspecified folder"
		Response.End
		Exit Function
	End If
	Dim Name, Ext, Pos
	Name = ""
	Ext = ""
	Pos = InStrRev(FileName, ".")
	If Pos = 0 Then
		Name = FileName
		Ext = ""
	Else
		Name = Mid(FileName, 1, Pos-1)
		Ext = Mid(FileName, Pos+1)
	End If
	Folder = ew_IncludeTrailingDelimiter(Folder, True)
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Folder) Then
		If Not ew_CreateFolder(Folder) Then
			Response.Write "Folder does not exist: " & Folder
			Set fso = Nothing
			Exit Function
		End If
	End If
	Dim Suffix, Index
	Index = 0
	Suffix = ""

	' Check to see if filename exists
	While fso.FileExists(folder & Name & Suffix & "." & Ext)
		Index = Index + 1
		Suffix = "(" & Index & ")"
	Wend
	Set fso = Nothing

	' Return unique file name
	ew_UniqueFileName = Name & Suffix & "." & Ext
End Function

' Create a default file name (yyyymmddhhmmss.bin)
Function ew_DefaultFileName
	Dim dt
	dt = Now()
	ew_DefaultFileName = ew_ZeroPad(Year(dt), 4) & ew_ZeroPad(Month(dt), 2) &  _
		ew_ZeroPad(Day(dt), 2) & ew_ZeroPad(Hour(dt), 2) & _
		ew_ZeroPad(Minute(dt), 2) & ew_ZeroPad(Second(dt), 2) & ".bin"
End Function

' Include the last delimiter for a path
Function ew_IncludeTrailingDelimiter(Path, PhyPath)
	If PhyPath Then
		If Right(Path, 1) <> "\" Then Path = Path & "\"
	Else
		If Right(Path, 1) <> "/" Then Path = Path & "/"
	End If
	ew_IncludeTrailingDelimiter = Path
End Function

' Write the paths for config/debug only
Sub ew_WriteUploadPaths
	Response.Write "Request.ServerVariables(""APPL_PHYSICAL_PATH"")=" & _
		Request.ServerVariables("APPL_PHYSICAL_PATH") & "<br>"
	Response.Write "Request.ServerVariables(""APPL_MD_PATH"")=" & _
		Request.ServerVariables("APPL_MD_PATH") & "<br>"
End Sub 

' Get full url
Function ew_FullUrl()
	Dim sUrl
	sUrl = "http"
	If Request.ServerVariables("HTTPS") <> "off" Then sUrl = sUrl & "s"
	sUrl = sUrl & "://"
	sUrl = sUrl & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	ew_FullUrl = sUrl
End Function

' Convert to full url
Function ew_ConvertFullUrl(url)
	Dim sUrl
	If url = "" Then
		ew_ConvertFullUrl = ""
	ElseIf Instr(url, "://") > 0 Then
		ew_ConvertFullUrl = url
	Else
		sUrl = ew_FullUrl
		ew_ConvertFullUrl = Mid(sUrl, 1, InStrRev(sUrl, "/")) & url
	End If
End Function

' Check if folder exists
Function ew_FolderExists(Folder)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	ew_FolderExists = fso.FolderExists(Folder)
	Set fso = Nothing
End Function

' Check if file exists
Function ew_FileExists(Folder, File)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	ew_FileExists = fso.FileExists(Folder & File)
	Set fso = Nothing
End Function

' Delete file
Sub ew_DeleteFile(FilePath)
	On Error Resume Next
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If FilePath <> "" And fso.FileExists(FilePath) Then
		fso.DeleteFile(FilePath)
	End If
	Set fso = Nothing
End Sub

' Rename file
Sub ew_RenameFile(OldFilePath, NewFilePath)
	On Error Resume Next
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If OldFilePath <> "" And fso.FileExists(OldFilePath) Then
		fso.MoveFile OldFilePath, NewFilePath
	End If
	Set fso = Nothing
End Sub

' Create folder
Function ew_CreateFolder(Folder)
	On Error Resume Next
	ew_CreateFolder = False
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Folder) Then
		If ew_CreateFolder(fso.GetParentFolderName(Folder)) Then
			fso.CreateFolder(Folder)
			If Err.Number = 0 Then ew_CreateFolder = True
		End If
	Else
		ew_CreateFolder = True
	End If
	Set fso = Nothing
End Function

' Add an element to a position of an array
Function ew_AddItemToArray(ar, pos, aritem)
	Dim newar(), d1, d2, d3, p
	Dim i, j
	If not IsArray(aritem) Then
		ew_AddItemToArray = ar
		Exit Function
	End If
	d3 = UBound(aritem)
	If not IsArray(ar) Then
		Redim newar(d3,0)
		For i = 0 to d3
			newar(i,0) = aritem(i)
		Next
		ew_AddItemToArray = newar
		Exit Function
	Else
		d1 = UBound(ar,1)
		d2 = UBound(ar,2)
		p = pos
		If p < 0 Then p = 0 ' add at front
		If p > d2 Then p = d2 ' add at end
		Redim newar(d1, d2+1)

		' Copy item before p
		For j = 0 to p-1
			For i = 0 to d1
				newar(i,j) = ar(i,j)
			Next
		Next

		' Copy new item
		For i = 0 to d1
			If i <= d3 Then
				newar(i,p) = aritem(i)
			Else
				newar(i,p) = "" ' Initialize to empty string
			End If
		Next

		' Copy the rest
		For j = p to d2
			For i = 0 to d1
				newar(i,j+1) = ar(i,j)
			Next
		Next
	End If
	ew_AddItemToArray = newar
End Function

' Remove an element from a position of an array
Function ew_RemoveItemFromArray(ar, pos)
	Dim newar(), d1, d2
	Dim i, j
	ew_RemoveItemFromArray = Null
	If IsArray(ar) Then
		d1 = UBound(ar,1)
		d2 = UBound(ar,2)
		If pos < 0 Or pos > d2 Then
			ew_RemoveItemFromArray = ar
			Exit Function
		End If
		If d2 = 0 Then
			ew_RemoveItemFromArray = Null
		Else
			Redim newar(d1, d2-1)

			' Copy items before pos
			For j = 0 to pos-1
				For i = 0 to d1
					newar(i,j) = ar(i,j)
				Next
			Next

			' Copy items after pos
			For j = pos+1 to d2
				For i = 0 to d1
					newar(i,j-1) = ar(i,j)
				Next
			Next
			ew_RemoveItemFromArray = newar
		End If
	End If
End Function

'
' ASPMaker 6 common functions (end)
' -------------------------------------------

%>
<%

' ----------------------------------------------------------------
'  ASPMaker 6 Default Request Form Object Class (begin)
'
Class cFormObj
	Dim Index ' Index to handle multiple form elements

	' Class Initialize
	Private Sub Class_Initialize
		Index = 0
	End Sub

	' Get form element name based on index
	Function GetIndexedName(name)
		If Index <= 0 Then
			GetIndexedName = name
		Else
			Dim Pos
			Pos = InStr(name, "_")
			If Pos = 2 Or Pos = 3 Then
				GetIndexedName = Mid(name, 1, Pos-1) & Index & Mid(name, Pos)
			Else
				GetIndexedName = name
			End If
		End If
	End Function

	' Get value for form element
	Function GetValue(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If Request.Form(wrkname).Count > 0 Then
			GetValue = Request.Form(wrkname)
		Else
			GetValue = Null
		End If
	End Function
End Class

'
'  ASPMaker 6 Default Request Form Object Class (end)
' --------------------------------------------------------------

%>
<%

' ---------------------------------------------
'  ASPMaker 6 get upload object (begin)
'
Function ew_GetUploadObj()
		Set ew_GetUploadObj = New cUploadObj
End Function

'
'  ASPMaker 6 get upload object (end)
' -------------------------------------------

%>
<%

' ---------------------------------------------------------
'  ASPMaker 6 Default Upload Object Class (begin)
'
Class cUploadObj
	Dim rawData, separator, lenSeparator, dict
	Dim currentPos, inStrByte, tempValue, mValue, value
	Dim intDict, begPos, endPos
	Dim nameN, isValid, nameValue, midValue
	Dim rawStream
	Dim Index

	' Class Inialize
	Private Sub Class_Initialize
		Index = 0
		If Request.TotalBytes > 0 Then
			Set rawStream = Server.CreateObject("ADODB.Stream")
			rawStream.Type = 1 'adTypeBinary
			rawStream.Mode = 3 'adModeReadWrite
			rawStream.Open
			rawStream.Write Request.BinaryRead(Request.TotalBytes)
			rawStream.Position = 0
			rawData = rawStream.Read
			separator = MidB(rawData, 1, InStrB(1, rawData, ChrB(13)) - 1)
			lenSeparator = LenB(separator)
			Set dict = Server.CreateObject("Scripting.Dictionary")
			currentPos = 1
			inStrByte = 1
			tempValue = ""
			While inStrByte > 0
				inStrByte = InStrB(currentPos, rawData, separator)
				mValue = inStrByte - currentPos
				If mValue > 1 Then
					value = MidB(rawData, currentPos, mValue)
					Set intDict = Server.CreateObject("Scripting.Dictionary")
					begPos = 1 + InStrB(1, value, ChrB(34))
					endPos = InStrB(begPos + 1, value, ChrB(34))
					nameN = MidB(value, begPos, endPos - begPos)
					isValid = True
					If InStrB(1, value, StringToByte("Content-Type")) > 1 Then
						begPos = 1 + InStrB(endPos + 1, value, ChrB(34))
						endPos = InStrB(begPos + 1, value, ChrB(34))
						If endPos > 0 Then
							intDict.Add "FileName", ConvertToText(rawStream, currentPos + begPos - 2, endPos - begPos, MidB(value, begPos, endPos - begPos))
							begPos = 14 + InStrB(endPos + 1, value, StringToByte("Content-Type:"))
							endPos = InStrB(begPos, value, ChrB(13))
							intDict.Add "ContentType", ConvertToText(rawStream, currentPos + begPos - 2, endPos - begPos, MidB(value, begPos, endPos - begPos))
							begPos = endPos + 4
							endPos = LenB(value)
							nameValue = MidB(value, begPos, ((endPos - begPos) - 1))
						Else
							endPos = begPos + 1
							isValid = False
						End If
					Else
						nameValue = ConvertToText(rawStream, currentPos + endPos + 3, mValue - endPos - 4, MidB(value, endPos + 5))
					End If
					If isValid = True Then
						If dict.Exists(ByteToString(nameN)) Then
							Set intDict = dict.Item(ByteToString(nameN))
							If Right(intDict.Item("Value"), 2) = vbCrLf Then
								intDict.Item("Value") = Left(intDict.Item("Value"), Len(intDict.Item("Value"))-2)
							End If
							intDict.Item("Value") = intDict.Item("Value") & ", " & nameValue
						Else
							intDict.Add "Value", nameValue
							intDict.Add "Name", nameN
							dict.Add ByteToString(nameN), intDict
						End If
					End If
				End If
				currentPos = lenSeparator + inStrByte
			Wend
			rawStream.Close
			Set rawStream = Nothing
		End If
	End Sub

	' Get form element name based on index
	Function GetIndexedName(name)
		If Index <= 0 Then
			GetIndexedName = name
		Else
			GetIndexedName = Mid(name, 1, 1) & Index & Mid(name, 2)
		End If
	End Function

	' Get value for form element
	Function GetValue(name)
		Dim wrkname
		Dim gv
		GetValue = Null ' default return Null
		If IsObject(dict) Then
			wrkname = GetIndexedName(name)
			If dict.Exists(wrkname) Then
				gv = CStr(dict(wrkname).Item("Value"))
				gv = Left(gv, Len(gv)-2)
				GetValue = gv
			End If
		End If
	End Function

	' Get upload file size
	Function GetUploadFileSize(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileSize = LenB(dict(wrkname).Item("Value"))
		Else
			GetUploadFileSize = 0
		End If
	End Function

	' Get upload file name
	Function GetUploadFileName(name)
		Dim wrkname, temp, tempPos
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			temp = dict(wrkname).Item("FileName")
			tempPos = 1 + InStrRev(temp, "\")
			GetUploadFileName = Mid(temp, tempPos)
		Else
			GetUploadFileName = ""
		End If
	End Function

	' Get file content type
	Function GetUploadFileContentType(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileContentType = dict(wrkname).Item("ContentType")
		Else
			GetUploadFileContentType = ""
		End If
	End Function

	' Get upload file data
	Function GetUploadFileData(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		If dict.Exists(wrkname) Then
			GetUploadFileData = dict(wrkname).Item("Value")
			If LenB(GetUploadFileData) Mod 2 = 1 Then
				GetUploadFileData = GetUploadFileData & ChrB(0)
			End If
		Else
			GetUploadFileData = Null
		End If
	End Function

	' Get file image width
	Function GetUploadImageWidth(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		Dim ImageHeight
		Call GetImageDimension(GetUploadFileData(wrkname), GetUploadImageWidth, ImageHeight)
	End Function

	' Get file image height
	Function GetUploadImageHeight(name)
		Dim wrkname
		wrkname = GetIndexedName(name)
		Dim ImageWidth
		Call GetImageDimension(GetUploadFileData(wrkname), ImageWidth, GetUploadImageHeight)
	End Function

	' Convert length
	Private Function ConvertLength(b)
		ConvertLength = CLng(AscB(LeftB(b, 1)) + (AscB(RightB(b, 1)) * 256))
	End Function

	' Convert length 2
	Private Function ConvertLength2(b)
		ConvertLength2 = CLng(AscB(RightB(b, 1)) + (AscB(LeftB(b, 1)) * 256))
	End Function

	' Get image dimension
	Sub GetImageDimension(img, wd, ht)
		Dim sPNGHeader, sGIFHeader, sBMPHeader, sJPGHeader, sHeader, sImgType
		sImgType = "(unknown)"

		' Image headers, do not changed
		sPNGHeader = ChrB(137) & ChrB(80) & ChrB(78)
		sGIFHeader = ChrB(71) & ChrB(73) & ChrB(70)
		sBMPHeader = ChrB(66) & ChrB(77)
		sJPGHeader = ChrB(255) & ChrB(216) & ChrB(255)
		sHeader = MidB(img, 1, 3)

		' Handle GIF
		If sHeader = sGIFHeader Then
			sImgType = "GIF"
			wd = ConvertLength(MidB(img, 7, 2))
			ht = ConvertLength(MidB(img, 9, 2))

		' Handle BMP
		ElseIf LeftB(sHeader, 2) = sBMPHeader Then
			sImgType = "BMP"
			wd = ConvertLength(MidB(img, 19, 2))
			ht = ConvertLength(MidB(img, 23, 2))

		' Handle PNG
		ElseIf sHeader = sPNGHeader Then
			sImgType = "PNG"
			wd = ConvertLength2(MidB(img, 19, 2))
			ht = ConvertLength2(MidB(img, 23, 2))

		' Handle JPG
		Else
			Dim size, markersize, pos, bEndLoop
			size = LenB(img)
			pos = InStrB(img, sJPGHeader)
			If pos <= 0 Then
				wd = -1
				ht = -1
				Exit Sub
			End If
			sImgType = "JPG"
			pos = pos + 2
			bEndLoop = False
			Do While Not bEndLoop and pos < size
				Do While AscB(MidB(img, pos, 1)) = 255 and pos < size
					pos = pos + 1
				Loop
				If AscB(MidB(img, pos, 1)) < 192 or AscB(MidB(img, pos, 1)) > 195 Then
					markersize = ConvertLength2(MidB(img, pos+1, 2))
					pos = pos + markersize + 1
				Else
					bEndLoop = True
				End If
			Loop
			If Not bEndLoop Then
				wd = -1
				ht = -1
			Else
				wd = ConvertLength2(MidB(img, pos+6, 2))
				ht = ConvertLength2(MidB(img, pos+4, 2))
			End If
		End If
	End Sub

	' Convert string to byte
	Function StringToByte(toConv)
		Dim i, tempChar
		For i = 1 to Len(toConv)
			tempChar = Mid(toConv, i, 1)
			StringToByte = StringToByte & ChrB(AscB(tempChar))
		Next
	End Function

	' Convert byte to string
	Private Function ByteToString(ToConv)
		Dim i
		On Error Resume Next
		For i = 1 to LenB(ToConv)
			ByteToString = ByteToString & Chr(AscB(MidB(ToConv,i,1)))
		Next
	End Function

	' Convert to text
	Function ConvertToText(objStream, iStart, iLength, binData)
		On Error Resume Next
		If EW_UPLOAD_CHARSET <> "" Then
			Dim tmpStream
			Set tmpStream = Server.CreateObject("ADODB.Stream")
			tmpStream.Type = 1 'adTypeBinary
			tmpStream.Mode = 3 'adModeReadWrite
			tmpStream.Open
			objStream.Position = iStart
			objStream.CopyTo tmpStream, iLength
			tmpStream.Position = 0
			tmpStream.Type = 2 'adTypeText
			tmpStream.Charset = EW_UPLOAD_CHARSET
			ConvertToText = tmpStream.ReadText
			tmpStream.Close
			Set tmpStream = Nothing
		Else
			ConvertToText = ByteToString(binData)
		End If
		ConvertToText = Trim(ConvertToText & "")
	End Function

	' Class terminate
	Private Sub Class_Terminate

		' Dispose dictionary
		If IsObject(intDict) Then
			intDict.RemoveAll
			Set intDict = Nothing
		End If
		If IsObject(dict) Then
			dict.RemoveAll
			Set dict = Nothing
		End If
	End Sub
End Class

'
'  ASPMaker 6 Default Upload Object Class (end)
' -------------------------------------------------------

%>
<%

' Save binary to file
Function ew_SaveFile(folder, fn, filedata)
	On Error Resume Next
	Dim oStream
	ew_SaveFile = False
	If Not ew_SaveFileByComponent(folder, fn, filedata) Then
		If ew_CreateFolder(folder) Then
			Set oStream = Server.CreateObject("ADODB.Stream")
			oStream.Type = 1 ' 1=adTypeBinary
			oStream.Open
			oStream.Write ew_ConvertToBinary(filedata)
			oStream.SaveToFile folder & fn, 2 ' 2=adSaveCreateOverwrite
			oStream.Close
			Set oStream = Nothing
			If Err.Number = 0 Then ew_SaveFile = True
		End If
	End If
End Function

' Convert raw to binary
Function ew_ConvertToBinary(rawdata)
	Dim oRs
	Set oRs = Server.CreateObject("ADODB.Recordset")

	' Create field in an empty RecordSet
	Call oRs.Fields.Append("Blob", 205, LenB(rawdata)) ' Add field with type adLongVarBinary
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Blob").AppendChunk(rawdata & ChrB(0))
	Call oRs.Update()

	' Save Blob Data
	ew_ConvertToBinary = oRs.Fields("Blob").GetChunk(LenB(rawdata))

	' Close RecordSet
	Call oRs.Close()
	Set oRs = Nothing
End Function
%>
<%

' Resize binary to thumbnail
Function ew_ResizeBinary(filedata, width, height, interpolation)
	ew_ResizeBinary = False ' No resize
End Function

' Resize file to thumbnail file
Function ew_ResizeFile(fn, tn, width, height, interpolation)
	On Error Resume Next
	Dim fso

	' Just copy across
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(fn) Then
		fso.CopyFile fn, tn, True
	End If
	Set fso = Nothing
	ew_ResizeFile = True
End Function

' Resize file to binary
Function ew_ResizeFileToBinary(fn, width, height, interpolation)
	On Error Resume Next
	Dim oStream, fso
	ew_ResizeFileToBinary = Null

	' Return file content in binary
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(fn) Then
		Set oStream = Server.CreateObject("ADODB.Stream")
		oStream.Type = 1 ' 1=adTypeBinary
		oStream.Open
		oStream.LoadFromFile fn
		ew_ResizeFileToBinary = oStream.Read
		oStream.Close
		Set oStream = Nothing
	End If
	Set fso = Nothing
End Function

' Save file by component
Function ew_SaveFileByComponent(folder, fn, filedata)
	ew_SaveFileByComponent = False
End Function
%>
<%

' Highlight value based on basic search / advanced search keywords
Function ew_Highlight(src, bkw, bkwtype, akw)
	Dim i, x, y, outstr, kwlist, kw, kwstr
	outstr = ""
	If Len(src) > 0 And (Len(bkw) > 0 Or Len(akw) > 0) Then
		kwstr = bkw
		If Len(akw) > 0 Then
			If Len(kwstr) > 0 Then kwstr = kwstr & " "
			kwstr = kwstr & akw
		End If
		kwlist = Split(kwstr, " ")
		x = 1
		Call ew_GetKeyword(src, kwlist, x, y, kw)
		Do While y > 0
			outstr = outstr & Mid(src, x, y-x) & _
				"<span name=""ewHighlightSearch"" id=""ewHighlightSearch"" class=""ewHighlightSearch"">" & _
				Mid(src, y, Len(kw)) & "</span>"
			x = y + Len(kw)
			Call ew_GetKeyword(src, kwlist, x, y, kw)
		Loop
		outstr = outstr & Mid(src, x)
	Else
		outstr = src
	End If
	ew_Highlight = outstr
End Function

' Get keyword
Sub ew_GetKeyword(src, kwlist, x, y, kw)
	Dim i, thisy, thiskw, wrky, wrkkw
	thisy = -1
	thiskw = ""
	For i = 0 to UBound(kwlist)
		wrkkw = Trim(kwlist(i))
		wrky = InStr(x, src, wrkkw, EW_HIGHLIGHT_COMPARE)
		If wrky > 0 Then
			If thisy = -1 Then
				thisy = wrky
				thiskw = wrkkw
			ElseIf wrky < thisy Then
				thisy = wrky
				thiskw = wrkkw
			End If
		End If
	Next
	y = thisy
	kw = thiskw
End Sub 
%>
<%

' Functions for backward compatibilty
' Get current user name
Function CurrentUserName()
	On Error Resume Next
	CurrentUserName = Security.CurrentUserName
End Function

' Get current user ID
Function CurrentUserID()
	On Error Resume Next
	CurrentUserID = Security.CurrentUserID
End Function

' Get current parent user ID
Function CurrentParentUserID()
	On Error Resume Next
	CurrentParentUserID = Security.CurrentParentUserID
End Function

' Get current user level
Function CurrentUserLevel()
	On Error Resume Next
	CurrentUserLevel = Security.CurrentUserLevelID
End Function 

' Allow list
Function AllowList(TableName)
	On Error Resume Next
	AllowList = Security.AllowList(TableName)
End Function

' Is Logged In
Function IsLoggedIn()
	On Error Resume Next
	IsLoggedIn = Security.IsLoggedIn
End Function

' Is System Admin
Function IsSysAdmin()
	On Error Resume Next
	IsSysAdmin = Security.IsSysAdmin
End Function
%>
<%

' Get server variable by name
Function ew_GetServerVariable(Name)
	ew_GetServerVariable = Request.ServerVariables(Name)
End Function

' Get user IP
Function ew_CurrentUserIP()
	ew_CurrentUserIP = ew_GetServerVariable("REMOTE_HOST")
End Function

' Get current host name, e.g. "www.mycompany.com"
Function ew_CurrentHost()
	ew_CurrentUserIP = ew_GetServerVariable("HTTP_HOST")
End Function

' Get current date in default date format
Function ew_CurrentDate()
	If EW_DEFAULT_DATE_FORMAT = 6 Or EW_DEFAULT_DATE_FORMAT = 7 Then
		ew_CurrentDate = ew_FormatDateTime(Date, EW_DEFAULT_DATE_FORMAT)
	Else
		ew_CurrentDate = ew_FormatDateTime(Date, 5)
    End If
End Function

' Get current time in hh:mm:ss format
Function ew_CurrentTime()
	Dim DT
	DT = Now()
	ew_CurrentTime = ew_ZeroPad(Hour(DT), 2) & ":" & _
		ew_ZeroPad(Minute(DT), 2) & ":" & ew_ZeroPad(Second(DT), 2)
End Function

' Get current date in default date format with
' Current time in hh:mm:ss format
Function ew_CurrentDateTime()
	Dim DT
	DT = Now()
	If EW_DEFAULT_DATE_FORMAT = 6 Or EW_DEFAULT_DATE_FORMAT = 7 Then
		ew_CurrentDateTime = ew_FormatDateTime(DT, EW_DEFAULT_DATE_FORMAT)
	Else
		ew_CurrentDateTime = ew_FormatDateTime(DT, 5)
	End If
	ew_CurrentDateTime = ew_CurrentDateTime & " " & _
		ew_ZeroPad(Hour(DT), 2) & ":" & ew_ZeroPad(Minute(DT), 2) & _
		":" & ew_ZeroPad(Second(DT), 2)
End Function

' Return master keys in a dictionary
Function ew_CurrentMasterKeys()

'###	Dim d, sName
'###	Set d = Server.CreateObject("Scripting.Dictionary")
'###	For Each sName in Session.Contents
'###		If Left(sName, Len(ewSessionTblMasterKey)) = ewSessionTblMasterKey Then
'###			d.Add Mid(sName, Len(ewSessionTblMasterKey)+2), Session(sName)
'###		End If
'###	Next
'###	Set ew_CurrentMasterKeys = d

End Function

' Return first master key
Function ew_CurrentMasterKey()

'###	Dim d, k
'###	Set d = ew_CurrentMasterKeys()
'###	If d.Count > 0 Then
'###		k = d.Keys
'###		ew_CurrentMasterKey = d(k(0))
'###	End If
'###	Set d = Nothing

End Function

function extraUnits(mQty,msg)

	if( (Session("specialxprice") & "x"="x") and (Session("FreexQty") &  "x"="x")) then
		'Response.write mQty
		if(mQty>19) then  Response.write msg
	end if

end function
%>
<script language="JScript" runat="server">
// Server-side JScript functions for ASPMaker 6+ (Requires script engine 5.5.+)
// encrytion key
EW_RANDOM_KEY = 'HRaO%X6qPZWeCLEJ';

function ew_Encode(str) {	
	return encodeURIComponent(str);
}

function ew_Decode(str) {	
	return decodeURIComponent(str);	
}
// encrytion key
//EW_RANDOM_KEY = 'i6J65kmW$%MZ&xu4';
// JavaScript implementation of Block TEA by Chris Veness
// http://www.movable-type.co.uk/scripts/TEAblock.html
//
// TEAencrypt: Use Corrected Block TEA to encrypt plaintext using password
//            (note plaintext & password must be strings not string objects)
//
// Return encrypted text as string
//

function TEAencrypt(plaintext, password)
{
    if (plaintext.length == 0) return('');  // nothing to encrypt
    // 'escape' plaintext so chars outside ISO-8859-1 work in single-byte packing, but  
    // keep spaces as spaces (not '%20') so encrypted text doesn't grow too long, and 
    // convert result to longs
    var v = strToLongs(escape(plaintext).replace(/%20/g,' '));
    if (v.length == 1) v[1] = 0;  // algorithm doesn't work for n<2 so fudge by adding nulls
    var k = strToLongs(password.slice(0,16));  // simply convert first 16 chars of password as key
    var n = v.length;
    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = 0;
    while (q-- > 0) {  // 6 + 52/n operations gives between 6 & 32 mixes on each word
        sum += delta;
        e = sum>>>2 & 3;
        for (var p = 0; p < n-1; p++) {
            y = v[p+1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            z = v[p] += mx;
        }
        y = v[0];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        z = v[n-1] += mx;
    }
    // note use of >>> in place of >> due to lack of 'unsigned' type in JavaScript 
    return escCtrlCh(longsToStr(v));
}
//
// TEAdecrypt: Use Corrected Block TEA to decrypt ciphertext using password
//

function TEAdecrypt(ciphertext, password)
{
    if (ciphertext.length == 0) return('');
    var v = strToLongs(unescCtrlCh(ciphertext));
    var k = strToLongs(password.slice(0,16)); 
    var n = v.length;
    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = q*delta;
    while (sum != 0) {
        e = sum>>>2 & 3;
        for (var p = n-1; p > 0; p--) {
            z = v[p-1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            y = v[p] -= mx;
        }
        z = v[n-1];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        y = v[0] -= mx;
        sum -= delta;
    }
    var plaintext = longsToStr(v);
    // strip trailing null chars resulting from filling 4-char blocks:
    if (plaintext.search(/\0/) != -1) plaintext = plaintext.slice(0, plaintext.search(/\0/));
    return unescape(plaintext);
}
// supporting functions

function strToLongs(s) {  // convert string to array of longs, each containing 4 chars
    // note chars must be within ISO-8859-1 (with Unicode code-point < 256) to fit 4/long
    var l = new Array(Math.ceil(s.length/4))
    for (var i=0; i<l.length; i++) {
        // note little-endian encoding - endianness is irrelevant as long as 
        // it is the same in longsToStr() 
        l[i] = s.charCodeAt(i*4) + (s.charCodeAt(i*4+1)<<8) + 
               (s.charCodeAt(i*4+2)<<16) + (s.charCodeAt(i*4+3)<<24);
    }
    return l;  // note running off the end of the string generates nulls since 
}              // bitwise operators treat NaN as 0

function longsToStr(l) {  // convert array of longs back to string
    var a = new Array(l.length);
    for (var i=0; i<l.length; i++) {
        a[i] = String.fromCharCode(l[i] & 0xFF, l[i]>>>8 & 0xFF, 
                                   l[i]>>>16 & 0xFF, l[i]>>>24 & 0xFF);
    }
    return a.join('');  // use Array.join() rather than repeated string appends for efficiency
}

function escCtrlCh(str) {  // escape control chars which might cause problems with encrypted texts
    return str.replace(/[\0\n\v\f\r!]/g, function(c) { return '!' + c.charCodeAt(0) + '!'; });
}

function unescCtrlCh(str) {  // unescape potentially problematic nulls and control characters
    return str.replace(/!\d\d?!/g, function(c) { return String.fromCharCode(c.slice(1,-1)); });
}
</script>
