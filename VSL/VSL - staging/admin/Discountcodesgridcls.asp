<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountcodes_grid

	' Page ID
	Public Property Get PageID()
		PageID = "grid"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Discountcodes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Discountcodes_grid"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Discountcodes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Discountcodes.TableVar & "&" ' add page token
	End Property

	' Message
	Public Property Get Message()
		Message = Session(EW_SESSION_MESSAGE)
	End Property

	Public Property Let Message(v)
		Dim msg
		msg = Session(EW_SESSION_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_MESSAGE) = msg
	End Property

	Public Property Get FailureMessage()
		FailureMessage = Session(EW_SESSION_FAILURE_MESSAGE)
	End Property

	Public Property Let FailureMessage(v)
		Dim msg
		msg = Session(EW_SESSION_FAILURE_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_FAILURE_MESSAGE) = msg
	End Property

	Public Property Get SuccessMessage()
		SuccessMessage = Session(EW_SESSION_SUCCESS_MESSAGE)
	End Property

	Public Property Let SuccessMessage(v)
		Dim msg
		msg = Session(EW_SESSION_SUCCESS_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_SUCCESS_MESSAGE) = msg
	End Property

	' Show Message
	Public Sub ShowMessage()
		Dim sMessage
		sMessage = Message
		Call Message_Showing(sMessage, "")
		If sMessage <> "" Then Response.Write "<p class=""ewMessage"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then Response.Write "<p class=""ewSuccessMessage"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then Response.Write "<p class=""ewErrorMessage"">" & sErrorMessage & "</p>"
		Session(EW_SESSION_FAILURE_MESSAGE) = "" ' Clear message in Session
	End Sub
	Dim PageHeader
	Dim PageFooter

	' Show Page Header
	Public Sub ShowPageHeader()
		Dim sHeader
		sHeader = PageHeader
		Call Page_DataRendering(sHeader)
		If sHeader <> "" Then ' Header exists, display
			Response.Write "<p class=""aspmaker"">" & sHeader & "</p>"
		End If
	End Sub

	' Show Page Footer
	Public Sub ShowPageFooter()
		Dim sFooter
		sFooter = PageFooter
		Call Page_DataRendered(sFooter)
		If sFooter <> "" Then ' Footer exists, display
			Response.Write "<p class=""aspmaker"">" & sFooter & "</p>"
		End If
	End Sub

	' -----------------------
	'  Validate Page request
	'
	Public Function IsPageRequest()
		If Discountcodes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Discountcodes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Discountcodes.TableVar = Request.QueryString("t"))
			End If
		Else
			IsPageRequest = True
		End If
	End Function

	' -----------------------------------------------------------------
	'  Class initialize
	'  - init objects
	'  - open ADO connection
	'
	Private Sub Class_Initialize()
		If IsEmpty(StartTimer) Then StartTimer = Timer ' Init start time

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes

'		Set MasterTable = Table
		Set Table = Discountcodes

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "grid"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Discountcodes"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Initialize list options
		Set ListOptions = New cListOptions
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Init
	'  - called before page main
	'  - check Security
	'  - set up response header
	'  - call page load events
	'
	Sub Page_Init()
		Set Security = New cAdvancedSecurity
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				Discountcodes.GridAddRowCount = gridaddcnt
			End If
		End If

		' Set up list options
		SetupListOptions()

		' Global page loading event (in userfn7.asp)
		Call Page_Loading()

		' Page load event, used in current page
		Call Page_Load()
	End Sub

	' -----------------------------------------------------------------
	'  Class terminate
	'  - clean up page object
	'
	Private Sub Class_Terminate()
		Call Page_Terminate("")
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Terminate
	'  - called when exit page
	'  - clean up ADO connection and objects
	'  - if url specified, redirect to url
	'
	Sub Page_Terminate(url)

'		Set Table = MasterTable
		If url = "" Then
			Exit Sub
		End If

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		Set Security = Nothing
		Set Discountcodes = Nothing

		' Go to url if specified
		If sReDirectUrl <> "" Then
			If Response.Buffer Then Response.Clear
			Response.Redirect sReDirectUrl
		End If
	End Sub

	'
	'  Subroutine Page_Terminate (End)
	' ----------------------------------------

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim SearchWhere
	Dim RecCnt
	Dim EditRowCnt
	Dim RowCnt, RowIndex
	Dim RecPerRow, ColCnt
	Dim KeyCount
	Dim RowAction
	Dim RowOldKey ' Row old key (for copy)
	Dim DbMasterFilter, DbDetailFilter
	Dim MasterRecordExists
	Dim ListOptions
	Dim ExportOptions
	Dim MultiSelectKey
	Dim RestoreSearch
	Dim Recordset, OldRecordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		DisplayRecs = 20
		RecRange = 10
		RecCnt = 0 ' Record count
		KeyCount = 0 ' Key count

		' Search filters
		Dim sSrchAdvanced, sSrchBasic, sFilter
		sSrchAdvanced = "" ' Advanced search filter
		sSrchBasic = "" ' Basic search filter
		SearchWhere = "" ' Search where clause
		sFilter = ""

		' Master/Detail
		DbMasterFilter = "" ' Master filter
		DbDetailFilter = "" ' Detail filter
		If IsPageRequest Then ' Validate request

			' Handle reset command
			ResetCmd()

			' Set up master detail parameters
			SetUpMasterParms()

			' Hide all options
			If Discountcodes.Export <> "" Or Discountcodes.CurrentAction = "gridadd" Or Discountcodes.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
			End If

			' Show grid delete link for grid add / grid edit
			If Discountcodes.AllowAddDeleteRow Then
				If Discountcodes.CurrentAction = "gridadd" Or Discountcodes.CurrentAction = "gridedit" Then
					ListOptions.GetItem("griddelete").Visible = True
				End If
			End If

			' Set Up Sorting Order
			SetUpSortOrder()
		End If ' End Validate Request

		' Restore display records
		If Discountcodes.RecordsPerPage <> "" Then
			DisplayRecs = Discountcodes.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()
		sFilter = ""

		' Restore master/detail filter
		DbMasterFilter = Discountcodes.MasterFilter ' Restore master filter
		DbDetailFilter = Discountcodes.DetailFilter ' Restore detail filter
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)
		Dim RsMaster

		' Load master record
		If Discountcodes.MasterFilter <> "" And Discountcodes.CurrentMasterTable = "DiscountTypes" Then
			Set RsMaster = DiscountTypes.LoadRs(DbMasterFilter)
			MasterRecordExists = Not (RsMaster Is Nothing)
			If Not MasterRecordExists Then
			Else
				Call DiscountTypes.LoadListRowValues(RsMaster)
				DiscountTypes.RowType = EW_ROWTYPE_MASTER ' Master row
				Call DiscountTypes.RenderListRow()
				RsMaster.Close
				Set RsMaster = Nothing
			End If
		End If

		' Set up filter in Session
		Discountcodes.SessionWhere = sFilter
		Discountcodes.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	'  Exit out of inline mode
	'
	Sub ClearInlineMode()
		Discountcodes.LastAction = Discountcodes.CurrentAction ' Save last action
		Discountcodes.CurrentAction = "" ' Clear action
		Session(EW_SESSION_INLINE_MODE) = "" ' Clear inline mode
	End Sub

	' -----------------------------------------------------------------
	' Switch to Grid Add Mode
	'
	Sub GridAddMode()
		Session(EW_SESSION_INLINE_MODE) = "gridadd" ' Enabled grid add
	End Sub

	' -----------------------------------------------------------------
	' Switch to Grid Edit Mode
	'
	Sub GridEditMode()
		Session(EW_SESSION_INLINE_MODE) = "gridedit" ' Enabled grid edit
	End Sub

	' -----------------------------------------------------------------
	' Peform update to grid
	'
	Function GridUpdate()
		Dim rowindex
		Dim bGridUpdate
		Dim sKey, sThisKey
		Dim Rs, RsOld, RsNew, sSql
		rowindex = 1
		bGridUpdate = True

		' Get old recordset
		Discountcodes.CurrentFilter  = BuildKeyFilter()
		sSql = Discountcodes.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Set RsOld = ew_CloneRs(Rs)
		sKey = ""

		' Update row index and get row key
		Dim rowcnt
		ObjForm.Index = 0
		rowcnt = ObjForm.GetValue("key_count")
		If rowcnt = "" Or Not IsNumeric(rowcnt) Then
			rowcnt = 0
		End If

		' Update all rows based on key
		Dim rowkey, rowaction
		For rowindex = 1 to rowcnt
			ObjForm.Index = rowindex
			rowkey = ObjForm.GetValue("k_key") & ""
			rowaction = ObjForm.GetValue("k_action") & ""

			' Load all values & keys
			If rowaction <> "insertdelete" Then ' Skip insert then deleted rows
				Call LoadFormValues() ' Get form values
				If rowaction = "" Or rowaction = "edit" Or rowaction = "delete" Then
					bGridUpdate = SetupKeyValues(rowkey) ' Set up key values
				Else
					bGridUpdate = True
				End If

				' Skip empty row
				If rowaction = "insert" And EmptyRow() Then

					' No action required
				' Validate form and insert/update/delete record

				ElseIf bGridUpdate Then
					If rowaction = "delete" Then
						Discountcodes.CurrentFilter = Discountcodes.KeyFilter
						bGridUpdate = DeleteRows() ' Delete this row
					ElseIf Not ValidateForm() Then
						bGridUpdate = False ' Form error, reset action
						FailureMessage = gsFormError
					Else
						If rowaction = "insert" Then
							bGridUpdate = AddRow(Null) ' Insert this row
						Else
							If rowkey <> "" Then
								Discountcodes.SendEmail = False ' Do not send email on update success

								' Set detail key fields disabled flag to skip update
								If Discountcodes.CurrentMasterTable = "DiscountTypes" Then
									Discountcodes.DiscountTypeId.Disabled = True ' Set field disabled flag to skip update
								End If
								bGridUpdate = EditRow() ' Update this row

								' Reset detail key fields disabled flag
								If Discountcodes.CurrentMasterTable = "DiscountTypes" Then
									Discountcodes.DiscountTypeId.Disabled = False ' Reset field disabled flag
								End If
							End If
						End If ' End update
					End If
				End If
				If bGridUpdate Then
					If sKey <> "" Then sKey = sKey & ", "
					sKey = sKey & rowkey
				Else
					Exit For
				End If
			End If
		Next
		If bGridUpdate Then

			' Get new recordset
			Set Rs = Conn.Execute(sSql)
			Set RsNew = ew_CloneRs(Rs)
			Call ClearInlineMode() ' Clear inline edit mode
		Else
			If FailureMessage = "" Then
				FailureMessage = Language.Phrase("UpdateFailed") ' Set update failed message
			End If
			Discountcodes.EventCancelled = True ' Set event cancelled
			Discountcodes.CurrentAction = "gridedit" ' Stay in gridedit mode
		End If
		Set Rs = Nothing
		Set RsOld = Nothing
		Set RsNew = Nothing
		GridUpdate = bGridUpdate
	End Function

	' -----------------------------------------------------------------
	'  Build filter for all keys
	'
	Function BuildKeyFilter()
		Dim rowindex, sThisKey
		Dim sKey
		Dim sWrkFilter, sFilter
		sWrkFilter = ""

		' Update row index and get row key
		rowindex = 1
		ObjForm.Index = rowindex
		sThisKey = ObjForm.GetValue("k_key") & ""
		Do While (sThisKey <> "")
			If SetupKeyValues(sThisKey) Then
				sFilter = Discountcodes.KeyFilter
				If sWrkFilter <> "" Then sWrkFilter = sWrkFilter & " OR "
				sWrkFilter = sWrkFilter & sFilter
			Else
				sWrkFilter = "0=1"
				Exit Do
			End If

			' Update row index and get row key
			rowindex = rowindex + 1 ' next row
			ObjForm.Index = rowindex
			sThisKey = ObjForm.GetValue("k_key") & ""
		Loop
		BuildKeyFilter = sWrkFilter
	End Function

	' -----------------------------------------------------------------
	' Set up key values
	'
	Function SetupKeyValues(key)
		Dim arrKeyFlds
		arrKeyFlds = Split(key&"", EW_COMPOSITE_KEY_SEPARATOR)
		If UBound(arrKeyFlds) >= 0 Then
			Discountcodes.Discountid.FormValue = arrKeyFlds(0)
			If Not IsNumeric(Discountcodes.Discountid.FormValue) Then
				SetupKeyValues = False
				Exit Function
			End If
		End If
		SetupKeyValues = True
	End Function

	' Grid Insert
	' Peform insert to grid
	Function GridInsert()
		Dim addcnt
		Dim rowindex, rowcnt
		Dim bGridInsert
		Dim sSql, sWrkFilter, sFilter, sKey, sThisKey
		Dim Rs, RsNew
		rowindex = 1
		bGridInsert = False

		' Init key filter
		sWrkFilter = ""
		addcnt = 0
		sKey = ""

		' Get row count
		ObjForm.Index = 0
		rowcnt = ObjForm.GetValue("key_count") & ""
		If rowcnt = "" Or Not IsNumeric(rowcnt) Then rowcnt = 0

		' Insert all rows
		For rowindex = 1 to rowcnt

			' Load current row values
			ObjForm.Index = rowindex
			Dim rowaction
			rowaction = ObjForm.GetValue("k_action") & ""
			If rowaction = "" Or rowaction = "insert" Then
				If rowaction = "insert" Then
					RowOldKey = ObjForm.GetValue("k_oldkey") & ""
					LoadOldRecord() ' Load old recordset
				End If
				Call LoadFormValues() ' Get form values
				If Not EmptyRow() Then
					addcnt = addcnt + 1
					Discountcodes.SendEmail = False ' Do not send email on insert success

					' Validate Form
					If Not ValidateForm() Then
						bGridInsert = False ' Form error, reset action
						FailureMessage = gsFormError
					Else
						bGridInsert = AddRow(OldRecordset) ' Insert this row
					End If
					If bGridInsert Then
						If sKey <> "" Then sKey = sKey & EW_COMPOSITE_KEY_SEPARATOR
						sKey = sKey & Discountcodes.Discountid.CurrentValue

						' Add filter for this record
						sFilter = Discountcodes.KeyFilter
						If sWrkFilter <> "" Then sWrkFilter = sWrkFilter & " OR "
						sWrkFilter = sWrkFilter & sFilter
					Else
						Exit For
					End If
				End If
			End If
		Next
		If addcnt = 0 Then ' No record inserted
			Call ClearInlineMode() ' Clear grid add mode and return
			GridInsert = True
			Exit Function
		End If
		If bGridInsert Then

			' Get new recordset
			Discountcodes.CurrentFilter  = sWrkFilter
			sSql = Discountcodes.SQL
			Set Rs = Conn.Execute(sSql)
			Set RsNew = ew_CloneRs(Rs)
			Call ClearInlineMode() ' Clear grid add mode
		Else
			If FailureMessage = "" Then
				FailureMessage = Language.Phrase("InsertFailed") ' Set insert failed message
			End If
			Discountcodes.EventCancelled = True ' Set event cancelled
			Discountcodes.CurrentAction = "gridadd" ' Stay in gridadd mode
		End If
		Set Rs = Nothing
		Set RsNew = Nothing
		GridInsert = bGridInsert
	End Function

	' Check if empty row
	Function EmptyRow()
		EmptyRow = True
		If EmptyRow And ObjForm.HasValue("x_DiscountCode") And ObjForm.HasValue("o_DiscountCode") Then EmptyRow = (Discountcodes.DiscountCode.CurrentValue&"" = Discountcodes.DiscountCode.OldValue&"")
		If EmptyRow And ObjForm.HasValue("x_Active") And ObjForm.HasValue("o_Active") Then EmptyRow = (Discountcodes.Active.CurrentValue&"" = Discountcodes.Active.OldValue&"")
		If EmptyRow And ObjForm.HasValue("x_used") And ObjForm.HasValue("o_used") Then EmptyRow = (Discountcodes.used.CurrentValue&"" = Discountcodes.used.OldValue&"")
		If EmptyRow And ObjForm.HasValue("x_OrderId") And ObjForm.HasValue("o_OrderId") Then EmptyRow = (Discountcodes.OrderId.CurrentValue&"" = Discountcodes.OrderId.OldValue&"")
		If EmptyRow And ObjForm.HasValue("x_Use_date") And ObjForm.HasValue("o_Use_date") Then EmptyRow = (Discountcodes.Use_date.CurrentValue&"" = Discountcodes.Use_date.OldValue&"")
		If EmptyRow And ObjForm.HasValue("x_DiscountTypeId") And ObjForm.HasValue("o_DiscountTypeId") Then EmptyRow = (Discountcodes.DiscountTypeId.CurrentValue&"" = Discountcodes.DiscountTypeId.OldValue&"")
	End Function

	' Validate grid form
	Function ValidateGridForm()
		Dim rowindex, rowcnt, rowaction

		' Get row count
		ObjForm.Index = 0
		rowcnt = ObjForm.GetValue("key_count")&""
		If rowcnt = "" Or Not IsNumeric(rowcnt) Then
			rowcnt = 0
		End If

		' Validate all records
		ValidateGridForm = True
		For rowindex = 1 to rowcnt

			' Load current row values
			ObjForm.Index = rowindex
			rowaction = ObjForm.GetValue("k_action") & ""
			If rowaction <> "delete" And rowaction <> "insertdelete" Then
				LoadFormValues() ' Get form values
				If rowaction = "insert" And EmptyRow() Then

					' Ignore
				ElseIf Not ValidateForm() Then
					ValidateGridForm = False
					Exit For
				End If
			End If
		Next
	End Function

	' -----------------------------------------------------------------
	' Restore form values for current row
	'
	Sub RestoreCurrentRowFormValues(idx)

		' Get row based on current index
		ObjForm.Index = idx
		Call LoadFormValues() ' Load form values
	End Sub

	' -----------------------------------------------------------------
	' Set up Sort parameters based on Sort Links clicked
	'
	Sub SetUpSortOrder()
		Dim sOrderBy
		Dim sSortField, sLastSort, sThisSort
		Dim bCtrl

		' Check for an Order parameter
		If Request.QueryString("order").Count > 0 Then
			Discountcodes.CurrentOrder = Request.QueryString("order")
			Discountcodes.CurrentOrderType = Request.QueryString("ordertype")
			Discountcodes.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Discountcodes.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Discountcodes.SqlOrderBy <> "" Then
				sOrderBy = Discountcodes.SqlOrderBy
				Discountcodes.SessionOrderBy = sOrderBy
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Reset command based on querystring parameter cmd=
	' - RESET: reset search parameters
	' - RESETALL: reset search & master/detail parameters
	' - RESETSORT: reset sort parameters
	'
	Sub ResetCmd()
		Dim sCmd

		' Get reset cmd
		If Request.QueryString("cmd").Count > 0 Then
			sCmd = Request.QueryString("cmd")

			' Reset master/detail keys
			If LCase(sCmd) = "resetall" Then
				Discountcodes.CurrentMasterTable = "" ' Clear master table
				DbMasterFilter = ""
				DbDetailFilter = ""
				Discountcodes.DiscountTypeId.SessionValue = ""
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				Discountcodes.SessionOrderBy = sOrderBy
			End If

			' Reset start position
			StartRec = 1
			Discountcodes.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item

		' "griddelete"
		If Discountcodes.AllowAddDeleteRow Then
			ListOptions.Add("griddelete")
			Set item = ListOptions.GetItem("griddelete")
			item.CssStyle = "white-space: nowrap;"
			item.OnLeft = True
			item.Visible = False ' Default hidden
		End If
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()

		' Set up row action and key
		If IsNumeric(RowIndex) Then
			ObjForm.Index = RowIndex
			If RowAction <> "" Then
				MultiSelectKey = MultiSelectKey & "<input type=""hidden"" name=""k" & RowIndex & "_action"" id=""k" & RowIndex & "_action"" value=""" & RowAction & """>"
			End If
			If ObjForm.HasValue("k_oldkey") Then
				RowOldKey = ObjForm.GetValue("k_oldkey") & ""
			End If
			If RowOldKey <> "" Then
				MultiSelectKey = MultiSelectKey & "<input type=""hidden"" name=""k" & RowIndex & "_oldkey"" id=""k" & RowIndex & "_oldkey"" value = """ & ew_HtmlEncode(RowOldKey) & """>"
			End If
			If RowAction = "delete" Then
				Dim sKey
				sKey = ObjForm.GetValue("k_key") & ""
				Call SetupKeyValues(sKey)
			End If
		End If

		' "delete"
		If Discountcodes.AllowAddDeleteRow Then
			If Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Or Discountcodes.CurrentMode = "edit" Then
				Set item = ListOptions.GetItem("griddelete")
				item.Body = "<a class=""ewGridLink"" href=""javascript:void(0);"" onclick=""ew_DeleteGridRow(this, Discountcodes_grid, " & RowIndex & ");"">" & "<img src=""images/delete.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
			End If
		End If
		If Discountcodes.CurrentMode = "edit" And RowIndex <> "" And IsNumeric(RowIndex) Then
			MultiSelectKey = MultiSelectKey & "<input type=""hidden"" name=""k" & RowIndex & "_key"" id=""k" & RowIndex & "_key"" value=""" & Discountcodes.Discountid.CurrentValue & """>"
		End If
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	' Set record key
	Function SetRecordKey(rs)
		Dim key
		key = ""
		SetRecordKey = key
		If rs.Eof Then Exit Function
		If (key <> "") Then key = key & EW_COMPOSITE_KEY_SEPARATOR
		key = key & rs("Discountid")
		SetRecordKey = key
	End Function

	Function RenderListOptionsExt()
	End Function
	Dim Pager

	' -----------------------------------------------------------------
	' Set up Starting Record parameters based on Pager Navigation
	'
	Sub SetUpStartRec()
		Dim PageNo

		' Exit if DisplayRecs = 0
		If DisplayRecs = 0 Then Exit Sub
		If IsPageRequest Then ' Validate request

			' Check for a START parameter
			If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
				StartRec = Request.QueryString(EW_TABLE_START_REC)
				Discountcodes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Discountcodes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Discountcodes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Discountcodes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Discountcodes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Discountcodes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
		Dim index, confirmPage
		index = ObjForm.Index ' Save form index
		ObjForm.Index = 0
		confirmPage = (ObjForm.GetValue("a_confirm") & "" <> "")
		ObjForm.Index = index ' Restore form index
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		Discountcodes.DiscountCode.CurrentValue = Null
		Discountcodes.DiscountCode.OldValue = Discountcodes.DiscountCode.CurrentValue
		Discountcodes.Active.CurrentValue = "0"
		Discountcodes.Active.OldValue = Discountcodes.Active.CurrentValue
		Discountcodes.used.CurrentValue = "0"
		Discountcodes.used.OldValue = Discountcodes.used.CurrentValue
		Discountcodes.OrderId.CurrentValue = Null
		Discountcodes.OrderId.OldValue = Discountcodes.OrderId.CurrentValue
		Discountcodes.Use_date.CurrentValue = Null
		Discountcodes.Use_date.OldValue = Discountcodes.Use_date.CurrentValue
		Discountcodes.DiscountTypeId.CurrentValue = Null
		Discountcodes.DiscountTypeId.OldValue = Discountcodes.DiscountTypeId.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not Discountcodes.DiscountCode.FldIsDetailKey Then Discountcodes.DiscountCode.FormValue = ObjForm.GetValue("x_DiscountCode")
		Discountcodes.DiscountCode.OldValue = ObjForm.GetValue("o_DiscountCode")
		If Not Discountcodes.Active.FldIsDetailKey Then Discountcodes.Active.FormValue = ObjForm.GetValue("x_Active")
		Discountcodes.Active.OldValue = ObjForm.GetValue("o_Active")
		If Not Discountcodes.used.FldIsDetailKey Then Discountcodes.used.FormValue = ObjForm.GetValue("x_used")
		Discountcodes.used.OldValue = ObjForm.GetValue("o_used")
		If Not Discountcodes.OrderId.FldIsDetailKey Then Discountcodes.OrderId.FormValue = ObjForm.GetValue("x_OrderId")
		Discountcodes.OrderId.OldValue = ObjForm.GetValue("o_OrderId")
		If Not Discountcodes.Use_date.FldIsDetailKey Then Discountcodes.Use_date.FormValue = ObjForm.GetValue("x_Use_date")
		If Not Discountcodes.Use_date.FldIsDetailKey Then Discountcodes.Use_date.CurrentValue = ew_UnFormatDateTime(Discountcodes.Use_date.CurrentValue, 8)
		Discountcodes.Use_date.OldValue = ObjForm.GetValue("o_Use_date")
		If Not Discountcodes.DiscountTypeId.FldIsDetailKey Then Discountcodes.DiscountTypeId.FormValue = ObjForm.GetValue("x_DiscountTypeId")
		Discountcodes.DiscountTypeId.OldValue = ObjForm.GetValue("o_DiscountTypeId")
		If Not Discountcodes.Discountid.FldIsDetailKey And Discountcodes.CurrentAction <> "gridadd" And Discountcodes.CurrentAction <> "add" Then Discountcodes.Discountid.FormValue = ObjForm.GetValue("x_Discountid")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Discountcodes.DiscountCode.CurrentValue = Discountcodes.DiscountCode.FormValue
		Discountcodes.Active.CurrentValue = Discountcodes.Active.FormValue
		Discountcodes.used.CurrentValue = Discountcodes.used.FormValue
		Discountcodes.OrderId.CurrentValue = Discountcodes.OrderId.FormValue
		Discountcodes.Use_date.CurrentValue = Discountcodes.Use_date.FormValue
		Discountcodes.Use_date.CurrentValue = ew_UnFormatDateTime(Discountcodes.Use_date.CurrentValue, 8)
		Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.FormValue
		If Discountcodes.CurrentAction <> "gridadd" And Discountcodes.CurrentAction <> "add" Then Discountcodes.Discountid.CurrentValue = Discountcodes.Discountid.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Discountcodes.CurrentFilter
		Call Discountcodes.Recordset_Selecting(sFilter)
		Discountcodes.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Discountcodes.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Discountcodes.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Discountcodes.KeyFilter

		' Call Row Selecting event
		Call Discountcodes.Row_Selecting(sFilter)

		' Load sql based on filter
		Discountcodes.CurrentFilter = sFilter
		sSql = Discountcodes.SQL
		Call ew_SetDebugMsg("LoadRow: " & sSql) ' Show SQL for debugging
		Set RsRow = ew_LoadRow(sSql)
		If RsRow.Eof Then
			LoadRow = False
		Else
			LoadRow = True
			RsRow.MoveFirst
			Call LoadRowValues(RsRow) ' Load row values
		End If
		RsRow.Close
		Set RsRow = Nothing
	End Function

	' -----------------------------------------------------------------
	' Load row values from recordset
	'
	Sub LoadRowValues(RsRow)
		Dim sDetailFilter
		If RsRow.Eof Then Exit Sub

		' Call Row Selected event
		Call Discountcodes.Row_Selected(RsRow)
		Discountcodes.Discountid.DbValue = RsRow("Discountid")
		Discountcodes.DiscountCode.DbValue = RsRow("DiscountCode")
		Discountcodes.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Discountcodes.used.DbValue = ew_IIf(RsRow("used"), "1", "0")
		Discountcodes.OrderId.DbValue = RsRow("OrderId")
		Discountcodes.Use_date.DbValue = RsRow("Use_date")
		Discountcodes.DiscountTypeId.DbValue = RsRow("DiscountTypeId")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		Dim arKeys, cnt
		ReDim arKeys(0)
		arKeys(0) = RowOldKey
		cnt = UBound(arKeys)+1
		If cnt >= 1 Then
			If arKeys(0) & "" <> "" Then
				Discountcodes.Discountid.CurrentValue = arKeys(0) & "" ' Discountid
			Else
				bValidKey = False
			End If
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Discountcodes.CurrentFilter = Discountcodes.KeyFilter
			Dim sSql
			sSql = Discountcodes.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Discountcodes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Discountid
		' DiscountCode
		' Active
		' used
		' OrderId
		' Use_date
		' DiscountTypeId
		' -----------
		'  View  Row
		' -----------

		If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Discountid
			Discountcodes.Discountid.ViewValue = Discountcodes.Discountid.CurrentValue
			Discountcodes.Discountid.ViewCustomAttributes = ""

			' DiscountCode
			Discountcodes.DiscountCode.ViewValue = Discountcodes.DiscountCode.CurrentValue
			Discountcodes.DiscountCode.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			Else
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			End If
			Discountcodes.Active.ViewCustomAttributes = ""

			' used
			If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			Else
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			End If
			Discountcodes.used.ViewCustomAttributes = ""

			' OrderId
			Discountcodes.OrderId.ViewValue = Discountcodes.OrderId.CurrentValue
			Discountcodes.OrderId.ViewCustomAttributes = ""

			' Use_date
			Discountcodes.Use_date.ViewValue = Discountcodes.Use_date.CurrentValue
			Discountcodes.Use_date.ViewCustomAttributes = ""

			' DiscountTypeId
			If Discountcodes.DiscountTypeId.CurrentValue & "" <> "" Then
				sFilterWrk = "[DiscountTypeId] = " & ew_AdjustSql(Discountcodes.DiscountTypeId.CurrentValue) & ""
			sSqlWrk = "SELECT [DiscountType] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Discountcodes.DiscountTypeId.ViewValue = RsWrk("DiscountType")
				Else
					Discountcodes.DiscountTypeId.ViewValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Discountcodes.DiscountTypeId.ViewValue = Null
			End If
			Discountcodes.DiscountTypeId.ViewCustomAttributes = ""

			' View refer script
			' DiscountCode

			Discountcodes.DiscountCode.LinkCustomAttributes = ""
			Discountcodes.DiscountCode.HrefValue = ""
			Discountcodes.DiscountCode.TooltipValue = ""

			' Active
			Discountcodes.Active.LinkCustomAttributes = ""
			Discountcodes.Active.HrefValue = ""
			Discountcodes.Active.TooltipValue = ""

			' used
			Discountcodes.used.LinkCustomAttributes = ""
			Discountcodes.used.HrefValue = ""
			Discountcodes.used.TooltipValue = ""

			' OrderId
			Discountcodes.OrderId.LinkCustomAttributes = ""
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.ViewValue<>"", Discountcodes.OrderId.ViewValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If
			Discountcodes.OrderId.TooltipValue = ""

			' Use_date
			Discountcodes.Use_date.LinkCustomAttributes = ""
			Discountcodes.Use_date.HrefValue = ""
			Discountcodes.Use_date.TooltipValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.LinkCustomAttributes = ""
			Discountcodes.DiscountTypeId.HrefValue = ""
			Discountcodes.DiscountTypeId.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add row

			' DiscountCode
			Discountcodes.DiscountCode.EditCustomAttributes = ""
			Discountcodes.DiscountCode.EditValue = ew_HtmlEncode(Discountcodes.DiscountCode.CurrentValue)

			' Active
			Discountcodes.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			Discountcodes.Active.EditValue = arwrk

			' used
			Discountcodes.used.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			Discountcodes.used.EditValue = arwrk

			' OrderId
			Discountcodes.OrderId.EditCustomAttributes = ""
			Discountcodes.OrderId.EditValue = ew_HtmlEncode(Discountcodes.OrderId.CurrentValue)

			' Use_date
			Discountcodes.Use_date.EditCustomAttributes = ""
			Discountcodes.Use_date.EditValue = Discountcodes.Use_date.CurrentValue

			' DiscountTypeId
			Discountcodes.DiscountTypeId.EditCustomAttributes = ""
			If Discountcodes.DiscountTypeId.SessionValue <> "" Then
				Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.SessionValue
				Discountcodes.DiscountTypeId.OldValue = Discountcodes.DiscountTypeId.CurrentValue
			If Discountcodes.DiscountTypeId.CurrentValue & "" <> "" Then
				sFilterWrk = "[DiscountTypeId] = " & ew_AdjustSql(Discountcodes.DiscountTypeId.CurrentValue) & ""
			sSqlWrk = "SELECT [DiscountType] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Discountcodes.DiscountTypeId.ViewValue = RsWrk("DiscountType")
				Else
					Discountcodes.DiscountTypeId.ViewValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Discountcodes.DiscountTypeId.ViewValue = Null
			End If
			Discountcodes.DiscountTypeId.ViewCustomAttributes = ""
			Else
			End If

			' Edit refer script
			' DiscountCode

			Discountcodes.DiscountCode.HrefValue = ""

			' Active
			Discountcodes.Active.HrefValue = ""

			' used
			Discountcodes.used.HrefValue = ""

			' OrderId
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.EditValue<>"", Discountcodes.OrderId.EditValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If

			' Use_date
			Discountcodes.Use_date.HrefValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.HrefValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' DiscountCode
			Discountcodes.DiscountCode.EditCustomAttributes = ""
			Discountcodes.DiscountCode.EditValue = ew_HtmlEncode(Discountcodes.DiscountCode.CurrentValue)

			' Active
			Discountcodes.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			Discountcodes.Active.EditValue = arwrk

			' used
			Discountcodes.used.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			Discountcodes.used.EditValue = arwrk

			' OrderId
			Discountcodes.OrderId.EditCustomAttributes = ""
			Discountcodes.OrderId.EditValue = ew_HtmlEncode(Discountcodes.OrderId.CurrentValue)

			' Use_date
			Discountcodes.Use_date.EditCustomAttributes = ""
			Discountcodes.Use_date.EditValue = Discountcodes.Use_date.CurrentValue

			' DiscountTypeId
			Discountcodes.DiscountTypeId.EditCustomAttributes = ""
			If Discountcodes.DiscountTypeId.SessionValue <> "" Then
				Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.SessionValue
				Discountcodes.DiscountTypeId.OldValue = Discountcodes.DiscountTypeId.CurrentValue
			If Discountcodes.DiscountTypeId.CurrentValue & "" <> "" Then
				sFilterWrk = "[DiscountTypeId] = " & ew_AdjustSql(Discountcodes.DiscountTypeId.CurrentValue) & ""
			sSqlWrk = "SELECT [DiscountType] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Discountcodes.DiscountTypeId.ViewValue = RsWrk("DiscountType")
				Else
					Discountcodes.DiscountTypeId.ViewValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Discountcodes.DiscountTypeId.ViewValue = Null
			End If
			Discountcodes.DiscountTypeId.ViewCustomAttributes = ""
			Else
			End If

			' Edit refer script
			' DiscountCode

			Discountcodes.DiscountCode.HrefValue = ""

			' Active
			Discountcodes.Active.HrefValue = ""

			' used
			Discountcodes.used.HrefValue = ""

			' OrderId
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.EditValue<>"", Discountcodes.OrderId.EditValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If

			' Use_date
			Discountcodes.Use_date.HrefValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.HrefValue = ""
		End If
		If Discountcodes.RowType = EW_ROWTYPE_ADD Or Discountcodes.RowType = EW_ROWTYPE_EDIT Or Discountcodes.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Discountcodes.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Discountcodes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Discountcodes.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(Discountcodes.OrderId.FormValue) Then
			Call ew_AddMessage(gsFormError, Discountcodes.OrderId.FldErrMsg)
		End If
		If Not ew_CheckDate(Discountcodes.Use_date.FormValue) Then
			Call ew_AddMessage(gsFormError, Discountcodes.Use_date.FldErrMsg)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld
		DeleteRows = True
		sSql = Discountcodes.SQL
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = Discountcodes.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("Discountid")
				RsDelete.Delete
				If Err.Number <> 0 Then
					FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If Discountcodes.CancelMessage <> "" Then
				FailureMessage = Discountcodes.CancelMessage
				Discountcodes.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
		Else
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			If Not RsOld.Eof Then RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call Discountcodes.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

	' -----------------------------------------------------------------
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = Discountcodes.KeyFilter
		Discountcodes.CurrentFilter  = sFilter
		sSql = Discountcodes.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field DiscountCode
			Call Discountcodes.DiscountCode.SetDbValue(Rs, Discountcodes.DiscountCode.CurrentValue, Null, Discountcodes.DiscountCode.ReadOnly)

			' Field Active
			boolwrk = Discountcodes.Active.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Discountcodes.Active.SetDbValue(Rs, boolwrk, Null, Discountcodes.Active.ReadOnly)

			' Field used
			boolwrk = Discountcodes.used.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Discountcodes.used.SetDbValue(Rs, boolwrk, Null, Discountcodes.used.ReadOnly)

			' Field OrderId
			Call Discountcodes.OrderId.SetDbValue(Rs, Discountcodes.OrderId.CurrentValue, Null, Discountcodes.OrderId.ReadOnly)

			' Field Use_date
			Call Discountcodes.Use_date.SetDbValue(Rs, Discountcodes.Use_date.CurrentValue, Null, Discountcodes.Use_date.ReadOnly)

			' Field DiscountTypeId
			Call Discountcodes.DiscountTypeId.SetDbValue(Rs, Discountcodes.DiscountTypeId.CurrentValue, Null, Discountcodes.DiscountTypeId.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Discountcodes.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				Rs.Update
				If Err.Number <> 0 Then
					FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
			Else
				Rs.CancelUpdate
				If Discountcodes.CancelMessage <> "" Then
					FailureMessage = Discountcodes.CancelMessage
					Discountcodes.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Discountcodes.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' -----------------------------------------------------------------
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear

		' Set up foreign key field value from Session
		If Discountcodes.CurrentMasterTable = "DiscountTypes" Then
			Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.SessionValue
		End If

		' Add new record
		sFilter = "(0 = 1)"
		Discountcodes.CurrentFilter = sFilter
		sSql = Discountcodes.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field DiscountCode
		Call Discountcodes.DiscountCode.SetDbValue(Rs, Discountcodes.DiscountCode.CurrentValue, Null, False)

		' Field Active
		boolwrk = Discountcodes.Active.CurrentValue
		If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
		Call Discountcodes.Active.SetDbValue(Rs, boolwrk, Null, (Discountcodes.Active.CurrentValue&"" = ""))

		' Field used
		boolwrk = Discountcodes.used.CurrentValue
		If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
		Call Discountcodes.used.SetDbValue(Rs, boolwrk, Null, (Discountcodes.used.CurrentValue&"" = ""))

		' Field OrderId
		Call Discountcodes.OrderId.SetDbValue(Rs, Discountcodes.OrderId.CurrentValue, Null, False)

		' Field Use_date
		Call Discountcodes.Use_date.SetDbValue(Rs, Discountcodes.Use_date.CurrentValue, Null, False)

		' Field DiscountTypeId
		Call Discountcodes.DiscountTypeId.SetDbValue(Rs, Discountcodes.DiscountTypeId.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = Discountcodes.Row_Inserting(RsOld, Rs)
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
		Else
			Rs.CancelUpdate
			If Discountcodes.CancelMessage <> "" Then
				FailureMessage = Discountcodes.CancelMessage
				Discountcodes.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			Discountcodes.Discountid.DbValue = RsNew("Discountid")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call Discountcodes.Row_Inserted(RsOld, RsNew)
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' -----------------------------------------------------------------
	' Set up Master Detail based on querystring parameter
	'
	Sub SetUpMasterParms()
		Dim bValidMaster, sMasterTblVar

		' Hide foreign keys
		sMasterTblVar = Discountcodes.CurrentMasterTable
		If sMasterTblVar = "DiscountTypes" Then
			Discountcodes.DiscountTypeId.Visible = False
			If DiscountTypes.EventCancelled Then Discountcodes.EventCancelled = True
		End If
		DbMasterFilter = Discountcodes.MasterFilter '  Get master filter
		DbDetailFilter = Discountcodes.DetailFilter ' Get detail filter
	End Sub

	' Page Load event
	Sub Page_Load()

		'Response.Write "Page Load"
	End Sub

	' Page Unload event
	Sub Page_Unload()

		'Response.Write "Page Unload"
	End Sub

	' Page Redirecting event
	Sub Page_Redirecting(url)

		'url = newurl
	End Sub

	' Message Showing event
	' typ = ""|"success"|"failure"
	Sub Message_Showing(msg, typ)

		' Example:
		'If typ = "success" Then msg = "your success message"

	End Sub

	' Page Data Rendering event
	Sub Page_DataRendering(header)

		' Example:
		'header = "your header"

	End Sub

	' Page Data Rendered event
	Sub Page_DataRendered(footer)

		' Example:
		'footer = "your footer"

	End Sub

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function

	' ListOptions Load event
	Sub ListOptions_Load()

		'Example: 
		' Dim opt
		' Set opt = ListOptions.Add("new")
		' opt.OnLeft = True ' Link on left
		' opt.MoveTo 0 ' Move to first column

	End Sub

	' ListOptions Rendered event
	Sub ListOptions_Rendered()

		'Example: 
		'ListOptions.GetItem("new").Body = "xxx"

	End Sub
End Class
%>
