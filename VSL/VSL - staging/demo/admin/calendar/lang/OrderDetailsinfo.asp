<%

' ASPMaker configuration for Table OrderDetails
Dim OrderDetails

' Define table class
Class cOrderDetails

	' Class Initialize
	Private Sub Class_Initialize()
		UseTokenInUrl = EW_USE_TOKEN_IN_URL
		ExportOriginalValue = EW_EXPORT_ORIGINAL_VALUE
		ExportAll = True
		Set RowAttrs = New cAttributes ' Row attributes
		AllowAddDeleteRow = ew_AllowAddDeleteRow() ' Allow add/delete row
		DetailAdd = False ' Allow detail add
		DetailEdit = False ' Allow detail edit
		GridAddRowCount = 5 ' Grid add row count
		Call ew_SetArObj(Fields, "OrderDetailsId", OrderDetailsId)
		Call ew_SetArObj(Fields, "OrderId", OrderId)
		Call ew_SetArObj(Fields, "ProductId", ProductId)
		Call ew_SetArObj(Fields, "Quantity", Quantity)
		Call ew_SetArObj(Fields, "Price", Price)
	End Sub

	' Reset attributes for table object
	Public Sub ResetAttrs()
		CssClass = ""
		CssStyle = ""
		RowAttrs.Clear()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				Call fld.ResetAttrs()
			Next
		End If
	End Sub

	' Setup field titles
	Public Sub SetupFieldTitles()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				If fld.FldTitle <> "" Then
					fld.EditAttrs.AddAttribute "onmouseover", "ew_ShowTitle(this, '" & ew_JsEncode3(fld.FldTitle) & "');", True
					fld.EditAttrs.AddAttribute "onmouseout", "ew_HideTooltip();", True
				End If
			Next
		End If
	End Sub

	' Define table level constants
	' Use table token in Url

	Dim UseTokenInUrl

	' Table variable
	Public Property Get TableVar()
		TableVar = "OrderDetails"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "OrderDetails"
	End Property

	' Table type
	Public Property Get TableType()
		TableType = "TABLE"
	End Property

	' Table caption
	Public Property Get TableCaption()
		TableCaption = Language.TablePhrase(TableVar, "TblCaption")
	End Property

	' Page caption
	Public Property Get PageCaption(Page)
		PageCaption = Language.TablePhrase(TableVar, "TblPageCaption" & Page)
		If PageCaption = "" Then PageCaption = "Page " & Page
	End Property

	' Export Return Page
	Public Property Get ExportReturnUrl()
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) <> "" Then
			ExportReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL)
		Else
			ExportReturnUrl = ew_CurrentPage
		End If
	End Property

	Public Property Let ExportReturnUrl(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) = v
	End Property

	' Records per page
	Public Property Get RecordsPerPage()
		RecordsPerPage = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE)
	End Property

	Public Property Let RecordsPerPage(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE) = v
	End Property

	' Start record number
	Public Property Get StartRecordNumber()
		StartRecordNumber = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC)
	End Property

	Public Property Let StartRecordNumber(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC) = v
	End Property

	' Search Highlight Name
	Public Property Get HighlightName()
		HighlightName = "OrderDetails_Highlight"
	End Property

	' Advanced search
	Public Function GetAdvancedSearch(fld)
		GetAdvancedSearch = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld)
	End Function

	Public Function SetAdvancedSearch(fld, v)
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) <> v Then
			Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) = v
		End If
	End Function
	Dim BasicSearchKeyword
	Dim BasicSearchType

	' Basic search Keyword
	Public Property Get SessionBasicSearchKeyword()
		SessionBasicSearchKeyword = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH)
	End Property

	Public Property Let SessionBasicSearchKeyword(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH) = v
	End Property

	' Basic Search Type
	Public Property Get SessionBasicSearchType()
		SessionBasicSearchType = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH_TYPE)
	End Property

	Public Property Let SessionBasicSearchType(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH_TYPE) = v
	End Property

	' Search where clause
	Public Property Get SearchWhere()
		SearchWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE)
	End Property

	Public Property Let SearchWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE) = v
	End Property

	' Single column sort
	Public Sub UpdateSort(ofld)
		Dim sSortField, sLastSort, sThisSort
		If CurrentOrder = ofld.FldName Then
			sSortField = ofld.FldExpression
			sLastSort = ofld.Sort
			If CurrentOrderType = "ASC" Or CurrentOrderType = "DESC" Then
				sThisSort = CurrentOrderType
			Else
				If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			End If
			ofld.Sort = sThisSort
			SessionOrderBy = sSortField & " " & sThisSort ' Save to Session
		Else
			ofld.Sort = ""
		End If
	End Sub

	' Session WHERE Clause
	Public Property Get SessionWhere()
		SessionWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE)
	End Property

	Public Property Let SessionWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE) = v
	End Property

	' Session ORDER BY
	Public Property Get SessionOrderBy()
		SessionOrderBy = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY)
	End Property

	Public Property Let SessionOrderBy(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY) = v
	End Property

	' Session Key
	Public Function GetKey(fld)
		GetKey = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld)
	End Function

	Public Function SetKey(fld, v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld) = v
	End Function

	' Current master table name
	Public Property Get CurrentMasterTable()
		CurrentMasterTable = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_MASTER_TABLE)
	End Property

	Public Property Let CurrentMasterTable(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_MASTER_TABLE) = v
	End Property

	' Current master table object
	Public Property Get MasterTable()
		If CurrentMasterTable = "Orders" Then
			Set MasterTable = Orders
			Exit Property
		End If
		Set MasterTable = Nothing
	End Property

	' Session master where clause
	Public Property Get MasterFilter()

		' Master filter
		Dim sMasterFilter
		sMasterFilter = ""
		If CurrentMasterTable = "Orders" Then
			If OrderId.SessionValue <> "" Then
				sMasterFilter = sMasterFilter & "[OrderId]=" & ew_QuotedValue(OrderId.SessionValue, EW_DATATYPE_NUMBER)
			Else
				MasterFilter = ""
				Exit Property
			End If
		End If
		MasterFilter = sMasterFilter
	End Property

	' Session detail where clause
	Public Property Get DetailFilter()

		' Detail filter
		Dim sDetailFilter
		sDetailFilter = ""
		If CurrentMasterTable = "Orders" Then
			If OrderId.SessionValue <> "" Then
				sDetailFilter = sDetailFilter & "[OrderId]=" & ew_QuotedValue(OrderId.SessionValue, EW_DATATYPE_NUMBER)
			Else
				DetailFilter = ""
				Exit Property
			End If
		End If
		DetailFilter = sDetailFilter
	End Property

	' Master filter
	Public Property Get SqlMasterFilter_Orders
		SqlMasterFilter_Orders = "[OrderId]=@OrderId@"
	End Property

	' Detail filter
	Public Property Get SqlDetailFilter_Orders
		SqlDetailFilter_Orders = "[OrderId]=@OrderId@"
	End Property

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [OrderDetails]"
	End Property

	Private Property Get TableFilter()
		TableFilter = ""
	End Property

	Public Property Get SqlWhere() ' Where
		Dim sWhere
		sWhere = ""
		Call ew_AddFilter(sWhere, TableFilter)
		SqlWhere = sWhere
	End Property

	Public Property Get SqlGroupBy() ' Group By
		SqlGroupBy = ""
	End Property

	Public Property Get SqlHaving() ' Having
		SqlHaving = ""
	End Property

	Public Property Get SqlOrderBy() ' Order By
		SqlOrderBy = ""
	End Property

	' SQL variables
	Dim CurrentFilter ' Current filter
	Dim CurrentOrder ' Current order
	Dim CurrentOrderType ' Current order type

	' Get sql
	Public Function GetSQL(where, orderby)
		GetSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, where, orderby)
	End Function

	' Table sql
	Public Property Get SQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = SessionOrderBy
		SQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Return table sql with list page filter
	Public Property Get ListSQL()
		Dim sFilter, sSort
		sFilter = SessionWhere
		Call ew_AddFilter(sFilter, CurrentFilter)
		sSort = SessionOrderBy
		ListSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Key filter for table
	Private Property Get SqlKeyFilter()
		SqlKeyFilter = "[OrderDetailsId] = @OrderDetailsId@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(OrderDetailsId.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@OrderDetailsId@", ew_AdjustSql(OrderDetailsId.CurrentValue)) ' Replace key value
		KeyFilter = sKeyFilter
	End Property

	' Return url
	Public Property Get ReturnUrl()

		' Get referer url automatically
		If Request.ServerVariables("HTTP_REFERER") <> "" Then
			If ew_ReferPage <> ew_CurrentPage And ew_ReferPage <> "login.asp" Then ' Referer not same page or login page
				Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = Request.ServerVariables("HTTP_REFERER") ' Save to Session
			End If
		End If
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "OrderDetailslist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "OrderDetailslist.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("OrderDetailsview.asp", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "OrderDetailsadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("OrderDetailsedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("OrderDetailsadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("OrderDetailsdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(OrderDetailsId.CurrentValue) Then
			sUrl = sUrl & "OrderDetailsId=" & OrderDetailsId.CurrentValue
		Else
			KeyUrl = "javascript:alert(ewLanguage.Phrase('InvalidRecord'));"
			Exit Function
		End If
		KeyUrl = sUrl
	End Function

	' Sort Url
	Public Property Get SortUrl(fld)
		If CurrentAction <> "" Or Export <> "" Or (fld.FldType = 201 Or fld.FldType = 203 Or fld.FldType = 205 Or fld.FldType = 141) Then
			SortUrl = ""
		ElseIf fld.Sortable Then
			SortUrl = ew_CurrentPage
			Dim sUrlParm
			sUrlParm = UrlParm("order=" & Server.URLEncode(fld.FldName) & "&amp;ordertype=" & fld.ReverseSort)
			SortUrl = SortUrl & "?" & sUrlParm
		Else
			SortUrl = ""
		End If
	End Property

	' Url parm
	Function UrlParm(parm)
		If UseTokenInUrl Then
			UrlParm = "t=OrderDetails"
		Else
			UrlParm = ""
		End If
		If parm <> "" Then
			If UrlParm <> "" Then UrlParm = UrlParm & "&"
			UrlParm = UrlParm & parm
		End If
	End Function

	' Get record keys from Form/QueryString/Session
	Public Function GetRecordKeys()
		Dim arKeys, arKey, cnt, i, bHasKey
		bHasKey = False

		' Check ObjForm first
		If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
			ObjForm.Index = 0
			If ObjForm.HasValue("key_m") Then
				arKeys = ObjForm.GetValue("key_m")
				If Not IsArray(arKeys) Then
					arKeys = Array(arKeys)
				End If
				bHasKey = True
			End If
		End If

		' Check Form/QueryString
		If Not bHasKey Then
			If Request.Form("key_m").Count > 0 Then
				cnt = Request.Form("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.Form("key_m")(i)
				Next
			ElseIf Request.QueryString("key_m").Count > 0 Then
				cnt = Request.QueryString("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.QueryString("key_m")(i)
				Next
			ElseIf Request.QueryString <> "" Then
				ReDim arKeys(0)
				arKeys(0) = Request.QueryString("OrderDetailsId") ' OrderDetailsId

				'GetRecordKeys = arKeys ' do not return yet, so the values will also be checked by the following code
			End If
		End If

		' Check keys
		Dim ar, key
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
						Dim skip
						skip = False
						If Not IsNumeric(key) Then skip = True
						If Not skip Then
							If IsArray(ar) Then
								ReDim Preserve ar(UBound(ar)+1)
							Else
								ReDim ar(0)
							End If
							ar(UBound(ar)) = key
						End If
			Next
		End If
		GetRecordKeys = ar
	End Function

	' Get key filter
	Public Function GetKeyFilter()
		Dim arKeys, sKeyFilter, i, key
		arKeys = GetRecordKeys()
		sKeyFilter = ""
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
				If sKeyFilter <> "" Then sKeyFilter = sKeyFilter & " OR "
				OrderDetailsId.CurrentValue = key
				sKeyFilter = sKeyFilter & "(" & KeyFilter & ")"
			Next
		End If
		GetKeyFilter = sKeyFilter
	End Function

	' Function LoadRecordCount
	' - Load record count based on filter
	Public Function LoadRecordCount(sFilter)
		Dim wrkrs
		Set wrkrs = LoadRs(sFilter)
		If Not wrkrs Is Nothing Then
			LoadRecordCount = wrkrs.RecordCount
		Else
			LoadRecordCount = 0
		End If
		Set wrkrs = Nothing
	End Function

	' Function LoadRs
	' - Load Rows based on filter
	Public Function LoadRs(sFilter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim RsRows, sSql

		' Set up filter (Sql Where Clause) and get Return Sql
		CurrentFilter = sFilter
		sSql = SQL
		Err.Clear
		Set RsRows = Server.CreateObject("ADODB.Recordset")
		RsRows.CursorLocation = EW_CURSORLOCATION
		RsRows.Open sSql, Conn, 3, 1, 1 ' adOpenStatic, adLockReadOnly, adCmdText
		If Err.Number <> 0 Then
			Err.Clear
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		ElseIf RsRows.Eof Then
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		Else
			Set LoadRs = RsRows
		End If
	End Function

	' Load row values from recordset
	Public Sub LoadListRowValues(RsRow)
		OrderDetailsId.DbValue = RsRow("OrderDetailsId")
		OrderId.DbValue = RsRow("OrderId")
		ProductId.DbValue = RsRow("ProductId")
		Quantity.DbValue = RsRow("Quantity")
		Price.DbValue = RsRow("Price")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' OrderDetailsId
		' OrderId
		' ProductId
		' Quantity
		' Price
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' OrderDetailsId

		OrderDetailsId.ViewValue = OrderDetailsId.CurrentValue
		OrderDetailsId.ViewCustomAttributes = ""

		' OrderId
		OrderId.ViewValue = OrderId.CurrentValue
		OrderId.ViewCustomAttributes = ""

		' ProductId
		If ProductId.CurrentValue & "" <> "" Then
			sFilterWrk = "[ItemId] = " & ew_AdjustSql(ProductId.CurrentValue) & ""
		sSqlWrk = "SELECT [Description] FROM [Products]"
		sWhereWrk = ""
		Call ew_AddFilter(sWhereWrk, sFilterWrk)
		If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			Set RsWrk = Conn.Execute(sSqlWrk)
			If Not RsWrk.Eof Then
				ProductId.ViewValue = RsWrk("Description")
			Else
				ProductId.ViewValue = ProductId.CurrentValue
			End If
			RsWrk.Close
			Set RsWrk = Nothing
		Else
			ProductId.ViewValue = Null
		End If
		ProductId.ViewCustomAttributes = ""

		' Quantity
		Quantity.ViewValue = Quantity.CurrentValue
		Quantity.ViewCustomAttributes = ""

		' Price
		Price.ViewValue = Price.CurrentValue
		Price.ViewCustomAttributes = ""

		' OrderDetailsId
		OrderDetailsId.LinkCustomAttributes = ""
		OrderDetailsId.HrefValue = ""
		OrderDetailsId.TooltipValue = ""

		' OrderId
		OrderId.LinkCustomAttributes = ""
		OrderId.HrefValue = ""
		OrderId.TooltipValue = ""

		' ProductId
		ProductId.LinkCustomAttributes = ""
		ProductId.HrefValue = ""
		ProductId.TooltipValue = ""

		' Quantity
		Quantity.LinkCustomAttributes = ""
		Quantity.HrefValue = ""
		Quantity.TooltipValue = ""

		' Price
		Price.LinkCustomAttributes = ""
		Price.HrefValue = ""
		Price.TooltipValue = ""

		' Call Row Rendered event
		If RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Row_Rendered()
		End If
	End Sub

	' Aggregate list row values
	Public Sub AggregateListRowValues()
	End Sub

	' Aggregate list row (for rendering)
	Sub AggregateListRow()
	End Sub

	' Export data in Xml Format
	Public Sub ExportXmlDocument(XmlDoc, HasParent, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(XmlDoc) Then
			Exit Sub
		End If
		If Not HasParent Then
			Call XmlDoc.AddRoot(TableVar)
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And RecCnt < StopRec
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				If HasParent Then
					Call XmlDoc.AddRow(TableVar, "")
				Else
					Call XmlDoc.AddRow("", "")
				End If
				If ExportPageType = "view" Then
					Call XmlDoc.AddField("OrderDetailsId", OrderDetailsId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("OrderId", OrderId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ProductId", ProductId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Quantity", Quantity.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price", Price.ExportValue(Export, ExportOriginalValue))
				Else
					Call XmlDoc.AddField("OrderDetailsId", OrderDetailsId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("OrderId", OrderId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ProductId", ProductId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Quantity", Quantity.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price", Price.ExportValue(Export, ExportOriginalValue))
				End If
			End If
			Recordset.MoveNext()
		Loop
	End Sub

	' Export data in HTML/CSV/Word/Excel/Email format
	Public Sub ExportDocument(Doc, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(Doc) Then
			Exit Sub
		End If

		' Write header
		Call Doc.ExportTableHeader()
		If Doc.Horizontal Then ' Horizontal format, write header
			Call Doc.BeginExportRow(0)
			If ExportPageType = "view" Then
				Call Doc.ExportCaption(OrderDetailsId)
				Call Doc.ExportCaption(OrderId)
				Call Doc.ExportCaption(ProductId)
				Call Doc.ExportCaption(Quantity)
				Call Doc.ExportCaption(Price)
			Else
				Call Doc.ExportCaption(OrderDetailsId)
				Call Doc.ExportCaption(OrderId)
				Call Doc.ExportCaption(ProductId)
				Call Doc.ExportCaption(Quantity)
				Call Doc.ExportCaption(Price)
			End If
			Call Doc.EndExportRow()
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And CLng(RecCnt) < CLng(StopRec)
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				Call Doc.BeginExportRow(RowCnt)
				If ExportPageType = "view" Then
					Call Doc.ExportField(OrderDetailsId)
					Call Doc.ExportField(OrderId)
					Call Doc.ExportField(ProductId)
					Call Doc.ExportField(Quantity)
					Call Doc.ExportField(Price)
				Else
					Call Doc.ExportField(OrderDetailsId)
					Call Doc.ExportField(OrderId)
					Call Doc.ExportField(ProductId)
					Call Doc.ExportField(Quantity)
					Call Doc.ExportField(Price)
				End If
				Call Doc.EndExportRow()
			End If
			Recordset.MoveNext()
		Loop
		Call Doc.ExportTableFooter()
	End Sub
	Dim CurrentAction ' Current action
	Dim LastAction ' Last action
	Dim CurrentMode ' Current mode
	Dim UpdateConflict ' Update conflict
	Dim EventName ' Event name
	Dim EventCancelled ' Event cancelled
	Dim CancelMessage ' Cancel message
	Dim AllowAddDeleteRow ' Allow add/delete row
	Dim DetailAdd ' Allow detail add
	Dim DetailEdit ' Allow detail edit
	Dim GridAddRowCount ' Grid add row count

	' Check current action
	' - Add
	Public Function IsAdd()
		IsAdd = (CurrentAction = "add")
	End Function

	' - Copy
	Public Function IsCopy()
		IsCopy = (CurrentAction = "copy" Or CurrentAction = "C")
	End Function

	' - Edit
	Public Function IsEdit()
		IsEdit = (CurrentAction = "edit")
	End Function

	' - Delete
	Public Function IsDelete()
		IsDelete = (CurrentAction = "D")
	End Function

	' - Confirm
	Public Function IsConfirm()
		IsConfirm = (CurrentAction = "F")
	End Function

	' - Overwrite
	Public Function IsOverwrite()
		IsOverwrite = (CurrentAction = "overwrite")
	End Function

	' - Cancel
	Public Function IsCancel()
		IsCancel = (CurrentAction = "cancel")
	End Function

	' - Grid add
	Public Function IsGridAdd()
		IsGridAdd = (CurrentAction = "gridadd")
	End Function

	' - Grid edit
	Public Function IsGridEdit()
		IsGridEdit = (CurrentAction = "gridedit")
	End Function

	' - Insert
	Public Function IsInsert()
		IsInsert = (CurrentAction = "insert" Or CurrentAction = "A")
	End Function

	' - Update
	Public Function IsUpdate()
		IsUpdate = (CurrentAction = "update" Or CurrentAction = "U")
	End Function

	' - Grid update
	Public Function IsGridUpdate()
		IsGridUpdate = (CurrentAction = "gridupdate")
	End Function

	' - Grid insert
	Public Function IsGridInsert()
		IsGridInsert = (CurrentAction = "gridinsert")
	End Function

	' - Grid overwrite
	Public Function IsGridOverwrite()
		IsGridOverwrite = (CurrentAction = "gridoverwrite")
	End Function

	' Check last action
	' - Cancelled
	Public Function IsCancelled()
		IsCancelled = (LastAction = "cancel" And CurrentAction = "")
	End Function

	' - Inline inserted
	Public Function IsInlineInserted()
		IsInlineInserted = (LastAction = "insert" And CurrentAction = "")
	End Function

	' - Inline updated
	Public Function IsInlineUpdated()
		IsInlineUpdated = (LastAction = "update" And CurrentAction = "")
	End Function

	' - Grid updated
	Public Function IsGridUpdated()
		IsGridUpdated = (LastAction = "gridupdate" And CurrentAction = "")
	End Function

	' - Grid inserted
	Public Function IsGridInserted()
		IsGridInserted = (LastAction = "gridinsert" And CurrentAction = "")
	End Function

	' Row Type
	Private m_RowType

	Public Property Get RowType()
		RowType = m_RowType
	End Property

	Public Property Let RowType(v)
		m_RowType = v
	End Property
	Dim CssClass ' Css class
	Dim CssStyle' Css style

'	Dim RowClientEvents ' Row client events
	Dim RowAttrs ' Row attributes

	' Row Styles
	Public Property Get RowStyles()
		Dim sAtt, Value
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If RowAttrs.Exists("style") Then
			Value = RowAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If RowAttrs.Exists("class") Then
			Value = RowAttrs.Item("class")
			If Trim(Value) <> "" Then
				sClass = sClass & " " & Value
			End If
		End If
		If Trim(sStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(sStyle) & """" 
		End If
		If Trim(sClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(sClass) & """" 
		End If
		RowStyles = sAtt
	End Property

	' Row Attribute
	Public Property Get RowAttributes()
		Dim sAtt, Attr, Value, i
		sAtt = RowStyles
		If m_Export = "" Then

'			If Trim(RowClientEvents) <> "" Then
'				sAtt = sAtt & " " & Trim(RowClientEvents)
'			End If

			For i = 0 to UBound(RowAttrs.Attributes)
				Attr = RowAttrs.Attributes(i)(0)
				Value = RowAttrs.Attributes(i)(1)
				If Attr <> "style" And Attr <> "class" And Attr <> "" And Value <> "" Then
					sAtt = sAtt & " " & Attr & "=""" & Value & """"
				End If
			Next
		End If
		RowAttributes = sAtt
	End Property

	' Export
	Private m_Export

	Public Property Get Export()
		Export = m_Export
	End Property

	Public Property Let Export(v)
		m_Export = v
	End Property

	' Export Original Value
	Dim ExportOriginalValue

	' Export All
	Dim ExportAll

	' Send Email
	Dim SendEmail

	' Custom Inner Html
	Dim TableCustomInnerHtml

	' ----------------
	'  Field objects
	' ----------------
	' Field OrderDetailsId
	Private m_OrderDetailsId

	Public Property Get OrderDetailsId()
		If Not IsObject(m_OrderDetailsId) Then
			Set m_OrderDetailsId = NewFldObj("OrderDetails", "OrderDetails", "x_OrderDetailsId", "OrderDetailsId", "[OrderDetailsId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_OrderDetailsId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set OrderDetailsId = m_OrderDetailsId
	End Property

	' Field OrderId
	Private m_OrderId

	Public Property Get OrderId()
		If Not IsObject(m_OrderId) Then
			Set m_OrderId = NewFldObj("OrderDetails", "OrderDetails", "x_OrderId", "OrderId", "[OrderId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_OrderId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set OrderId = m_OrderId
	End Property

	' Field ProductId
	Private m_ProductId

	Public Property Get ProductId()
		If Not IsObject(m_ProductId) Then
			Set m_ProductId = NewFldObj("OrderDetails", "OrderDetails", "x_ProductId", "ProductId", "[ProductId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_ProductId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set ProductId = m_ProductId
	End Property

	' Field Quantity
	Private m_Quantity

	Public Property Get Quantity()
		If Not IsObject(m_Quantity) Then
			Set m_Quantity = NewFldObj("OrderDetails", "OrderDetails", "x_Quantity", "Quantity", "[Quantity]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_Quantity.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Quantity = m_Quantity
	End Property

	' Field Price
	Private m_Price

	Public Property Get Price()
		If Not IsObject(m_Price) Then
			Set m_Price = NewFldObj("OrderDetails", "OrderDetails", "x_Price", "Price", "[Price]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_Price.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Price = m_Price
	End Property
	Dim Fields ' Fields

	' Create new field object
	Private Function NewFldObj(TblVar, TblName, FldVar, FldName, FldExpression, FldType, FldDtFormat, FldVirtualExp, FldVirtual, FldForceSelect, FldViewTag)
		Dim fld
		Set fld = New cField
		fld.TblVar = TblVar
		fld.TblName = TblName
		fld.FldVar = FldVar
		fld.FldName = FldName
		fld.FldExpression = FldExpression
		fld.FldType = FldType
		fld.FldDataType = ew_FieldDataType(FldType)
		fld.FldDateTimeFormat = FldDtFormat
		fld.FldVirtualExpression = FldVirtualExp
		fld.FldIsVirtual = FldVirtual
		fld.FldForceSelection = FldForceSelect
		fld.FldViewTag = FldViewTag
		Set NewFldObj = fld
	End Function

	' Table level events
	' Recordset Selecting event
	Sub Recordset_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Recordset Selected event
	Sub Recordset_Selected(rs)

		'Response.Write "Recordset Selected"
	End Sub

	' Recordset Search Validated event
	Sub Recordset_SearchValidated()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
	End Sub

	' Recordset Searching event
	Sub Recordset_Searching(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row_Selecting event
	Sub Row_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Selected event
	Sub Row_Selected(rs)

		'Response.Write "Row Selected"
	End Sub

	' Row Inserting event
	Function Row_Inserting(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Inserting = True
	End Function

	' Row Inserted event
	Sub Row_Inserted(rsold, rsnew)

		' Response.Write "Row Inserted"
	End Sub

	' Row Updating event
	Function Row_Updating(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Updating = True
	End Function

	' Row Updated event
	Sub Row_Updated(rsold, rsnew)

		' Response.Write "Row Updated"
	End Sub

	' Row Update Conflict event
	Function Row_UpdateConflict(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To ignore conflict, set return value to False

		Row_UpdateConflict = True
	End Function

	' Row Deleting event
	Function Row_Deleting(rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Deleting = True
	End Function

	' Row Deleted event
	Sub Row_Deleted(rs)

		' Response.Write "Row Deleted"
	End Sub

	' Email Sending event
	Function Email_Sending(Email, Args)

		'Response.Write Email.AsString
		'Response.Write "Keys of Args: " & Join(Args.Keys, ", ")
		'Response.End

		Email_Sending = True
	End Function

	' Row Rendering event
	Sub Row_Rendering()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Rendered event
	Sub Row_Rendered()

		' To view properties of field class, use:
		' Response.Write <FieldName>.AsString() 

	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_OrderDetailsId) Then Set m_OrderDetailsId = Nothing
		If IsObject(m_OrderId) Then Set m_OrderId = Nothing
		If IsObject(m_ProductId) Then Set m_ProductId = Nothing
		If IsObject(m_Quantity) Then Set m_Quantity = Nothing
		If IsObject(m_Price) Then Set m_Price = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
