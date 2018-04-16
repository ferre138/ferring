<%

' ASPMaker configuration for Table DiscountTypes
Dim DiscountTypes

' Define table class
Class cDiscountTypes

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
		Call ew_SetArObj(Fields, "DiscountTypeId", DiscountTypeId)
		Call ew_SetArObj(Fields, "DiscountType", DiscountType)
		Call ew_SetArObj(Fields, "DiscountTitle", DiscountTitle)
		Call ew_SetArObj(Fields, "freeShipping", freeShipping)
		Call ew_SetArObj(Fields, "FreePerQty", FreePerQty)
		Call ew_SetArObj(Fields, "SpecialPrice", SpecialPrice)
		Call ew_SetArObj(Fields, "fDiscountTitle", fDiscountTitle)
		Call ew_SetArObj(Fields, "StartDate", StartDate)
		Call ew_SetArObj(Fields, "EndDate", EndDate)
		Call ew_SetArObj(Fields, "DiscountPerc", DiscountPerc)
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
		TableVar = "DiscountTypes"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "DiscountTypes"
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
		HighlightName = "DiscountTypes_Highlight"
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

	' Current detail table name
	Public Property Get CurrentDetailTable()
		CurrentDetailTable = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_DETAIL_TABLE)
	End Property

	Public Property Let CurrentDetailTable(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_DETAIL_TABLE) = v
	End Property

	' Get detail url
	Public Property Get DetailUrl()

		' Detail url
		Dim sDetailUrl
		sDetailUrl = ""
		If CurrentDetailTable = "Discountcodes" Then
			sDetailUrl = Discountcodes.ListUrl & "?showmaster=" & TableVar
			sDetailUrl = sDetailUrl & "&DiscountTypeId=" & DiscountTypeId.CurrentValue
		End If
		DetailUrl = sDetailUrl
	End Property

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [DiscountTypes]"
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
		SqlKeyFilter = "[DiscountTypeId] = @DiscountTypeId@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(DiscountTypeId.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@DiscountTypeId@", ew_AdjustSql(DiscountTypeId.CurrentValue)) ' Replace key value
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
			ReturnUrl = "DiscountTypeslist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "DiscountTypeslist.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("DiscountTypesview.asp", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "DiscountTypesadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		If parm <> "" Then
			EditUrl = KeyUrl("DiscountTypesedit.asp", UrlParm(parm))
		Else
			EditUrl = KeyUrl("DiscountTypesedit.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		If parm <> "" Then
			CopyUrl = KeyUrl("DiscountTypesadd.asp", UrlParm(parm))
		Else
			CopyUrl = KeyUrl("DiscountTypesadd.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("DiscountTypesdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(DiscountTypeId.CurrentValue) Then
			sUrl = sUrl & "DiscountTypeId=" & DiscountTypeId.CurrentValue
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
			UrlParm = "t=DiscountTypes"
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
				arKeys(0) = Request.QueryString("DiscountTypeId") ' DiscountTypeId

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
				DiscountTypeId.CurrentValue = key
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
		DiscountTypeId.DbValue = RsRow("DiscountTypeId")
		DiscountType.DbValue = RsRow("DiscountType")
		DiscountTitle.DbValue = RsRow("DiscountTitle")
		freeShipping.DbValue = ew_IIf(RsRow("freeShipping"), "1", "0")
		FreePerQty.DbValue = RsRow("FreePerQty")
		SpecialPrice.DbValue = RsRow("SpecialPrice")
		fDiscountTitle.DbValue = RsRow("fDiscountTitle")
		StartDate.DbValue = RsRow("StartDate")
		EndDate.DbValue = RsRow("EndDate")
		DiscountPerc.DbValue = RsRow("DiscountPerc")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' DiscountTypeId
		' DiscountType
		' DiscountTitle
		' freeShipping
		' FreePerQty
		' SpecialPrice
		' fDiscountTitle
		' StartDate
		' EndDate
		' DiscountPerc
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' DiscountTypeId

		DiscountTypeId.ViewValue = DiscountTypeId.CurrentValue
		DiscountTypeId.ViewCustomAttributes = ""

		' DiscountType
		DiscountType.ViewValue = DiscountType.CurrentValue
		DiscountType.ViewCustomAttributes = ""

		' DiscountTitle
		DiscountTitle.ViewValue = DiscountTitle.CurrentValue
		DiscountTitle.ViewCustomAttributes = ""

		' freeShipping
		If ew_ConvertToBool(freeShipping.CurrentValue) Then
			freeShipping.ViewValue = ew_IIf(freeShipping.FldTagCaption(1) <> "", freeShipping.FldTagCaption(1), "Yes")
		Else
			freeShipping.ViewValue = ew_IIf(freeShipping.FldTagCaption(2) <> "", freeShipping.FldTagCaption(2), "No")
		End If
		freeShipping.ViewCustomAttributes = ""

		' FreePerQty
		FreePerQty.ViewValue = FreePerQty.CurrentValue
		FreePerQty.ViewCustomAttributes = ""

		' SpecialPrice
		SpecialPrice.ViewValue = SpecialPrice.CurrentValue
		SpecialPrice.ViewCustomAttributes = ""

		' fDiscountTitle
		fDiscountTitle.ViewValue = fDiscountTitle.CurrentValue
		fDiscountTitle.ViewCustomAttributes = ""

		' StartDate
		StartDate.ViewValue = StartDate.CurrentValue
		StartDate.ViewCustomAttributes = ""

		' EndDate
		EndDate.ViewValue = EndDate.CurrentValue
		EndDate.ViewCustomAttributes = ""

		' DiscountPerc
		DiscountPerc.ViewValue = DiscountPerc.CurrentValue
		DiscountPerc.ViewCustomAttributes = ""

		' DiscountTypeId
		DiscountTypeId.LinkCustomAttributes = ""
		DiscountTypeId.HrefValue = ""
		DiscountTypeId.TooltipValue = ""

		' DiscountType
		DiscountType.LinkCustomAttributes = ""
		DiscountType.HrefValue = ""
		DiscountType.TooltipValue = ""

		' DiscountTitle
		DiscountTitle.LinkCustomAttributes = ""
		DiscountTitle.HrefValue = ""
		DiscountTitle.TooltipValue = ""

		' freeShipping
		freeShipping.LinkCustomAttributes = ""
		freeShipping.HrefValue = ""
		freeShipping.TooltipValue = ""

		' FreePerQty
		FreePerQty.LinkCustomAttributes = ""
		FreePerQty.HrefValue = ""
		FreePerQty.TooltipValue = ""

		' SpecialPrice
		SpecialPrice.LinkCustomAttributes = ""
		SpecialPrice.HrefValue = ""
		SpecialPrice.TooltipValue = ""

		' fDiscountTitle
		fDiscountTitle.LinkCustomAttributes = ""
		fDiscountTitle.HrefValue = ""
		fDiscountTitle.TooltipValue = ""

		' StartDate
		StartDate.LinkCustomAttributes = ""
		StartDate.HrefValue = ""
		StartDate.TooltipValue = ""

		' EndDate
		EndDate.LinkCustomAttributes = ""
		EndDate.HrefValue = ""
		EndDate.TooltipValue = ""

		' DiscountPerc
		DiscountPerc.LinkCustomAttributes = ""
		DiscountPerc.HrefValue = ""
		DiscountPerc.TooltipValue = ""

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
					Call XmlDoc.AddField("DiscountTypeId", DiscountTypeId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountType", DiscountType.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountTitle", DiscountTitle.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("freeShipping", freeShipping.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("FreePerQty", FreePerQty.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("SpecialPrice", SpecialPrice.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fDiscountTitle", fDiscountTitle.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("StartDate", StartDate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("EndDate", EndDate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountPerc", DiscountPerc.ExportValue(Export, ExportOriginalValue))
				Else
					Call XmlDoc.AddField("DiscountTypeId", DiscountTypeId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountType", DiscountType.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountTitle", DiscountTitle.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("freeShipping", freeShipping.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("FreePerQty", FreePerQty.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("SpecialPrice", SpecialPrice.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fDiscountTitle", fDiscountTitle.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("StartDate", StartDate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("EndDate", EndDate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("DiscountPerc", DiscountPerc.ExportValue(Export, ExportOriginalValue))
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
				Call Doc.ExportCaption(DiscountTypeId)
				Call Doc.ExportCaption(DiscountType)
				Call Doc.ExportCaption(DiscountTitle)
				Call Doc.ExportCaption(freeShipping)
				Call Doc.ExportCaption(FreePerQty)
				Call Doc.ExportCaption(SpecialPrice)
				Call Doc.ExportCaption(fDiscountTitle)
				Call Doc.ExportCaption(StartDate)
				Call Doc.ExportCaption(EndDate)
				Call Doc.ExportCaption(DiscountPerc)
			Else
				Call Doc.ExportCaption(DiscountTypeId)
				Call Doc.ExportCaption(DiscountType)
				Call Doc.ExportCaption(DiscountTitle)
				Call Doc.ExportCaption(freeShipping)
				Call Doc.ExportCaption(FreePerQty)
				Call Doc.ExportCaption(SpecialPrice)
				Call Doc.ExportCaption(fDiscountTitle)
				Call Doc.ExportCaption(StartDate)
				Call Doc.ExportCaption(EndDate)
				Call Doc.ExportCaption(DiscountPerc)
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
					Call Doc.ExportField(DiscountTypeId)
					Call Doc.ExportField(DiscountType)
					Call Doc.ExportField(DiscountTitle)
					Call Doc.ExportField(freeShipping)
					Call Doc.ExportField(FreePerQty)
					Call Doc.ExportField(SpecialPrice)
					Call Doc.ExportField(fDiscountTitle)
					Call Doc.ExportField(StartDate)
					Call Doc.ExportField(EndDate)
					Call Doc.ExportField(DiscountPerc)
				Else
					Call Doc.ExportField(DiscountTypeId)
					Call Doc.ExportField(DiscountType)
					Call Doc.ExportField(DiscountTitle)
					Call Doc.ExportField(freeShipping)
					Call Doc.ExportField(FreePerQty)
					Call Doc.ExportField(SpecialPrice)
					Call Doc.ExportField(fDiscountTitle)
					Call Doc.ExportField(StartDate)
					Call Doc.ExportField(EndDate)
					Call Doc.ExportField(DiscountPerc)
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
	' Field DiscountTypeId
	Private m_DiscountTypeId

	Public Property Get DiscountTypeId()
		If Not IsObject(m_DiscountTypeId) Then
			Set m_DiscountTypeId = NewFldObj("DiscountTypes", "DiscountTypes", "x_DiscountTypeId", "DiscountTypeId", "[DiscountTypeId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_DiscountTypeId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set DiscountTypeId = m_DiscountTypeId
	End Property

	' Field DiscountType
	Private m_DiscountType

	Public Property Get DiscountType()
		If Not IsObject(m_DiscountType) Then
			Set m_DiscountType = NewFldObj("DiscountTypes", "DiscountTypes", "x_DiscountType", "DiscountType", "[DiscountType]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set DiscountType = m_DiscountType
	End Property

	' Field DiscountTitle
	Private m_DiscountTitle

	Public Property Get DiscountTitle()
		If Not IsObject(m_DiscountTitle) Then
			Set m_DiscountTitle = NewFldObj("DiscountTypes", "DiscountTypes", "x_DiscountTitle", "DiscountTitle", "[DiscountTitle]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set DiscountTitle = m_DiscountTitle
	End Property

	' Field freeShipping
	Private m_freeShipping

	Public Property Get freeShipping()
		If Not IsObject(m_freeShipping) Then
			Set m_freeShipping = NewFldObj("DiscountTypes", "DiscountTypes", "x_freeShipping", "freeShipping", "[freeShipping]", 11, 8, "", False, False, "FORMATTED TEXT")
			m_freeShipping.FldDataType = EW_DATATYPE_BOOLEAN
		End If
		Set freeShipping = m_freeShipping
	End Property

	' Field FreePerQty
	Private m_FreePerQty

	Public Property Get FreePerQty()
		If Not IsObject(m_FreePerQty) Then
			Set m_FreePerQty = NewFldObj("DiscountTypes", "DiscountTypes", "x_FreePerQty", "FreePerQty", "[FreePerQty]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_FreePerQty.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set FreePerQty = m_FreePerQty
	End Property

	' Field SpecialPrice
	Private m_SpecialPrice

	Public Property Get SpecialPrice()
		If Not IsObject(m_SpecialPrice) Then
			Set m_SpecialPrice = NewFldObj("DiscountTypes", "DiscountTypes", "x_SpecialPrice", "SpecialPrice", "[SpecialPrice]", 5, 8, "", False, False, "FORMATTED TEXT")
			m_SpecialPrice.FldDefaultErrMsg = Language.Phrase("IncorrectFloat")
		End If
		Set SpecialPrice = m_SpecialPrice
	End Property

	' Field fDiscountTitle
	Private m_fDiscountTitle

	Public Property Get fDiscountTitle()
		If Not IsObject(m_fDiscountTitle) Then
			Set m_fDiscountTitle = NewFldObj("DiscountTypes", "DiscountTypes", "x_fDiscountTitle", "fDiscountTitle", "[fDiscountTitle]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set fDiscountTitle = m_fDiscountTitle
	End Property

	' Field StartDate
	Private m_StartDate

	Public Property Get StartDate()
		If Not IsObject(m_StartDate) Then
			Set m_StartDate = NewFldObj("DiscountTypes", "DiscountTypes", "x_StartDate", "StartDate", "[StartDate]", 135, 8, "", False, False, "FORMATTED TEXT")
			m_StartDate.FldDefaultErrMsg = Replace(Language.Phrase("IncorrectDateMDY"), "%s", "/")
		End If
		Set StartDate = m_StartDate
	End Property

	' Field EndDate
	Private m_EndDate

	Public Property Get EndDate()
		If Not IsObject(m_EndDate) Then
			Set m_EndDate = NewFldObj("DiscountTypes", "DiscountTypes", "x_EndDate", "EndDate", "[EndDate]", 135, 8, "", False, False, "FORMATTED TEXT")
			m_EndDate.FldDefaultErrMsg = Replace(Language.Phrase("IncorrectDateMDY"), "%s", "/")
		End If
		Set EndDate = m_EndDate
	End Property

	' Field DiscountPerc
	Private m_DiscountPerc

	Public Property Get DiscountPerc()
		If Not IsObject(m_DiscountPerc) Then
			Set m_DiscountPerc = NewFldObj("DiscountTypes", "DiscountTypes", "x_DiscountPerc", "DiscountPerc", "[DiscountPerc]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_DiscountPerc.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set DiscountPerc = m_DiscountPerc
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
		If IsObject(m_DiscountTypeId) Then Set m_DiscountTypeId = Nothing
		If IsObject(m_DiscountType) Then Set m_DiscountType = Nothing
		If IsObject(m_DiscountTitle) Then Set m_DiscountTitle = Nothing
		If IsObject(m_freeShipping) Then Set m_freeShipping = Nothing
		If IsObject(m_FreePerQty) Then Set m_FreePerQty = Nothing
		If IsObject(m_SpecialPrice) Then Set m_SpecialPrice = Nothing
		If IsObject(m_fDiscountTitle) Then Set m_fDiscountTitle = Nothing
		If IsObject(m_StartDate) Then Set m_StartDate = Nothing
		If IsObject(m_EndDate) Then Set m_EndDate = Nothing
		If IsObject(m_DiscountPerc) Then Set m_DiscountPerc = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
