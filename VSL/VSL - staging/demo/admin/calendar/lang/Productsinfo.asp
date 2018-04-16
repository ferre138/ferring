<%

' ASPMaker configuration for Table Products
Dim Products

' Define table class
Class cProducts

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
		Call ew_SetArObj(Fields, "ItemId", ItemId)
		Call ew_SetArObj(Fields, "Description", Description)
		Call ew_SetArObj(Fields, "Price", Price)
		Call ew_SetArObj(Fields, "Active", Active)
		Call ew_SetArObj(Fields, "Image", Image)
		Call ew_SetArObj(Fields, "Sizes", Sizes)
		Call ew_SetArObj(Fields, "Image_Thumb", Image_Thumb)
		Call ew_SetArObj(Fields, "ProductName", ProductName)
		Call ew_SetArObj(Fields, "ItemNo", ItemNo)
		Call ew_SetArObj(Fields, "UPC", UPC)
		Call ew_SetArObj(Fields, "Price_rebate", Price_rebate)
		Call ew_SetArObj(Fields, "fDescription", fDescription)
		Call ew_SetArObj(Fields, "fImage", fImage)
		Call ew_SetArObj(Fields, "fSizes", fSizes)
		Call ew_SetArObj(Fields, "fImage_Thumb", fImage_Thumb)
		Call ew_SetArObj(Fields, "fProductName", fProductName)
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
		TableVar = "Products"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Products"
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
		HighlightName = "Products_Highlight"
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

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [Products]"
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
		SqlKeyFilter = "[ItemId] = @ItemId@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(ItemId.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@ItemId@", ew_AdjustSql(ItemId.CurrentValue)) ' Replace key value
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
			ReturnUrl = "Productslist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "Productslist.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("Productsview.asp", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "Productsadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("Productsedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("Productsadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("Productsdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(ItemId.CurrentValue) Then
			sUrl = sUrl & "ItemId=" & ItemId.CurrentValue
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
			UrlParm = "t=Products"
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
				arKeys(0) = Request.QueryString("ItemId") ' ItemId

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
				ItemId.CurrentValue = key
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
		ItemId.DbValue = RsRow("ItemId")
		Description.DbValue = RsRow("Description")
		Price.DbValue = RsRow("Price")
		Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Image.Upload.DbValue = RsRow("Image")
		Sizes.DbValue = RsRow("Sizes")
		Image_Thumb.Upload.DbValue = RsRow("Image_Thumb")
		ProductName.DbValue = RsRow("ProductName")
		ItemNo.DbValue = RsRow("ItemNo")
		UPC.DbValue = RsRow("UPC")
		Price_rebate.DbValue = RsRow("Price_rebate")
		fDescription.DbValue = RsRow("fDescription")
		fImage.Upload.DbValue = RsRow("fImage")
		fSizes.DbValue = RsRow("fSizes")
		fImage_Thumb.Upload.DbValue = RsRow("fImage_Thumb")
		fProductName.DbValue = RsRow("fProductName")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' ItemId
		' Description
		' Price
		' Active
		' Image
		' Sizes
		' Image_Thumb
		' ProductName
		' ItemNo
		' UPC
		' Price_rebate
		' fDescription
		' fImage
		' fSizes
		' fImage_Thumb
		' fProductName
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' ItemId

		ItemId.ViewValue = ItemId.CurrentValue
		ItemId.ViewCustomAttributes = ""

		' Description
		Description.ViewValue = Description.CurrentValue
		Description.ViewCustomAttributes = ""

		' Price
		Price.ViewValue = Price.CurrentValue
		Price.ViewCustomAttributes = ""

		' Active
		If ew_ConvertToBool(Active.CurrentValue) Then
			Active.ViewValue = ew_IIf(Active.FldTagCaption(1) <> "", Active.FldTagCaption(1), "Yes")
		Else
			Active.ViewValue = ew_IIf(Active.FldTagCaption(2) <> "", Active.FldTagCaption(2), "No")
		End If
		Active.ViewCustomAttributes = ""

		' Image
		If Not ew_Empty(Image.Upload.DbValue) Then
			Image.ViewValue = Image.Upload.DbValue
			Image.ImageAlt = Image.FldAlt
		Else
			Image.ViewValue = ""
		End If
		Image.ViewCustomAttributes = ""

		' Sizes
		Sizes.ViewValue = Sizes.CurrentValue
		Sizes.ViewCustomAttributes = ""

		' Image_Thumb
		If Not ew_Empty(Image_Thumb.Upload.DbValue) Then
			Image_Thumb.ViewValue = Image_Thumb.Upload.DbValue
			Image_Thumb.ImageAlt = Image_Thumb.FldAlt
		Else
			Image_Thumb.ViewValue = ""
		End If
		Image_Thumb.ViewCustomAttributes = ""

		' ProductName
		ProductName.ViewValue = ProductName.CurrentValue
		ProductName.ViewCustomAttributes = ""

		' ItemNo
		ItemNo.ViewValue = ItemNo.CurrentValue
		ItemNo.ViewCustomAttributes = ""

		' UPC
		UPC.ViewValue = UPC.CurrentValue
		UPC.ViewCustomAttributes = ""

		' Price_rebate
		Price_rebate.ViewValue = Price_rebate.CurrentValue
		Price_rebate.ViewCustomAttributes = ""

		' fDescription
		fDescription.ViewValue = fDescription.CurrentValue
		fDescription.ViewCustomAttributes = ""

		' fImage
		If Not ew_Empty(fImage.Upload.DbValue) Then
			fImage.ViewValue = fImage.Upload.DbValue
			fImage.ImageAlt = fImage.FldAlt
		Else
			fImage.ViewValue = ""
		End If
		fImage.ViewCustomAttributes = ""

		' fSizes
		fSizes.ViewValue = fSizes.CurrentValue
		fSizes.ViewCustomAttributes = ""

		' fImage_Thumb
		If Not ew_Empty(fImage_Thumb.Upload.DbValue) Then
			fImage_Thumb.ViewValue = fImage_Thumb.Upload.DbValue
			fImage_Thumb.ImageAlt = fImage_Thumb.FldAlt
		Else
			fImage_Thumb.ViewValue = ""
		End If
		fImage_Thumb.ViewCustomAttributes = ""

		' fProductName
		fProductName.ViewValue = fProductName.CurrentValue
		fProductName.ViewCustomAttributes = ""

		' ItemId
		ItemId.LinkCustomAttributes = ""
		ItemId.HrefValue = ""
		ItemId.TooltipValue = ""

		' Description
		Description.LinkCustomAttributes = ""
		Description.HrefValue = ""
		Description.TooltipValue = ""

		' Price
		Price.LinkCustomAttributes = ""
		Price.HrefValue = ""
		Price.TooltipValue = ""

		' Active
		Active.LinkCustomAttributes = ""
		Active.HrefValue = ""
		Active.TooltipValue = ""

		' Image
		Image.LinkCustomAttributes = ""
		Image.HrefValue = ""
		Image.TooltipValue = ""

		' Sizes
		Sizes.LinkCustomAttributes = ""
		Sizes.HrefValue = ""
		Sizes.TooltipValue = ""

		' Image_Thumb
		Image_Thumb.LinkCustomAttributes = ""
		Image_Thumb.HrefValue = ""
		Image_Thumb.TooltipValue = ""

		' ProductName
		ProductName.LinkCustomAttributes = ""
		ProductName.HrefValue = ""
		ProductName.TooltipValue = ""

		' ItemNo
		ItemNo.LinkCustomAttributes = ""
		ItemNo.HrefValue = ""
		ItemNo.TooltipValue = ""

		' UPC
		UPC.LinkCustomAttributes = ""
		UPC.HrefValue = ""
		UPC.TooltipValue = ""

		' Price_rebate
		Price_rebate.LinkCustomAttributes = ""
		Price_rebate.HrefValue = ""
		Price_rebate.TooltipValue = ""

		' fDescription
		fDescription.LinkCustomAttributes = ""
		fDescription.HrefValue = ""
		fDescription.TooltipValue = ""

		' fImage
		fImage.LinkCustomAttributes = ""
		fImage.HrefValue = ""
		fImage.TooltipValue = ""

		' fSizes
		fSizes.LinkCustomAttributes = ""
		fSizes.HrefValue = ""
		fSizes.TooltipValue = ""

		' fImage_Thumb
		fImage_Thumb.LinkCustomAttributes = ""
		fImage_Thumb.HrefValue = ""
		fImage_Thumb.TooltipValue = ""

		' fProductName
		fProductName.LinkCustomAttributes = ""
		fProductName.HrefValue = ""
		fProductName.TooltipValue = ""

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
					Call XmlDoc.AddField("ItemId", ItemId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Description", Description.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price", Price.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Active", Active.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Image", Image.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Sizes", Sizes.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Image_Thumb", Image_Thumb.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ProductName", ProductName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ItemNo", ItemNo.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("UPC", UPC.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price_rebate", Price_rebate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fDescription", fDescription.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fImage", fImage.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fSizes", fSizes.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fImage_Thumb", fImage_Thumb.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fProductName", fProductName.ExportValue(Export, ExportOriginalValue))
				Else
					Call XmlDoc.AddField("ItemId", ItemId.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Description", Description.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price", Price.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Active", Active.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Image", Image.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Sizes", Sizes.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Image_Thumb", Image_Thumb.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ProductName", ProductName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("ItemNo", ItemNo.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("UPC", UPC.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Price_rebate", Price_rebate.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fDescription", fDescription.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fImage", fImage.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fSizes", fSizes.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fImage_Thumb", fImage_Thumb.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("fProductName", fProductName.ExportValue(Export, ExportOriginalValue))
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
				Call Doc.ExportCaption(ItemId)
				Call Doc.ExportCaption(Description)
				Call Doc.ExportCaption(Price)
				Call Doc.ExportCaption(Active)
				Call Doc.ExportCaption(Image)
				Call Doc.ExportCaption(Sizes)
				Call Doc.ExportCaption(Image_Thumb)
				Call Doc.ExportCaption(ProductName)
				Call Doc.ExportCaption(ItemNo)
				Call Doc.ExportCaption(UPC)
				Call Doc.ExportCaption(Price_rebate)
				Call Doc.ExportCaption(fDescription)
				Call Doc.ExportCaption(fImage)
				Call Doc.ExportCaption(fSizes)
				Call Doc.ExportCaption(fImage_Thumb)
				Call Doc.ExportCaption(fProductName)
			Else
				Call Doc.ExportCaption(ItemId)
				Call Doc.ExportCaption(Description)
				Call Doc.ExportCaption(Price)
				Call Doc.ExportCaption(Active)
				Call Doc.ExportCaption(Image)
				Call Doc.ExportCaption(Sizes)
				Call Doc.ExportCaption(Image_Thumb)
				Call Doc.ExportCaption(ProductName)
				Call Doc.ExportCaption(ItemNo)
				Call Doc.ExportCaption(UPC)
				Call Doc.ExportCaption(Price_rebate)
				Call Doc.ExportCaption(fDescription)
				Call Doc.ExportCaption(fImage)
				Call Doc.ExportCaption(fSizes)
				Call Doc.ExportCaption(fImage_Thumb)
				Call Doc.ExportCaption(fProductName)
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
					Call Doc.ExportField(ItemId)
					Call Doc.ExportField(Description)
					Call Doc.ExportField(Price)
					Call Doc.ExportField(Active)
					Call Doc.ExportField(Image)
					Call Doc.ExportField(Sizes)
					Call Doc.ExportField(Image_Thumb)
					Call Doc.ExportField(ProductName)
					Call Doc.ExportField(ItemNo)
					Call Doc.ExportField(UPC)
					Call Doc.ExportField(Price_rebate)
					Call Doc.ExportField(fDescription)
					Call Doc.ExportField(fImage)
					Call Doc.ExportField(fSizes)
					Call Doc.ExportField(fImage_Thumb)
					Call Doc.ExportField(fProductName)
				Else
					Call Doc.ExportField(ItemId)
					Call Doc.ExportField(Description)
					Call Doc.ExportField(Price)
					Call Doc.ExportField(Active)
					Call Doc.ExportField(Image)
					Call Doc.ExportField(Sizes)
					Call Doc.ExportField(Image_Thumb)
					Call Doc.ExportField(ProductName)
					Call Doc.ExportField(ItemNo)
					Call Doc.ExportField(UPC)
					Call Doc.ExportField(Price_rebate)
					Call Doc.ExportField(fDescription)
					Call Doc.ExportField(fImage)
					Call Doc.ExportField(fSizes)
					Call Doc.ExportField(fImage_Thumb)
					Call Doc.ExportField(fProductName)
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
	' Field ItemId
	Private m_ItemId

	Public Property Get ItemId()
		If Not IsObject(m_ItemId) Then
			Set m_ItemId = NewFldObj("Products", "Products", "x_ItemId", "ItemId", "[ItemId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_ItemId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set ItemId = m_ItemId
	End Property

	' Field Description
	Private m_Description

	Public Property Get Description()
		If Not IsObject(m_Description) Then
			Set m_Description = NewFldObj("Products", "Products", "x_Description", "Description", "[Description]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Description = m_Description
	End Property

	' Field Price
	Private m_Price

	Public Property Get Price()
		If Not IsObject(m_Price) Then
			Set m_Price = NewFldObj("Products", "Products", "x_Price", "Price", "[Price]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Price = m_Price
	End Property

	' Field Active
	Private m_Active

	Public Property Get Active()
		If Not IsObject(m_Active) Then
			Set m_Active = NewFldObj("Products", "Products", "x_Active", "Active", "[Active]", 11, 8, "", False, False, "FORMATTED TEXT")
			m_Active.FldDataType = EW_DATATYPE_BOOLEAN
		End If
		Set Active = m_Active
	End Property

	' Field Image
	Private m_Image

	Public Property Get Image()
		If Not IsObject(m_Image) Then
			Set m_Image = NewFldObj("Products", "Products", "x_Image", "Image", "[Image]", 202, 8, "", False, False, "IMAGE")
			m_Image.UploadPath = "VSLPayPal/products"
		End If
		Set Image = m_Image
	End Property

	' Field Sizes
	Private m_Sizes

	Public Property Get Sizes()
		If Not IsObject(m_Sizes) Then
			Set m_Sizes = NewFldObj("Products", "Products", "x_Sizes", "Sizes", "[Sizes]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Sizes = m_Sizes
	End Property

	' Field Image_Thumb
	Private m_Image_Thumb

	Public Property Get Image_Thumb()
		If Not IsObject(m_Image_Thumb) Then
			Set m_Image_Thumb = NewFldObj("Products", "Products", "x_Image_Thumb", "Image_Thumb", "[Image_Thumb]", 202, 8, "", False, False, "IMAGE")
			m_Image_Thumb.UploadPath = "/VSLPayPal/products/thumbs"
		End If
		Set Image_Thumb = m_Image_Thumb
	End Property

	' Field ProductName
	Private m_ProductName

	Public Property Get ProductName()
		If Not IsObject(m_ProductName) Then
			Set m_ProductName = NewFldObj("Products", "Products", "x_ProductName", "ProductName", "[ProductName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set ProductName = m_ProductName
	End Property

	' Field ItemNo
	Private m_ItemNo

	Public Property Get ItemNo()
		If Not IsObject(m_ItemNo) Then
			Set m_ItemNo = NewFldObj("Products", "Products", "x_ItemNo", "ItemNo", "[ItemNo]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set ItemNo = m_ItemNo
	End Property

	' Field UPC
	Private m_UPC

	Public Property Get UPC()
		If Not IsObject(m_UPC) Then
			Set m_UPC = NewFldObj("Products", "Products", "x_UPC", "UPC", "[UPC]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set UPC = m_UPC
	End Property

	' Field Price_rebate
	Private m_Price_rebate

	Public Property Get Price_rebate()
		If Not IsObject(m_Price_rebate) Then
			Set m_Price_rebate = NewFldObj("Products", "Products", "x_Price_rebate", "Price_rebate", "[Price_rebate]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Price_rebate = m_Price_rebate
	End Property

	' Field fDescription
	Private m_fDescription

	Public Property Get fDescription()
		If Not IsObject(m_fDescription) Then
			Set m_fDescription = NewFldObj("Products", "Products", "x_fDescription", "fDescription", "[fDescription]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set fDescription = m_fDescription
	End Property

	' Field fImage
	Private m_fImage

	Public Property Get fImage()
		If Not IsObject(m_fImage) Then
			Set m_fImage = NewFldObj("Products", "Products", "x_fImage", "fImage", "[fImage]", 202, 8, "", False, False, "IMAGE")
			m_fImage.UploadPath = "VSLPayPal/French/products/thumbs"
		End If
		Set fImage = m_fImage
	End Property

	' Field fSizes
	Private m_fSizes

	Public Property Get fSizes()
		If Not IsObject(m_fSizes) Then
			Set m_fSizes = NewFldObj("Products", "Products", "x_fSizes", "fSizes", "[fSizes]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set fSizes = m_fSizes
	End Property

	' Field fImage_Thumb
	Private m_fImage_Thumb

	Public Property Get fImage_Thumb()
		If Not IsObject(m_fImage_Thumb) Then
			Set m_fImage_Thumb = NewFldObj("Products", "Products", "x_fImage_Thumb", "fImage_Thumb", "[fImage_Thumb]", 202, 8, "", False, False, "IMAGE")
			m_fImage_Thumb.UploadPath = "VSLPayPal/French/products"
		End If
		Set fImage_Thumb = m_fImage_Thumb
	End Property

	' Field fProductName
	Private m_fProductName

	Public Property Get fProductName()
		If Not IsObject(m_fProductName) Then
			Set m_fProductName = NewFldObj("Products", "Products", "x_fProductName", "fProductName", "[fProductName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set fProductName = m_fProductName
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
		If IsObject(m_ItemId) Then Set m_ItemId = Nothing
		If IsObject(m_Description) Then Set m_Description = Nothing
		If IsObject(m_Price) Then Set m_Price = Nothing
		If IsObject(m_Active) Then Set m_Active = Nothing
		If IsObject(m_Image) Then Set m_Image = Nothing
		If IsObject(m_Sizes) Then Set m_Sizes = Nothing
		If IsObject(m_Image_Thumb) Then Set m_Image_Thumb = Nothing
		If IsObject(m_ProductName) Then Set m_ProductName = Nothing
		If IsObject(m_ItemNo) Then Set m_ItemNo = Nothing
		If IsObject(m_UPC) Then Set m_UPC = Nothing
		If IsObject(m_Price_rebate) Then Set m_Price_rebate = Nothing
		If IsObject(m_fDescription) Then Set m_fDescription = Nothing
		If IsObject(m_fImage) Then Set m_fImage = Nothing
		If IsObject(m_fSizes) Then Set m_fSizes = Nothing
		If IsObject(m_fImage_Thumb) Then Set m_fImage_Thumb = Nothing
		If IsObject(m_fProductName) Then Set m_fProductName = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
