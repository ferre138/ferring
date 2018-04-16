<%

' ASPMaker configuration for Table Customers
Dim Customers

' Define table class
Class cCustomers

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
		Call ew_SetArObj(Fields, "CustomerID", CustomerID)
		Call ew_SetArObj(Fields, "Inv_FirstName", Inv_FirstName)
		Call ew_SetArObj(Fields, "Inv_LastName", Inv_LastName)
		Call ew_SetArObj(Fields, "Inv_Address", Inv_Address)
		Call ew_SetArObj(Fields, "inv_City", inv_City)
		Call ew_SetArObj(Fields, "inv_Province", inv_Province)
		Call ew_SetArObj(Fields, "inv_PostalCode", inv_PostalCode)
		Call ew_SetArObj(Fields, "inv_Country", inv_Country)
		Call ew_SetArObj(Fields, "inv_PhoneNumber", inv_PhoneNumber)
		Call ew_SetArObj(Fields, "inv_EmailAddress", inv_EmailAddress)
		Call ew_SetArObj(Fields, "Notes", Notes)
		Call ew_SetArObj(Fields, "inv_Fax", inv_Fax)
		Call ew_SetArObj(Fields, "Inv_Address2", Inv_Address2)
		Call ew_SetArObj(Fields, "UserName", UserName)
		Call ew_SetArObj(Fields, "passwrd", passwrd)
		Call ew_SetArObj(Fields, "NewCustomer", NewCustomer)
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
		TableVar = "Customers"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Customers"
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
		HighlightName = "Customers_Highlight"
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
		If CurrentDetailTable = "Orders" Then
			sDetailUrl = Orders.ListUrl & "?showmaster=" & TableVar
			sDetailUrl = sDetailUrl & "&CustomerId=" & CustomerID.CurrentValue
		End If
		DetailUrl = sDetailUrl
	End Property

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [Customers]"
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
		SqlOrderBy = "[CustomerID] DESC"
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
		SqlKeyFilter = "[CustomerID] = @CustomerID@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(CustomerID.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@CustomerID@", ew_AdjustSql(CustomerID.CurrentValue)) ' Replace key value
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
			ReturnUrl = "Customerslist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "Customerslist.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("Customersview.asp", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "Customersadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		If parm <> "" Then
			EditUrl = KeyUrl("Customersedit.asp", UrlParm(parm))
		Else
			EditUrl = KeyUrl("Customersedit.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		If parm <> "" Then
			CopyUrl = KeyUrl("Customersadd.asp", UrlParm(parm))
		Else
			CopyUrl = KeyUrl("Customersadd.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("Customersdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(CustomerID.CurrentValue) Then
			sUrl = sUrl & "CustomerID=" & CustomerID.CurrentValue
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
			UrlParm = "t=Customers"
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
				arKeys(0) = Request.QueryString("CustomerID") ' CustomerID

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
				CustomerID.CurrentValue = key
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
		CustomerID.DbValue = RsRow("CustomerID")
		Inv_FirstName.DbValue = RsRow("Inv_FirstName")
		Inv_LastName.DbValue = RsRow("Inv_LastName")
		Inv_Address.DbValue = RsRow("Inv_Address")
		inv_City.DbValue = RsRow("inv_City")
		inv_Province.DbValue = RsRow("inv_Province")
		inv_PostalCode.DbValue = RsRow("inv_PostalCode")
		inv_Country.DbValue = RsRow("inv_Country")
		inv_PhoneNumber.DbValue = RsRow("inv_PhoneNumber")
		inv_EmailAddress.DbValue = RsRow("inv_EmailAddress")
		Notes.DbValue = RsRow("Notes")
		inv_Fax.DbValue = RsRow("inv_Fax")
		Inv_Address2.DbValue = RsRow("Inv_Address2")
		UserName.DbValue = RsRow("UserName")
		passwrd.DbValue = RsRow("passwrd")
		NewCustomer.DbValue = ew_IIf(RsRow("NewCustomer"), "1", "0")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' CustomerID
		' Inv_FirstName
		' Inv_LastName
		' Inv_Address
		' inv_City
		' inv_Province
		' inv_PostalCode
		' inv_Country
		' inv_PhoneNumber
		' inv_EmailAddress
		' Notes
		' inv_Fax
		' Inv_Address2
		' UserName
		' passwrd
		' NewCustomer
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' CustomerID

		CustomerID.ViewValue = CustomerID.CurrentValue
		CustomerID.ViewCustomAttributes = ""

		' Inv_FirstName
		Inv_FirstName.ViewValue = Inv_FirstName.CurrentValue
		Inv_FirstName.ViewCustomAttributes = ""

		' Inv_LastName
		Inv_LastName.ViewValue = Inv_LastName.CurrentValue
		Inv_LastName.ViewCustomAttributes = ""

		' Inv_Address
		Inv_Address.ViewValue = Inv_Address.CurrentValue
		Inv_Address.ViewCustomAttributes = ""

		' inv_City
		inv_City.ViewValue = inv_City.CurrentValue
		inv_City.ViewCustomAttributes = ""

		' inv_Province
		inv_Province.ViewValue = inv_Province.CurrentValue
		inv_Province.ViewCustomAttributes = ""

		' inv_PostalCode
		inv_PostalCode.ViewValue = inv_PostalCode.CurrentValue
		inv_PostalCode.ViewCustomAttributes = ""

		' inv_Country
		inv_Country.ViewValue = inv_Country.CurrentValue
		inv_Country.ViewCustomAttributes = ""

		' inv_PhoneNumber
		inv_PhoneNumber.ViewValue = inv_PhoneNumber.CurrentValue
		inv_PhoneNumber.ViewCustomAttributes = ""

		' inv_EmailAddress
		inv_EmailAddress.ViewValue = inv_EmailAddress.CurrentValue
		inv_EmailAddress.ViewCustomAttributes = ""

		' Notes
		Notes.ViewValue = Notes.CurrentValue
		Notes.ViewCustomAttributes = ""

		' inv_Fax
		inv_Fax.ViewValue = inv_Fax.CurrentValue
		inv_Fax.ViewCustomAttributes = ""

		' Inv_Address2
		Inv_Address2.ViewValue = Inv_Address2.CurrentValue
		Inv_Address2.ViewCustomAttributes = ""

		' UserName
		UserName.ViewValue = UserName.CurrentValue
		UserName.ViewCustomAttributes = ""

		' passwrd
		passwrd.ViewValue = passwrd.CurrentValue
		passwrd.ViewCustomAttributes = ""

		' NewCustomer
		If ew_ConvertToBool(NewCustomer.CurrentValue) Then
			NewCustomer.ViewValue = ew_IIf(NewCustomer.FldTagCaption(1) <> "", NewCustomer.FldTagCaption(1), "Yes")
		Else
			NewCustomer.ViewValue = ew_IIf(NewCustomer.FldTagCaption(2) <> "", NewCustomer.FldTagCaption(2), "No")
		End If
		NewCustomer.ViewCustomAttributes = ""

		' CustomerID
		CustomerID.LinkCustomAttributes = ""
		CustomerID.HrefValue = ""
		CustomerID.TooltipValue = ""

		' Inv_FirstName
		Inv_FirstName.LinkCustomAttributes = ""
		Inv_FirstName.HrefValue = ""
		Inv_FirstName.TooltipValue = ""

		' Inv_LastName
		Inv_LastName.LinkCustomAttributes = ""
		Inv_LastName.HrefValue = ""
		Inv_LastName.TooltipValue = ""

		' Inv_Address
		Inv_Address.LinkCustomAttributes = ""
		Inv_Address.HrefValue = ""
		Inv_Address.TooltipValue = ""

		' inv_City
		inv_City.LinkCustomAttributes = ""
		inv_City.HrefValue = ""
		inv_City.TooltipValue = ""

		' inv_Province
		inv_Province.LinkCustomAttributes = ""
		inv_Province.HrefValue = ""
		inv_Province.TooltipValue = ""

		' inv_PostalCode
		inv_PostalCode.LinkCustomAttributes = ""
		inv_PostalCode.HrefValue = ""
		inv_PostalCode.TooltipValue = ""

		' inv_Country
		inv_Country.LinkCustomAttributes = ""
		inv_Country.HrefValue = ""
		inv_Country.TooltipValue = ""

		' inv_PhoneNumber
		inv_PhoneNumber.LinkCustomAttributes = ""
		inv_PhoneNumber.HrefValue = ""
		inv_PhoneNumber.TooltipValue = ""

		' inv_EmailAddress
		inv_EmailAddress.LinkCustomAttributes = ""
		inv_EmailAddress.HrefValue = ""
		inv_EmailAddress.TooltipValue = ""

		' Notes
		Notes.LinkCustomAttributes = ""
		Notes.HrefValue = ""
		Notes.TooltipValue = ""

		' inv_Fax
		inv_Fax.LinkCustomAttributes = ""
		inv_Fax.HrefValue = ""
		inv_Fax.TooltipValue = ""

		' Inv_Address2
		Inv_Address2.LinkCustomAttributes = ""
		Inv_Address2.HrefValue = ""
		Inv_Address2.TooltipValue = ""

		' UserName
		UserName.LinkCustomAttributes = ""
		UserName.HrefValue = ""
		UserName.TooltipValue = ""

		' passwrd
		passwrd.LinkCustomAttributes = ""
		passwrd.HrefValue = ""
		passwrd.TooltipValue = ""

		' NewCustomer
		NewCustomer.LinkCustomAttributes = ""
		NewCustomer.HrefValue = ""
		NewCustomer.TooltipValue = ""

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
					Call XmlDoc.AddField("CustomerID", CustomerID.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_FirstName", Inv_FirstName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_LastName", Inv_LastName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_Address", Inv_Address.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_City", inv_City.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Province", inv_Province.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_PostalCode", inv_PostalCode.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Country", inv_Country.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_PhoneNumber", inv_PhoneNumber.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_EmailAddress", inv_EmailAddress.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Notes", Notes.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Fax", inv_Fax.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_Address2", Inv_Address2.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("UserName", UserName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("passwrd", passwrd.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("NewCustomer", NewCustomer.ExportValue(Export, ExportOriginalValue))
				Else
					Call XmlDoc.AddField("CustomerID", CustomerID.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_FirstName", Inv_FirstName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_LastName", Inv_LastName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_Address", Inv_Address.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_City", inv_City.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Province", inv_Province.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_PostalCode", inv_PostalCode.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Country", inv_Country.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_PhoneNumber", inv_PhoneNumber.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_EmailAddress", inv_EmailAddress.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("inv_Fax", inv_Fax.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("Inv_Address2", Inv_Address2.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("UserName", UserName.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("passwrd", passwrd.ExportValue(Export, ExportOriginalValue))
					Call XmlDoc.AddField("NewCustomer", NewCustomer.ExportValue(Export, ExportOriginalValue))
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
				Call Doc.ExportCaption(CustomerID)
				Call Doc.ExportCaption(Inv_FirstName)
				Call Doc.ExportCaption(Inv_LastName)
				Call Doc.ExportCaption(Inv_Address)
				Call Doc.ExportCaption(inv_City)
				Call Doc.ExportCaption(inv_Province)
				Call Doc.ExportCaption(inv_PostalCode)
				Call Doc.ExportCaption(inv_Country)
				Call Doc.ExportCaption(inv_PhoneNumber)
				Call Doc.ExportCaption(inv_EmailAddress)
				Call Doc.ExportCaption(Notes)
				Call Doc.ExportCaption(inv_Fax)
				Call Doc.ExportCaption(Inv_Address2)
				Call Doc.ExportCaption(UserName)
				Call Doc.ExportCaption(passwrd)
				Call Doc.ExportCaption(NewCustomer)
			Else
				Call Doc.ExportCaption(CustomerID)
				Call Doc.ExportCaption(Inv_FirstName)
				Call Doc.ExportCaption(Inv_LastName)
				Call Doc.ExportCaption(Inv_Address)
				Call Doc.ExportCaption(inv_City)
				Call Doc.ExportCaption(inv_Province)
				Call Doc.ExportCaption(inv_PostalCode)
				Call Doc.ExportCaption(inv_Country)
				Call Doc.ExportCaption(inv_PhoneNumber)
				Call Doc.ExportCaption(inv_EmailAddress)
				Call Doc.ExportCaption(inv_Fax)
				Call Doc.ExportCaption(Inv_Address2)
				Call Doc.ExportCaption(UserName)
				Call Doc.ExportCaption(passwrd)
				Call Doc.ExportCaption(NewCustomer)
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
					Call Doc.ExportField(CustomerID)
					Call Doc.ExportField(Inv_FirstName)
					Call Doc.ExportField(Inv_LastName)
					Call Doc.ExportField(Inv_Address)
					Call Doc.ExportField(inv_City)
					Call Doc.ExportField(inv_Province)
					Call Doc.ExportField(inv_PostalCode)
					Call Doc.ExportField(inv_Country)
					Call Doc.ExportField(inv_PhoneNumber)
					Call Doc.ExportField(inv_EmailAddress)
					Call Doc.ExportField(Notes)
					Call Doc.ExportField(inv_Fax)
					Call Doc.ExportField(Inv_Address2)
					Call Doc.ExportField(UserName)
					Call Doc.ExportField(passwrd)
					Call Doc.ExportField(NewCustomer)
				Else
					Call Doc.ExportField(CustomerID)
					Call Doc.ExportField(Inv_FirstName)
					Call Doc.ExportField(Inv_LastName)
					Call Doc.ExportField(Inv_Address)
					Call Doc.ExportField(inv_City)
					Call Doc.ExportField(inv_Province)
					Call Doc.ExportField(inv_PostalCode)
					Call Doc.ExportField(inv_Country)
					Call Doc.ExportField(inv_PhoneNumber)
					Call Doc.ExportField(inv_EmailAddress)
					Call Doc.ExportField(inv_Fax)
					Call Doc.ExportField(Inv_Address2)
					Call Doc.ExportField(UserName)
					Call Doc.ExportField(passwrd)
					Call Doc.ExportField(NewCustomer)
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
	' Field CustomerID
	Private m_CustomerID

	Public Property Get CustomerID()
		If Not IsObject(m_CustomerID) Then
			Set m_CustomerID = NewFldObj("Customers", "Customers", "x_CustomerID", "CustomerID", "[CustomerID]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_CustomerID.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set CustomerID = m_CustomerID
	End Property

	' Field Inv_FirstName
	Private m_Inv_FirstName

	Public Property Get Inv_FirstName()
		If Not IsObject(m_Inv_FirstName) Then
			Set m_Inv_FirstName = NewFldObj("Customers", "Customers", "x_Inv_FirstName", "Inv_FirstName", "[Inv_FirstName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_FirstName = m_Inv_FirstName
	End Property

	' Field Inv_LastName
	Private m_Inv_LastName

	Public Property Get Inv_LastName()
		If Not IsObject(m_Inv_LastName) Then
			Set m_Inv_LastName = NewFldObj("Customers", "Customers", "x_Inv_LastName", "Inv_LastName", "[Inv_LastName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_LastName = m_Inv_LastName
	End Property

	' Field Inv_Address
	Private m_Inv_Address

	Public Property Get Inv_Address()
		If Not IsObject(m_Inv_Address) Then
			Set m_Inv_Address = NewFldObj("Customers", "Customers", "x_Inv_Address", "Inv_Address", "[Inv_Address]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_Address = m_Inv_Address
	End Property

	' Field inv_City
	Private m_inv_City

	Public Property Get inv_City()
		If Not IsObject(m_inv_City) Then
			Set m_inv_City = NewFldObj("Customers", "Customers", "x_inv_City", "inv_City", "[inv_City]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_City = m_inv_City
	End Property

	' Field inv_Province
	Private m_inv_Province

	Public Property Get inv_Province()
		If Not IsObject(m_inv_Province) Then
			Set m_inv_Province = NewFldObj("Customers", "Customers", "x_inv_Province", "inv_Province", "[inv_Province]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_Province = m_inv_Province
	End Property

	' Field inv_PostalCode
	Private m_inv_PostalCode

	Public Property Get inv_PostalCode()
		If Not IsObject(m_inv_PostalCode) Then
			Set m_inv_PostalCode = NewFldObj("Customers", "Customers", "x_inv_PostalCode", "inv_PostalCode", "[inv_PostalCode]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_PostalCode = m_inv_PostalCode
	End Property

	' Field inv_Country
	Private m_inv_Country

	Public Property Get inv_Country()
		If Not IsObject(m_inv_Country) Then
			Set m_inv_Country = NewFldObj("Customers", "Customers", "x_inv_Country", "inv_Country", "[inv_Country]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_Country = m_inv_Country
	End Property

	' Field inv_PhoneNumber
	Private m_inv_PhoneNumber

	Public Property Get inv_PhoneNumber()
		If Not IsObject(m_inv_PhoneNumber) Then
			Set m_inv_PhoneNumber = NewFldObj("Customers", "Customers", "x_inv_PhoneNumber", "inv_PhoneNumber", "[inv_PhoneNumber]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_PhoneNumber = m_inv_PhoneNumber
	End Property

	' Field inv_EmailAddress
	Private m_inv_EmailAddress

	Public Property Get inv_EmailAddress()
		If Not IsObject(m_inv_EmailAddress) Then
			Set m_inv_EmailAddress = NewFldObj("Customers", "Customers", "x_inv_EmailAddress", "inv_EmailAddress", "[inv_EmailAddress]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_EmailAddress = m_inv_EmailAddress
	End Property

	' Field Notes
	Private m_Notes

	Public Property Get Notes()
		If Not IsObject(m_Notes) Then
			Set m_Notes = NewFldObj("Customers", "Customers", "x_Notes", "Notes", "[Notes]", 203, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Notes = m_Notes
	End Property

	' Field inv_Fax
	Private m_inv_Fax

	Public Property Get inv_Fax()
		If Not IsObject(m_inv_Fax) Then
			Set m_inv_Fax = NewFldObj("Customers", "Customers", "x_inv_Fax", "inv_Fax", "[inv_Fax]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_Fax = m_inv_Fax
	End Property

	' Field Inv_Address2
	Private m_Inv_Address2

	Public Property Get Inv_Address2()
		If Not IsObject(m_Inv_Address2) Then
			Set m_Inv_Address2 = NewFldObj("Customers", "Customers", "x_Inv_Address2", "Inv_Address2", "[Inv_Address2]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_Address2 = m_Inv_Address2
	End Property

	' Field UserName
	Private m_UserName

	Public Property Get UserName()
		If Not IsObject(m_UserName) Then
			Set m_UserName = NewFldObj("Customers", "Customers", "x_UserName", "UserName", "[UserName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set UserName = m_UserName
	End Property

	' Field passwrd
	Private m_passwrd

	Public Property Get passwrd()
		If Not IsObject(m_passwrd) Then
			Set m_passwrd = NewFldObj("Customers", "Customers", "x_passwrd", "passwrd", "[passwrd]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set passwrd = m_passwrd
	End Property

	' Field NewCustomer
	Private m_NewCustomer

	Public Property Get NewCustomer()
		If Not IsObject(m_NewCustomer) Then
			Set m_NewCustomer = NewFldObj("Customers", "Customers", "x_NewCustomer", "NewCustomer", "[NewCustomer]", 11, 8, "", False, False, "FORMATTED TEXT")
			m_NewCustomer.FldDataType = EW_DATATYPE_BOOLEAN
		End If
		Set NewCustomer = m_NewCustomer
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
		If IsObject(m_CustomerID) Then Set m_CustomerID = Nothing
		If IsObject(m_Inv_FirstName) Then Set m_Inv_FirstName = Nothing
		If IsObject(m_Inv_LastName) Then Set m_Inv_LastName = Nothing
		If IsObject(m_Inv_Address) Then Set m_Inv_Address = Nothing
		If IsObject(m_inv_City) Then Set m_inv_City = Nothing
		If IsObject(m_inv_Province) Then Set m_inv_Province = Nothing
		If IsObject(m_inv_PostalCode) Then Set m_inv_PostalCode = Nothing
		If IsObject(m_inv_Country) Then Set m_inv_Country = Nothing
		If IsObject(m_inv_PhoneNumber) Then Set m_inv_PhoneNumber = Nothing
		If IsObject(m_inv_EmailAddress) Then Set m_inv_EmailAddress = Nothing
		If IsObject(m_Notes) Then Set m_Notes = Nothing
		If IsObject(m_inv_Fax) Then Set m_inv_Fax = Nothing
		If IsObject(m_Inv_Address2) Then Set m_Inv_Address2 = Nothing
		If IsObject(m_UserName) Then Set m_UserName = Nothing
		If IsObject(m_passwrd) Then Set m_passwrd = Nothing
		If IsObject(m_NewCustomer) Then Set m_NewCustomer = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
