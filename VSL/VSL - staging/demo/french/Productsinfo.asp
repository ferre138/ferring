<%

' ASPMaker 6 configuration for Table Products
Dim Products
Set Products = New cProducts ' Initialize table object

' Define table class
Class cProducts

	' Define table level constants
	' Table variable
	Public Property Get TableVar()
		TableVar = "Products"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Products"
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

	' Advanced search
	Public Function GetAdvancedSearch(fld)
		GetAdvancedSearch = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld)
	End Function

	Public Function SetAdvancedSearch(fld, v)
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) <> v Then
			Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ADVANCED_SEARCH & "_" & fld) = v
		End If
	End Function

	' Basic search Keyword
	Public Property Get BasicSearchKeyword()
		BasicSearchKeyword = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH)
	End Property

	Public Property Let BasicSearchKeyword(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH) = v
	End Property

	' Basic Search Type
	Public Property Get BasicSearchType()
		BasicSearchType = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_BASIC_SEARCH_TYPE)
	End Property

	Public Property Let BasicSearchType(v)
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

	Public Property Get SqlWhere() ' Where
		SqlWhere = ""
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

	' Report table sql
	Public Property Get SQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = SessionOrderBy
		SQL = ew_BuildSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Return table sql with list page filter
	Public Property Get ListSQL()
		Dim sFilter, sSort
		sFilter = SessionWhere
		If CurrentFilter <> "" Then
			If sFilter <> "" Then sFilter = sFilter & " AND "
			sFilter = sFilter & CurrentFilter
		End If
		sSort = SessionOrderBy
		ListSQL = ew_BuildSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Key filter for table
	Public Property Get SqlKeyFilter()
		SqlKeyFilter = "[ItemId] = @ItemId@"
	End Property

	' Return url
	Public Property Get ReturnUrl()
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "Productslist.asp"
		End If
	End Property

	Public Property Let ReturnUrl(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = v
	End Property

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("Productsview.asp", "")
	End Function

	' Edit url
	Public Function EditUrl()
		EditUrl = KeyUrl("Productsedit.asp", "")
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl("Productslist.asp", "a=edit")
	End Function

	' Copy url
	Public Function CopyUrl()
		CopyUrl = KeyUrl("Productsadd.asp", "")
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl("Productslist.asp", "a=copy")
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("Productsdelete.asp", "")
	End Function

	' Key url
	Public Function KeyUrl(url, action)
		Dim sUrl: sUrl = url & "?"
		If action <> "" Then sUrl = sUrl & action & "&"
		If Not IsNull(ItemId.CurrentValue) Then
			sUrl = sUrl & "ItemId=" & Server.URLEncode(ItemId.CurrentValue)
		Else
			KeyUrl = "javascript:alert('Invalid Record! Key is null');"
			Exit Function
		End If
		KeyUrl = sUrl
	End Function

	' Function LoadRs
	' - Load Row based on Key Value
	Public Function LoadRs(sFilter)
		On Error Resume Next
		Dim rs, sSql

		' Set up filter (Sql Where Clause) and get Return Sql
		CurrentFilter = sFilter
		sSql = SQL
		Err.Clear
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = EW_CURSORLOCATION
		rs.Open sSql, conn, 1, 2
		If Err.Number <> 0 Then
			Set LoadRs = Nothing
			rs.Close
			Set rs = Nothing
		ElseIf rs.Eof Then
			Set LoadRs = Nothing
			rs.Close
			Set rs = Nothing
		Else
			Set LoadRs = rs
		End If
	End Function

	' Load row values from rs
	Public Sub LoadListRowValues(rs)
		ItemId.DbValue = rs("ItemId")
		ItemNo.DbValue = rs("ItemNo")
		UPC.DbValue = rs("UPC")
		Image_Thumb.Upload.DbValue = rs("Image_Thumb")
		fImage_Thumb.Upload.DbValue = rs("fImage_Thumb")
		

		ProductName.DbValue = rs("ProductName")
		Description.DbValue = rs("Description")
		fProductName.DbValue = rs("fProductName")
		fDescription.DbValue = rs("fDescription")
		Price.DbValue = rs("Price")
		Active.DbValue = ew_IIf(rs("Active"), "1", "0")
		Image.Upload.DbValue = rs("Image")
		fImage.Upload.DbValue = rs("fImage")
		Sizes.DbValue = rs("Sizes")
	End Sub

	' Render list row values
	Sub RenderListRow()

		' ItemId
		ItemId.ViewValue = ItemId.CurrentValue
		ItemId.CssStyle = ""
		ItemId.CssClass = ""
		ItemId.ViewCustomAttributes = ""

		' ItemNo
		ItemNo.ViewValue = ItemNo.CurrentValue
		ItemNo.CssStyle = ""
		ItemNo.CssClass = ""
		ItemNo.ViewCustomAttributes = ""

		' UPC
		UPC.ViewValue = UPC.CurrentValue
		UPC.CssStyle = ""
		UPC.CssClass = ""
		UPC.ViewCustomAttributes = ""

		' Image_Thumb
		If Not IsNull(Image_Thumb.Upload.DbValue) Then
			Image_Thumb.ViewValue = Image_Thumb.Upload.DbValue
			Image_Thumb.ImageAlt = ""
		Else
			Image_Thumb.ViewValue = ""
		End If
		Image_Thumb.CssStyle = ""
		Image_Thumb.CssClass = ""
		Image_Thumb.ViewCustomAttributes = ""


' Image_Thumb
		If Not IsNull(fImage_Thumb.Upload.DbValue) Then
			fImage_Thumb.ViewValue = fImage_Thumb.Upload.DbValue
			fImage_Thumb.ImageAlt = ""
		Else
			fImage_Thumb.ViewValue = ""
		End If
		fImage_Thumb.CssStyle = ""
		fImage_Thumb.CssClass = ""
		fImage_Thumb.ViewCustomAttributes = ""
		
		' ProductName
		ProductName.ViewValue = ProductName.CurrentValue
		ProductName.CssStyle = ""
		ProductName.CssClass = ""
		ProductName.ViewCustomAttributes = ""
		' ProductName
		fProductName.ViewValue = ProductName.CurrentValue
		fProductName.CssStyle = ""
		fProductName.CssClass = ""
		fProductName.ViewCustomAttributes = ""

		' Description
		Description.ViewValue = Description.CurrentValue
		If Not IsNull(Description.ViewValue) Then
			Description.ViewValue = Replace(Description.ViewValue, vbLf, "<br>")
		End If
		Description.CssStyle = ""
		Description.CssClass = ""
		Description.ViewCustomAttributes = ""

' Description
		fDescription.ViewValue = fDescription.CurrentValue
		If Not IsNull(fDescription.ViewValue) Then
			fDescription.ViewValue = Replace(fDescription.ViewValue, vbLf, "<br>")
		End If
		fDescription.CssStyle = ""
		fDescription.CssClass = ""
		fDescription.ViewCustomAttributes = ""
		' Price
		Price.ViewValue = Price.CurrentValue
		Price.CssStyle = ""
		Price.CssClass = ""
		Price.ViewCustomAttributes = ""

		' Active
		If Active.CurrentValue = "1" Then
			Active.ViewValue = "Yes"
		Else
			Active.ViewValue = "No"
		End If
		Active.CssStyle = ""
		Active.CssClass = ""
		Active.ViewCustomAttributes = ""

		' Image
		If Not IsNull(Image.Upload.DbValue) Then
			Image.ViewValue = Image.Upload.DbValue
			Image.ImageAlt = ""
		Else
			Image.ViewValue = ""
		End If
		Image.CssStyle = ""
		Image.CssClass = ""
		Image.ViewCustomAttributes = ""

		' Sizes
		Sizes.ViewValue = Sizes.CurrentValue
		Sizes.CssStyle = ""
		Sizes.CssClass = ""
		Sizes.ViewCustomAttributes = ""

		' ItemId
		ItemId.HrefValue = ""

		' ItemNo
		ItemNo.HrefValue = ""

		' UPC
		UPC.HrefValue = ""

		' Image_Thumb
		Image_Thumb.HrefValue = ""

		' ProductName
		ProductName.HrefValue = ""

		' Description
		Description.HrefValue = ""
' ProductName
		fProductName.HrefValue = ""

		' Description
		fDescription.HrefValue = ""
		' Price
		Price.HrefValue = ""

		' Active
		Active.HrefValue = ""

		' Image
		Image.HrefValue = ""

		' Sizes
		Sizes.HrefValue = ""
	End Sub
	Dim CurrentAction ' Current action
	Dim EventName ' Event name
	Dim EventCancelled ' Event cancelled
	Dim CancelMessage ' Cancel message

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
	Dim RowClientEvents ' Row client events

	' Display Attribute
	Public Property Get DisplayAttributes()
		Dim sAtt
		sAtt = ""
		If Trim(CssStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(CssStyle) & """" 
		End If
		If Trim(CssClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(CssClass) & """" 
		End If
		If m_Export = "" Then
			If Trim(RowClientEvents) <> "" Then
				sAtt = sAtt & " " & RowClientEvents
			End If
		End If
		DisplayAttributes = sAtt
	End Property

	' Export
	Private m_Export

	Public Property Get Export()
		Export = m_Export
	End Property

	Public Property Let Export(v)
		m_Export = v
	End Property

	' Send Email
	Dim SendEmail

	' ----------------
	'  Field objects
	' ----------------
	' Field ItemId
	Private m_ItemId

	Public Property Get ItemId()
		If Not IsObject(m_ItemId) Then Set m_ItemId = NewFldObj("Products", "x_ItemId", "ItemId", "[ItemId]", 3)
		Set ItemId = m_ItemId
	End Property

	' Field ItemNo
	Private m_ItemNo

	Public Property Get ItemNo()
		If Not IsObject(m_ItemNo) Then Set m_ItemNo = NewFldObj("Products", "x_ItemNo", "ItemNo", "[ItemNo]", 202)
		Set ItemNo = m_ItemNo
	End Property

	' Field UPC
	Private m_UPC

	Public Property Get UPC()
		If Not IsObject(m_UPC) Then Set m_UPC = NewFldObj("Products", "x_UPC", "UPC", "[UPC]", 202)
		Set UPC = m_UPC
	End Property

	' Field Image_Thumb
	Private m_Image_Thumb

	Public Property Get Image_Thumb()
		If Not IsObject(m_Image_Thumb) Then Set m_Image_Thumb = NewFldObj("Products", "x_Image_Thumb", "Image_Thumb", "[Image_Thumb]", 202)
		Set Image_Thumb = m_Image_Thumb
	End Property
' Field Image_Thumb
	Private m_fImage_Thumb

	Public Property Get fImage_Thumb()
		If Not IsObject(m_fImage_Thumb) Then Set m_fImage_Thumb = NewFldObj("Products", "x_fImage_Thumb", "fImage_Thumb", "[fImage_Thumb]", 202)
		Set fImage_Thumb = m_fImage_Thumb
	End Property
	' Field ProductName
	Private m_ProductName

	Public Property Get ProductName()
		If Not IsObject(m_ProductName) Then Set m_ProductName = NewFldObj("Products", "x_ProductName", "ProductName", "[ProductName]", 202)
		Set ProductName = m_ProductName
	End Property

	' Field Description
	Private m_Description

	Public Property Get Description()
		If Not IsObject(m_Description) Then Set m_Description = NewFldObj("Products", "x_Description", "Description", "[Description]", 202)
		Set Description = m_Description
	End Property

' Field ProductName
	Private m_fProductName

	Public Property Get fProductName()
		If Not IsObject(m_fProductName) Then Set m_fProductName = NewFldObj("Products", "x_fProductName", "fProductName", "[fProductName]", 202)
		Set fProductName = m_fProductName
	End Property

	' Field Description
	Private m_fDescription

	Public Property Get fDescription()
		If Not IsObject(m_fDescription) Then Set m_fDescription = NewFldObj("Products", "x_fDescription", "fDescription", "[fDescription]", 202)
		Set fDescription = m_fDescription
	End Property

	' Field Price
	Private m_Price

	Public Property Get Price()
		If Not IsObject(m_Price) Then Set m_Price = NewFldObj("Products", "x_Price", "Price", "[Price]", 202)
		Set Price = m_Price
	End Property

	' Field Active
	Private m_Active

	Public Property Get Active()
		If Not IsObject(m_Active) Then Set m_Active = NewFldObj("Products", "x_Active", "Active", "[Active]", 11)
		Set Active = m_Active
	End Property

	' Field Image
	Private m_Image

	Public Property Get Image()
		If Not IsObject(m_Image) Then Set m_Image = NewFldObj("Products", "x_Image", "Image", "[Image]", 202)
		Set Image = m_Image
	End Property
' Field Image
	Private m_fImage

	Public Property Get fImage()
		If Not IsObject(m_fImage) Then Set m_fImage = NewFldObj("Products", "x_fImage", "fImage", "[fImage]", 202)
		Set fImage = m_fImage
	End Property
	' Field Sizes
	Private m_Sizes

	Public Property Get Sizes()
		If Not IsObject(m_Sizes) Then Set m_Sizes = NewFldObj("Products", "x_Sizes", "Sizes", "[Sizes]", 202)
		Set Sizes = m_Sizes
	End Property

	' Create new field object
	Private Function NewFldObj(tblvar, fldvar, fldname, fldexpression, fldtype)
		Dim fld
		Set fld = New cField
		fld.TblVar = tblvar
		fld.FldVar = fldvar
		fld.FldName = fldname
		fld.FldExpression = fldexpression
		fld.FldType = fldtype
		fld.ImageWidth = 0
		fld.ImageHeight = 0
		Set NewFldObj = fld
	End Function

	' Table level events
	' Recordset Selecting event
	Sub Recordset_Selecting(filter)
		On Error Resume Next

		' Enter your code here	
	End Sub

	' Recordset Selected event
	Sub Recordset_Selected(rs)

	'***Response.Write "Recordset Selected"
	End Sub

	' Row_Selecting event
	Sub Row_Selecting(filter)
		On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Selected event
	Sub Row_Selected(rs)

	'***Response.Write "Row Selected"
	End Sub

	' Row Rendering event
	Sub Row_Rendering()
		On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Rendered event
	Sub Row_Rendered()

		' To view properties of field class, use:
		' Response.Write <FieldName>.AsString() 

	End Sub

	' Row Inserting event
	Function Row_Inserting(rs)
		On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Inserting = True
	End Function

	' Row Inserted event
	Sub Row_Inserted(rs)

		' Response.Write "Row Inserted"
	End Sub

	' Row Updating event
	Function Row_Updating(rsold, rsnew)
		On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Updating = True
	End Function

	' Row Updated event
	Sub Row_Updated(rsold, rsnew)

		' Response.Write "Row Updated"
	End Sub

	' Recordset Deleting event
	Function Recordset_Deleting(rs)
		On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Recordset_Deleting = True
	End Function

	' Recordset Deleted event
	Sub Recordset_Deleted(rs)

		' Response.Write "Recordset Deleted"
	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_ItemId) Then Set m_ItemId = Nothing
		If IsObject(m_ItemNo) Then Set m_ItemNo = Nothing
		If IsObject(m_UPC) Then Set m_UPC = Nothing
		If IsObject(m_Image_Thumb) Then Set m_Image_Thumb = Nothing
			If IsObject(m_fImage_Thumb) Then Set m_fImage_Thumb = Nothing
		If IsObject(m_ProductName) Then Set m_ProductName = Nothing
		If IsObject(m_Description) Then Set m_Description = Nothing
		If IsObject(m_PfroductName) Then Set m_fProductName = Nothing
		If IsObject(m_fescription) Then Set m_fDescription = Nothing
		If IsObject(m_Price) Then Set m_Price = Nothing
		If IsObject(m_Active) Then Set m_Active = Nothing
		If IsObject(m_Image) Then Set m_Image = Nothing
		If IsObject(m_fImage) Then Set m_fImage = Nothing
		If IsObject(m_Sizes) Then Set m_Sizes = Nothing
	End Sub
End Class
%>
