<%

' ASPMaker 6 configuration for Table Customers
Dim Customers
Set Customers = New cCustomers ' Initialize table object

' Define table class
Class cCustomers

	' Define table level constants
	' Table variable
	Public Property Get TableVar()
		TableVar = "Customers"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Customers"
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
		SqlSelect = "SELECT * FROM [Customers]"
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
		SqlKeyFilter = "[CustomerID] = @CustomerID@"
	End Property

	' Return url
	Public Property Get ReturnUrl()
	
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "vslCart.asp"
		End If

	End Property

	Public Property Let ReturnUrl(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = v
	End Property

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("Customersview.asp", "")
	End Function

	' Edit url
	Public Function EditUrl()
		EditUrl = KeyUrl("Customersedit.asp", "")
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl("Customerslist.asp", "a=edit")
	End Function

	' Copy url
	Public Function CopyUrl()
		CopyUrl = KeyUrl("Customersadd.asp", "")
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl("Customerslist.asp", "a=copy")
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("Customersdelete.asp", "")
	End Function

	' Key url
	Public Function KeyUrl(url, action)
		Dim sUrl: sUrl = url & "?"
		If action <> "" Then sUrl = sUrl & action & "&"
		If Not IsNull(CustomerID.CurrentValue) Then
			sUrl = sUrl & "CustomerID=" & Server.URLEncode(CustomerID.CurrentValue)
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
		CustomerID.DbValue = rs("CustomerID")
		Inv_FirstName.DbValue = rs("Inv_FirstName")
		Inv_LastName.DbValue = rs("Inv_LastName")
		Inv_Address.DbValue = rs("Inv_Address")
		Inv_Address2.DbValue = rs("Inv_Address2")
		inv_City.DbValue = rs("inv_City")
		inv_Province.DbValue = rs("inv_Province")
		inv_PostalCode.DbValue = rs("inv_PostalCode")
		inv_Country.DbValue = rs("inv_Country")
		inv_PhoneNumber.DbValue = rs("inv_PhoneNumber")
		inv_EmailAddress.DbValue = rs("inv_EmailAddress")
		inv_Fax.DbValue = rs("inv_Fax")
		Notes.DbValue = rs("Notes")
		UserName.DbValue = rs("UserName")
		passwrd.DbValue = rs("passwrd")
	End Sub

	' Render list row values
	Sub RenderListRow()

		' Inv_FirstName
		Inv_FirstName.ViewValue = Inv_FirstName.CurrentValue
		Inv_FirstName.CssStyle = ""
		Inv_FirstName.CssClass = ""
		Inv_FirstName.ViewCustomAttributes = ""

		' Inv_LastName
		Inv_LastName.ViewValue = Inv_LastName.CurrentValue
		Inv_LastName.CssStyle = ""
		Inv_LastName.CssClass = ""
		Inv_LastName.ViewCustomAttributes = ""

		' Inv_Address
		Inv_Address.ViewValue = Inv_Address.CurrentValue
		Inv_Address.CssStyle = ""
		Inv_Address.CssClass = ""
		Inv_Address.ViewCustomAttributes = ""

		' Inv_Address2
		Inv_Address2.ViewValue = Inv_Address2.CurrentValue
		Inv_Address2.CssStyle = ""
		Inv_Address2.CssClass = ""
		Inv_Address2.ViewCustomAttributes = ""

		' inv_City
		inv_City.ViewValue = inv_City.CurrentValue
		inv_City.CssStyle = ""
		inv_City.CssClass = ""
		inv_City.ViewCustomAttributes = ""

		' inv_Province
		If Not IsNull(inv_Province.CurrentValue) And inv_Province.CurrentValue <> "" Then
			sSqlWrk = "SELECT [Province] FROM [Province] WHERE [Prov] = '" & ew_AdjustSql(inv_Province.CurrentValue) & "'"
			sSqlWrk = sSqlWrk & " ORDER BY [Province] Asc"
			Set rswrk = conn.Execute(sSqlWrk)
			If Not rswrk.Eof Then
				inv_Province.ViewValue = rswrk("Province")
			Else
				inv_Province.ViewValue = inv_Province.CurrentValue
			End If
			rswrk.Close
			Set rswrk = Nothing
		Else
			inv_Province.ViewValue = Null
		End If
		inv_Province.CssStyle = ""
		inv_Province.CssClass = ""
		inv_Province.ViewCustomAttributes = ""

		' inv_EmailAddress
		inv_EmailAddress.ViewValue = inv_EmailAddress.CurrentValue
		inv_EmailAddress.CssStyle = ""
		inv_EmailAddress.CssClass = ""
		inv_EmailAddress.ViewCustomAttributes = ""

		' inv_Fax
		inv_Fax.ViewValue = inv_Fax.CurrentValue
		inv_Fax.CssStyle = ""
		inv_Fax.CssClass = ""
		inv_Fax.ViewCustomAttributes = ""

		' UserName
		UserName.ViewValue = UserName.CurrentValue
		UserName.CssStyle = ""
		UserName.CssClass = ""
		UserName.ViewCustomAttributes = ""

		' Inv_FirstName
		Inv_FirstName.HrefValue = ""

		' Inv_LastName
		Inv_LastName.HrefValue = ""

		' Inv_Address
		Inv_Address.HrefValue = ""

		' Inv_Address2
		Inv_Address2.HrefValue = ""

		' inv_City
		inv_City.HrefValue = ""

		' inv_Province
		inv_Province.HrefValue = ""

		' inv_EmailAddress
		inv_EmailAddress.HrefValue = ""

		' inv_Fax
		inv_Fax.HrefValue = ""

		' UserName
		UserName.HrefValue = ""
	End Sub

	' User id filter
	Private Property Get UserIDFilter()
		UserIDFilter = "[CustomerID]=@UserID@"
	End Property

	' Add user id filter
	Public Function AddUserIDFilter(sFilter, userid)
		Dim sFilterWrk
		sFilterWrk = ""
		sFilterWrk = AddParentUserIDFilter(UserIDFilter, "[CustomerID]", userid)
		If sFilter <> "" Then
			If sFilterWrk <> "" Then
				sFilterWrk = sFilter & " AND " & sFilterWrk
			Else
				sFilterWrk = sFilter
			End If
		End If
		sFilterWrk = Replace(sFilterWrk, "@UserID@", ew_AdjustSql(userid))
		AddUserIDFilter = sFilterWrk
	End Function

	' Add parent user id filter
	Public Function AddParentUserIDFilter(sUserIDFilter, sUserIDFld, userid)
		Dim sWrk, sSql, rs
		sWrk = Replace(EW_PARENT_USER_ID_SQL, "@ParentUserID@", userid)
		If sWrk <> "" Then
			sWrk = sUserIDFld & " IN (" & sWrk & ")"
		End If
		If sUserIDFilter <> "" Then
			sWrk = "((" & sUserIDFilter & ") OR (" & sWrk & "))"
		Else
			sWrk = "(" & sWrk & ")"
		End If
		AddParentUserIDFilter = sWrk
	End Function

	' Get user id subquery
	Public Function GetUserIDSubquery(fld, masterfld, userid)
		Dim sWrk, sSql, rs
		sWrk = ""
		sSql = "SELECT " & masterfld.FldExpression & " FROM [Customers] WHERE " & AddUserIDFilter("", userid)
		sWrk = sSql
		If sWrk <> "" Then
			sWrk = fld.FldExpression & " IN (" & sWrk & ")"
		End If
		GetUserIDSubquery = sWrk
	End Function
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
	' Field CustomerID
	Private m_CustomerID

	Public Property Get CustomerID()
		If Not IsObject(m_CustomerID) Then Set m_CustomerID = NewFldObj("Customers", "x_CustomerID", "CustomerID", "[CustomerID]", 3)
		Set CustomerID = m_CustomerID
	End Property

	' Field Inv_FirstName
	Private m_Inv_FirstName

	Public Property Get Inv_FirstName()
		If Not IsObject(m_Inv_FirstName) Then Set m_Inv_FirstName = NewFldObj("Customers", "x_Inv_FirstName", "Inv_FirstName", "[Inv_FirstName]", 202)
		Set Inv_FirstName = m_Inv_FirstName
	End Property

	' Field Inv_LastName
	Private m_Inv_LastName

	Public Property Get Inv_LastName()
		If Not IsObject(m_Inv_LastName) Then Set m_Inv_LastName = NewFldObj("Customers", "x_Inv_LastName", "Inv_LastName", "[Inv_LastName]", 202)
		Set Inv_LastName = m_Inv_LastName
	End Property

	' Field Inv_Address
	Private m_Inv_Address

	Public Property Get Inv_Address()
		If Not IsObject(m_Inv_Address) Then Set m_Inv_Address = NewFldObj("Customers", "x_Inv_Address", "Inv_Address", "[Inv_Address]", 202)
		Set Inv_Address = m_Inv_Address
	End Property

	' Field Inv_Address2
	Private m_Inv_Address2

	Public Property Get Inv_Address2()
		If Not IsObject(m_Inv_Address2) Then Set m_Inv_Address2 = NewFldObj("Customers", "x_Inv_Address2", "Inv_Address2", "[Inv_Address2]", 202)
		Set Inv_Address2 = m_Inv_Address2
	End Property

	' Field inv_City
	Private m_inv_City

	Public Property Get inv_City()
		If Not IsObject(m_inv_City) Then Set m_inv_City = NewFldObj("Customers", "x_inv_City", "inv_City", "[inv_City]", 202)
		Set inv_City = m_inv_City
	End Property

	' Field inv_Province
	Private m_inv_Province

	Public Property Get inv_Province()
		If Not IsObject(m_inv_Province) Then Set m_inv_Province = NewFldObj("Customers", "x_inv_Province", "inv_Province", "[inv_Province]", 202)
		Set inv_Province = m_inv_Province
	End Property

	' Field inv_PostalCode
	Private m_inv_PostalCode

	Public Property Get inv_PostalCode()
		If Not IsObject(m_inv_PostalCode) Then Set m_inv_PostalCode = NewFldObj("Customers", "x_inv_PostalCode", "inv_PostalCode", "[inv_PostalCode]", 202)
		Set inv_PostalCode = m_inv_PostalCode
	End Property

	' Field inv_Country
	Private m_inv_Country

	Public Property Get inv_Country()
		If Not IsObject(m_inv_Country) Then Set m_inv_Country = NewFldObj("Customers", "x_inv_Country", "inv_Country", "[inv_Country]", 202)
		Set inv_Country = m_inv_Country
	End Property

	' Field inv_PhoneNumber
	Private m_inv_PhoneNumber

	Public Property Get inv_PhoneNumber()
		If Not IsObject(m_inv_PhoneNumber) Then Set m_inv_PhoneNumber = NewFldObj("Customers", "x_inv_PhoneNumber", "inv_PhoneNumber", "[inv_PhoneNumber]", 202)
		Set inv_PhoneNumber = m_inv_PhoneNumber
	End Property

	' Field inv_EmailAddress
	Private m_inv_EmailAddress

	Public Property Get inv_EmailAddress()
		If Not IsObject(m_inv_EmailAddress) Then Set m_inv_EmailAddress = NewFldObj("Customers", "x_inv_EmailAddress", "inv_EmailAddress", "[inv_EmailAddress]", 202)
		Set inv_EmailAddress = m_inv_EmailAddress
	End Property

	' Field inv_Fax
	Private m_inv_Fax

	Public Property Get inv_Fax()
		If Not IsObject(m_inv_Fax) Then Set m_inv_Fax = NewFldObj("Customers", "x_inv_Fax", "inv_Fax", "[inv_Fax]", 202)
		Set inv_Fax = m_inv_Fax
	End Property

	' Field Notes
	Private m_Notes

	Public Property Get Notes()
		If Not IsObject(m_Notes) Then Set m_Notes = NewFldObj("Customers", "x_Notes", "Notes", "[Notes]", 203)
		Set Notes = m_Notes
	End Property

	' Field UserName
	Private m_UserName

	Public Property Get UserName()
		If Not IsObject(m_UserName) Then Set m_UserName = NewFldObj("Customers", "x_UserName", "UserName", "[UserName]", 202)
		Set UserName = m_UserName
	End Property

	' Field passwrd
	Private m_passwrd

	Public Property Get passwrd()
		If Not IsObject(m_passwrd) Then Set m_passwrd = NewFldObj("Customers", "x_passwrd", "passwrd", "[passwrd]", 202)
		Set passwrd = m_passwrd
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
		If IsObject(m_CustomerID) Then Set m_CustomerID = Nothing
		If IsObject(m_Inv_FirstName) Then Set m_Inv_FirstName = Nothing
		If IsObject(m_Inv_LastName) Then Set m_Inv_LastName = Nothing
		If IsObject(m_Inv_Address) Then Set m_Inv_Address = Nothing
		If IsObject(m_Inv_Address2) Then Set m_Inv_Address2 = Nothing
		If IsObject(m_inv_City) Then Set m_inv_City = Nothing
		If IsObject(m_inv_Province) Then Set m_inv_Province = Nothing
		If IsObject(m_inv_PostalCode) Then Set m_inv_PostalCode = Nothing
		If IsObject(m_inv_Country) Then Set m_inv_Country = Nothing
		If IsObject(m_inv_PhoneNumber) Then Set m_inv_PhoneNumber = Nothing
		If IsObject(m_inv_EmailAddress) Then Set m_inv_EmailAddress = Nothing
		If IsObject(m_inv_Fax) Then Set m_inv_Fax = Nothing
		If IsObject(m_Notes) Then Set m_Notes = Nothing
		If IsObject(m_UserName) Then Set m_UserName = Nothing
		If IsObject(m_passwrd) Then Set m_passwrd = Nothing
	End Sub
End Class
%>
