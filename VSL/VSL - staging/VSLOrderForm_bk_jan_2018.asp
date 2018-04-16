<%
Const EW_PAGE_ID = "list"
Const EW_TABLE_NAME = "Products"
%>
<!--#include file="ewcfg60.asp"--> 
<!--#include file="Productsinfo.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%

' Open connection to the database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
%>

<%

' Common page loading event (in userfn60.asp)
Call Page_Loading()
%>
<%

' Page load event, used in current page
Call Page_Load()
%>
<%
Products.Export = Request.QueryString("export") ' Get export parameter
sExport = Products.Export ' Get export parameter, used in header
sExportFile = Products.TableVar ' Get export file, used in header
%>
<%

' Paging variables
Dim Pager, PagerItem ' Pager
Dim nDisplayRecs ' Number of display records
Dim nRecRange ' Record display range
Dim nStartRec, nStopRec, nTotalRecs
nStartRec = 0 ' Start record index
nStopRec = 0 ' Stop record index
nTotalRecs = 0 ' Total number of records
nDisplayRecs = 20
nRecRange = 10
Dim i
Dim nRecCount
nRecCount = 0 ' Record count
Dim RowCnt, RowIndex, OptionCnt

' Sort
Dim sSortOrder

' Search filters
Dim sSrchAdvanced, sSrchBasic, sSrchWhere, sFilter
sSrchAdvanced = "" ' Advanced search filter
sSrchBasic = "" ' Basic search filter
sSrchWhere = "" ' Search where clause
sFilter = ""
Dim bEditRow, nEditRowCnt ' Edit row

' Master/Detail
Dim sDbMasterFilter, sDbDetailFilter
sDbMasterFilter = "" ' Master filter
sDbDetailFilter = "" ' Detail filter
Dim sSqlMaster
sSqlMaster = "" ' Sql for master record

' Handle reset command
ResetCmd()

' Get basic search criteria
sSrchBasic = BasicSearchWhere()

' Build search criteria
If sSrchAdvanced <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchAdvanced & ")"
End If
If sSrchBasic <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchBasic & ")"
End If

' Save search criteria
If sSrchWhere <> "" Then
	If sSrchBasic = "" Then Call ResetBasicSearchParms()
	Products.SearchWhere = sSrchWhere ' Save to Session
	nStartRec = 1 ' Reset start record counter
	Products.StartRecordNumber = nStartRec
Else
	Call RestoreSearchParms()
End If

' Build filter
sFilter = "active=yes"
If sDbDetailFilter <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sDbDetailFilter & ")"
End If
If sSrchWhere <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sSrchWhere & ")"
End If

' Set up filter in Session
Products.SessionWhere = sFilter
Products.CurrentFilter = ""

' Set Up Sorting Order
SetUpSortOrder()

' Set Return Url
Products.ReturnUrl = "vslOrderForm.asp"
%>
<!--#include file="header.asp"-->


<table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
            <tr>
            <td  ><h1 class="h1-header" style="padding:20px 0px 20px 0px; color:#0067A7;">Order VSL#3 Here</h1></td></tr><tr>
              <td  ><div align="left"><span style="color:#333333;font-weight: bold;">Please note:   VSL#3 orders are shipped Monday, Tuesday and Wednesday's only.<br/>  
Any orders received Wednesday after 10am EST; up to and including Friday, will ship the following Monday.</span><br/>
</div></td>
           <!--   <td width="28" valign="top"><img src="images/FontSize.png" border="0" alt=""> </td>
              <td width="24" valign="top"> <a href="#"
				onmouseover="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseout="changeImages('login_13', 'images/font1.png'); return true;"
				onmousedown="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseup="changeImages('login_13', 'images/login_13-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('default'); 
return false;"> <img name="login_13" src="images/font1.png" width="24" height="27" border="0" alt=""></a></td>
              <td width="24"  valign="top"> <a href="#"
				onmouseover="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseout="changeImages('login_14', 'images/font2.png'); return true;"
				onmousedown="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseup="changeImages('login_14', 'images/login_14-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Medium'); 
return false;"> <img name="login_14" src="images/font2.png" width="24" height="27" border="0" alt=""></a></td>
              <td width="26"  valign="top"> <a href="#"
				onmouseover="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseout="changeImages('login_15', 'images/font3.png'); return true;"
				onmousedown="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseup="changeImages('login_15', 'images/login_15-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Large'); 
return false;"><img name="login_15" src="images/font3.png" width="24" height="27" border="0" alt=""></a></td>
            </tr>
            <tr>
              <td colspan="4" valign="top"><div align="right">
                <p><a href="french/VSLOrderForm.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
              </div></td>-->
              </tr>
        </table>
        


 
<script type="text/javascript">
<!--
var EW_PAGE_ID = "list"; // Page id
//-->
</script>
<script type="text/javascript">
<!--
var firstrowoffset = 1; // First data row start at
var lastrowoffset = 0; // Last data row end at
var EW_LIST_TABLE_NAME = 'ewlistmain'; // Table name for list page
var rowclass = 'ewTableRow'; // Row class
var rowaltclass = 'ewTableAltRow'; // Row alternate class
var rowmoverclass = 'ewTableHighlightRow'; // Row mouse over class
var rowselectedclass = 'ewTableSelectRow'; // Row selected class
var roweditclass = 'ewTableEditRow'; // Row edit class
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>


<%

' Load recordset
Dim rs
Set rs = LoadRecordset()
nTotalRecs = rs.RecordCount

nStartRec = 1
If nDisplayRecs <= 0 Then ' Display all records
	nDisplayRecs = nTotalRecs
End If
If Not (EW_EXPORT_ALL And Products.Export <> "") Then
	SetUpStartRec() ' Set up start record position
End If
%>

<table width="785" height="179" border="0" align="center" cellpadding="0" cellspacing="0" class="ewTableNoBorder" id="ewlistmain">
<tr><td colspan="3"><span class="ewmsg">We are closed for the Holiday Season. Back in January 2nd, 2018.<br>VSL#3 is also available at most pharmacies and health food stores.<br> Visit our <a href="store-locator.html">Store Locator</a> to find one near you.<br><br>Thank you and Happy Holidays!
<br>
</span>
</td></tr>
</table>




<%

' Close recordset and connection
rs.Close
Set rs = Nothing
%>


<!--#include file="footer.asp"-->

<%

' If control is passed here, simply terminate the page without redirect
Call Page_Terminate("")

' -----------------------------------------------------------------
'  Subroutine Page_Terminate
'  - called when exit page
'  - clean up ADO connection and objects
'  - if url specified, redirect to url, otherwise end response
'
Sub Page_Terminate(url)

	' Page unload event, used in current page
	Call Page_Unload()

	' Global page unloaded event (in userfn60.asp)
	Call Page_Unloaded()
	conn.Close ' Close Connection
	Set conn = Nothing
	Set Security = Nothing
	Set Products = Nothing

	' Go to url if specified
	If url <> "" Then
		Response.Clear
		Response.Redirect url
	End If

	' Terminate response
	Response.End
End Sub

'
'  Subroutine Page_Terminate (End)
' ----------------------------------------

%>
<%

' Return Basic Search sql
Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = ew_AdjustSql(Keyword)
	BasicSearchSQL = ""
	BasicSearchSQL = BasicSearchSQL & "[ItemNo] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[UPC] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Image_Thumb] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[ProductName] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Description] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Price] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Image] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Sizes] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

' Return Basic Search Where based on search keyword and type
Function BasicSearchWhere()
	Dim sSearchStr, sSearchKeyword, sSearchType
	Dim sSearch, arKeyword, sKeyword
	sSearchStr = ""
	sSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
	sSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	If sSearchKeyword <> "" Then
		sSearch = Trim(sSearchKeyword)
		If sSearchType <> "" Then
			While InStr(sSearch, "  ") > 0
				sSearch = Replace(sSearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(sSearch), " ")
			For Each sKeyword In arKeyword
				If sSearchStr <> "" Then sSearchStr = sSearchStr & " " & sSearchType & " "
				sSearchStr = sSearchStr & "(" & BasicSearchSQL(sKeyword) & ")"
			Next
		Else
			sSearchStr = BasicSearchSQL(sSearch)
		End If
	End If
	If sSearchKeyword <> "" then
		Products.BasicSearchKeyword = sSearchKeyword
		Products.BasicSearchType = sSearchType
	End If
	BasicSearchWhere = sSearchStr
End Function

' Clear all search parameters
Sub ResetSearchParms()

	' Clear search where
	sSrchWhere = ""
	Products.SearchWhere = sSrchWhere

	' Clear basic search parameters
	Call ResetBasicSearchParms()
End Sub

' Clear all basic search parameters
Sub ResetBasicSearchParms()

	' Clear basic search parameters
	Products.BasicSearchKeyword = ""
	Products.BasicSearchType = ""
End Sub

' Restore all search parameters
Sub RestoreSearchParms()
	sSrchWhere = Products.SearchWhere
End Sub

' Set up Sort parameters based on Sort Links clicked
Sub SetUpSortOrder()
	Dim sOrderBy
	Dim sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		Products.CurrentOrder = Request.QueryString("order")
		Products.CurrentOrderType = Request.QueryString("ordertype")

		' Field ItemId
		Call Products.UpdateSort(Products.ItemId)

		' Field ItemNo
		Call Products.UpdateSort(Products.ItemNo)

		' Field UPC
		Call Products.UpdateSort(Products.UPC)

		' Field Image_Thumb
		Call Products.UpdateSort(Products.Image_Thumb)

		' Field ProductName
		Call Products.UpdateSort(Products.ProductName)

		' Field Description
		Call Products.UpdateSort(Products.Description)

		' Field Price
		Call Products.UpdateSort(Products.Price)

		' Field Active
		Call Products.UpdateSort(Products.Active)

		' Field Image
		Call Products.UpdateSort(Products.Image)

		' Field Sizes
		Call Products.UpdateSort(Products.Sizes)
		Products.StartRecordNumber = 1 ' Reset start position
	End If
	sOrderBy = Products.SessionOrderBy ' Get order by from Session
	If sOrderBy = "" Then
		If Products.SqlOrderBy <> "" Then
			sOrderBy = Products.SqlOrderBy
			Products.SessionOrderBy = sOrderBy
		End If
	End If
End Sub

' Reset command based on querystring parameter cmd=
' - RESET: reset search parameters
' - RESETALL: reset search & master/detail parameters
' - RESETSORT: reset sort parameters
Sub ResetCmd()
	Dim sCmd

	' Get reset cmd
	If Request.QueryString("cmd").Count > 0 Then
		sCmd = Request.QueryString("cmd")

		' Reset search criteria
		If LCase(sCmd) = "reset" Or LCase(sCmd) = "resetall" Then
			Call ResetSearchParms()
		End If

		' Reset Sort Criteria
		If LCase(sCmd) = "resetsort" Then
			Dim sOrderBy
			sOrderBy = ""
			Products.SessionOrderBy = sOrderBy
			Products.ItemId.Sort = ""
			Products.ItemNo.Sort = ""
			Products.UPC.Sort = ""
			Products.Image_Thumb.Sort = ""
			Products.ProductName.Sort = ""
			Products.Description.Sort = ""
			Products.Price.Sort = ""
			Products.Active.Sort = ""
			Products.Image.Sort = ""
			Products.Sizes.Sort = ""
		End If

		' Reset start position
		nStartRec = 1
		Products.StartRecordNumber = nStartRec
	End If
End Sub
%>
<%

' Set up Starting Record parameters based on Pager Navigation
Sub SetUpStartRec()
	Dim nPageNo

	' Exit if nDisplayRecs = 0
	If nDisplayRecs = 0 Then Exit Sub

	' Check for a START parameter
	If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
		nStartRec = Request.QueryString(EW_TABLE_START_REC)
		Products.StartRecordNumber = nStartRec
	ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
		nPageNo = Request.QueryString(EW_TABLE_PAGE_NO)
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Products.StartRecordNumber = nStartRec
		Else
			nStartRec = Products.StartRecordNumber
		End If
	Else
		nStartRec = Products.StartRecordNumber
	End If

	' Check if correct start record counter
	If Not IsNumeric(nStartRec) Or nStartRec = "" Then ' Avoid invalid start record counter
		nStartRec = 1 ' Reset start record counter
		Products.StartRecordNumber = nStartRec
	ElseIf CLng(nStartRec) > CLng(nTotalRecs) Then ' Avoid starting record > total records
		nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to last page first record
		Products.StartRecordNumber = nStartRec
	ElseIf (nStartRec-1) Mod nDisplayRecs <> 0 Then
		nStartRec = ((nStartRec-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to page boundary
		Products.StartRecordNumber = nStartRec
	End If
End Sub
%>
<%

' Load recordset
Function LoadRecordset()

	' Call Recordset Selecting event
	Call Products.Recordset_Selecting(Products.CurrentFilter)

	' Load list page sql
	Dim sSql
	sSql = Products.ListSQL & "Order by Description desc"

	 'Response.Write sSql ' Uncomment to show SQL for debugging
	' Load recordset

	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = EW_CURSORLOCATION
	rs.Open sSql, conn, 1, 2

	' Call Recordset Selected event
	Call Products.Recordset_Selected(rs)
	Set LoadRecordset = rs
End Function
%>
<%

' Load row based on key values
Function LoadRow()
	Dim rs, sSql, sFilter
	sFilter = Products.SqlKeyFilter
	If Not IsNumeric(Products.ItemId.CurrentValue) Then
		LoadRow = False ' Invalid key, exit
		Exit Function
	End If
	sFilter = Replace(sFilter, "@ItemId@", ew_AdjustSql(Products.ItemId.CurrentValue)) ' Replace key value

	' Call Row Selecting event
	Call Products.Row_Selecting(sFilter)

	' Load sql based on filter
	Products.CurrentFilter = sFilter
	sSql = Products.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadRow = False
	Else
		LoadRow = True
		rs.MoveFirst
		Call LoadRowValues(rs) ' Load row values

		' Call Row Selected event
		Call Products.Row_Selected(rs)
	End If
	rs.Close
	Set rs = Nothing
End Function

' Load row values from recordset
Sub LoadRowValues(rs)
	Products.ItemId.DbValue = rs("ItemId")
	Products.ItemNo.DbValue = rs("ItemNo")
	Products.UPC.DbValue = rs("UPC")
	Products.Image_Thumb.Upload.DbValue = rs("Image_Thumb")
	Products.ProductName.DbValue = rs("ProductName")
	Products.Description.DbValue = rs("Description")
	Products.Price.DbValue = rs("Price")
	Products.Active.DbValue = ew_IIf(rs("Active"), "1", "0")
	Products.Image.Upload.DbValue = rs("Image")
	Products.Sizes.DbValue = rs("Sizes")
End Sub
%>
<%

' Render row values based on field settings
Sub RenderRow()

	' Call Row Rendering event
	Call Products.Row_Rendering()

	' Common render codes for all row types
	' ItemId

	Products.ItemId.CellCssStyle = ""
	Products.ItemId.CellCssClass = ""

	' ItemNo
	Products.ItemNo.CellCssStyle = ""
	Products.ItemNo.CellCssClass = ""

	' UPC
	Products.UPC.CellCssStyle = ""
	Products.UPC.CellCssClass = ""

	' Image_Thumb
	Products.Image_Thumb.CellCssStyle = ""
	Products.Image_Thumb.CellCssClass = ""

	' ProductName
	Products.ProductName.CellCssStyle = ""
	Products.ProductName.CellCssClass = ""

	' Description
	Products.Description.CellCssStyle = ""
	Products.Description.CellCssClass = ""

	' Price
	Products.Price.CellCssStyle = ""
	Products.Price.CellCssClass = ""

	' Active
	Products.Active.CellCssStyle = ""
	Products.Active.CellCssClass = ""

	' Image
	Products.Image.CellCssStyle = ""
	Products.Image.CellCssClass = ""

	' Sizes
	Products.Sizes.CellCssStyle = ""
	Products.Sizes.CellCssClass = ""
	If Products.RowType = EW_ROWTYPE_VIEW Then ' View row

		' ItemId
		Products.ItemId.ViewValue = Products.ItemId.CurrentValue
		Products.ItemId.CssStyle = ""
		Products.ItemId.CssClass = ""
		Products.ItemId.ViewCustomAttributes = ""

		' ItemNo
		Products.ItemNo.ViewValue = Products.ItemNo.CurrentValue
		Products.ItemNo.CssStyle = ""
		Products.ItemNo.CssClass = ""
		Products.ItemNo.ViewCustomAttributes = ""

		' UPC
		Products.UPC.ViewValue = Products.UPC.CurrentValue
		Products.UPC.CssStyle = ""
		Products.UPC.CssClass = ""
		Products.UPC.ViewCustomAttributes = ""

		' Image_Thumb
		If Not IsNull(Products.Image_Thumb.Upload.DbValue) Then
			Products.Image_Thumb.ViewValue = Products.Image_Thumb.Upload.DbValue
			Products.Image_Thumb.ImageAlt = ""
		Else
			Products.Image_Thumb.ViewValue = ""
		End If
		Products.Image_Thumb.CssStyle = ""
		Products.Image_Thumb.CssClass = ""
		Products.Image_Thumb.ViewCustomAttributes = ""

		' ProductName
		Products.ProductName.ViewValue = Products.ProductName.CurrentValue
		Products.ProductName.CssStyle = ""
		Products.ProductName.CssClass = ""
		Products.ProductName.ViewCustomAttributes = ""

		' Description
		Products.Description.ViewValue = Products.Description.CurrentValue
		If Not IsNull(Products.Description.ViewValue) Then
			Products.Description.ViewValue = Replace(Products.Description.ViewValue, vbLf, "<br>")
		End If
		Products.Description.CssStyle = ""
		Products.Description.CssClass = ""
		Products.Description.ViewCustomAttributes = ""

		' Price
		Products.Price.ViewValue = Products.Price.CurrentValue
		Products.Price.CssStyle = ""
		Products.Price.CssClass = ""
		Products.Price.ViewCustomAttributes = ""

		' Active
		If Products.Active.CurrentValue = "1" Then
			Products.Active.ViewValue = "Yes"
		Else
			Products.Active.ViewValue = "No"
		End If
		Products.Active.CssStyle = ""
		Products.Active.CssClass = ""
		Products.Active.ViewCustomAttributes = ""

		' Image
		If Not IsNull(Products.Image.Upload.DbValue) Then
			Products.Image.ViewValue = Products.Image.Upload.DbValue
			Products.Image.ImageAlt = ""
		Else
			Products.Image.ViewValue = ""
		End If
		Products.Image.CssStyle = ""
		Products.Image.CssClass = ""
		Products.Image.ViewCustomAttributes = ""

		' Sizes
		Products.Sizes.ViewValue = Products.Sizes.CurrentValue
		Products.Sizes.CssStyle = ""
		Products.Sizes.CssClass = ""
		Products.Sizes.ViewCustomAttributes = ""

		' ItemId
		Products.ItemId.HrefValue = ""

		' ItemNo
		Products.ItemNo.HrefValue = ""

		' UPC
		Products.UPC.HrefValue = ""

		' Image_Thumb
		Products.Image_Thumb.HrefValue = ""

		' ProductName
		Products.ProductName.HrefValue = ""

		' Description
		Products.Description.HrefValue = ""

		' Price
		Products.Price.HrefValue = ""

		' Active
		Products.Active.HrefValue = ""

		' Image
		Products.Image.HrefValue = ""

		' Sizes
		Products.Sizes.HrefValue = ""
	ElseIf Products.RowType = EW_ROWTYPE_ADD Then ' Add row
	ElseIf Products.RowType = EW_ROWTYPE_EDIT Then ' Edit row
	ElseIf Products.RowType = EW_ROWTYPE_SEARCH Then ' Search row
	End If

	' Call Row Rendered event
	Call Products.Row_Rendered()
End Sub
%>
<%

' Page Load event
Sub Page_Load()

'***Response.Write "Page Load"
End Sub

' Page Unload event
Sub Page_Unload()

'***Response.Write "Page Unload"
End Sub
%>
