<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Productsinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Products_view
Set Products_view = New cProducts_view
Set Page = Products_view

' Page init processing
Call Products_view.Page_Init()

' Page main processing
Call Products_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Products.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Products_view = new ew_Page("Products_view");
// page properties
Products_view.PageID = "view"; // page ID
Products_view.FormID = "fProductsview"; // form ID
var EW_PAGE_ID = Products_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Products_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Products_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Products_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Products_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<div id="ewDetailsDiv" style="visibility: hidden; z-index: 11000;" name="ewDetailsDivDiv"></div>
<script language="JavaScript" type="text/javascript">
<!--
// YUI container
var ewDetailsDiv;
var ew_AjaxDetailsTimer = null;
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Products_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Products.TableCaption %>
&nbsp;&nbsp;<% Products_view.ExportOptions.Render "body", "" %>
</p>
<% If Products.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Products_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Products_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Products_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Products_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Products_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Products.ItemId.Visible Then ' ItemId %>
	<tr id="r_ItemId"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.ItemId.FldCaption %></td>
		<td<%= Products.ItemId.CellAttributes %>>
<div<%= Products.ItemId.ViewAttributes %>><%= Products.ItemId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.Description.Visible Then ' Description %>
	<tr id="r_Description"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Description.FldCaption %></td>
		<td<%= Products.Description.CellAttributes %>>
<div<%= Products.Description.ViewAttributes %>><%= Products.Description.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.Price.Visible Then ' Price %>
	<tr id="r_Price"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Price.FldCaption %></td>
		<td<%= Products.Price.CellAttributes %>>
<div<%= Products.Price.ViewAttributes %>><%= Products.Price.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Active.FldCaption %></td>
		<td<%= Products.Active.CellAttributes %>>
<% If ew_ConvertToBool(Products.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Products.Active.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Products.Active.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
<% If Products.Image.Visible Then ' Image %>
	<tr id="r_Image"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Image.FldCaption %></td>
		<td<%= Products.Image.CellAttributes %>>
<% If Products.Image.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image.UploadPath) & Products.Image.Upload.DbValue %>" border=0<%= Products.Image.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image.UploadPath) & Products.Image.Upload.DbValue %>" border=0<%= Products.Image.ViewAttributes %>>
<% End If %>
<% End If %>
</td>
	</tr>
<% End If %>
<% If Products.Sizes.Visible Then ' Sizes %>
	<tr id="r_Sizes"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Sizes.FldCaption %></td>
		<td<%= Products.Sizes.CellAttributes %>>
<div<%= Products.Sizes.ViewAttributes %>><%= Products.Sizes.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.Image_Thumb.Visible Then ' Image_Thumb %>
	<tr id="r_Image_Thumb"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Image_Thumb.FldCaption %></td>
		<td<%= Products.Image_Thumb.CellAttributes %>>
<% If Products.Image_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% End If %>
</td>
	</tr>
<% End If %>
<% If Products.ProductName.Visible Then ' ProductName %>
	<tr id="r_ProductName"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.ProductName.FldCaption %></td>
		<td<%= Products.ProductName.CellAttributes %>>
<div<%= Products.ProductName.ViewAttributes %>><%= Products.ProductName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.ItemNo.Visible Then ' ItemNo %>
	<tr id="r_ItemNo"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.ItemNo.FldCaption %></td>
		<td<%= Products.ItemNo.CellAttributes %>>
<div<%= Products.ItemNo.ViewAttributes %>><%= Products.ItemNo.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.UPC.Visible Then ' UPC %>
	<tr id="r_UPC"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.UPC.FldCaption %></td>
		<td<%= Products.UPC.CellAttributes %>>
<div<%= Products.UPC.ViewAttributes %>><%= Products.UPC.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.Price_rebate.Visible Then ' Price_rebate %>
	<tr id="r_Price_rebate"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Price_rebate.FldCaption %></td>
		<td<%= Products.Price_rebate.CellAttributes %>>
<div<%= Products.Price_rebate.ViewAttributes %>><%= Products.Price_rebate.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.fDescription.Visible Then ' fDescription %>
	<tr id="r_fDescription"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fDescription.FldCaption %></td>
		<td<%= Products.fDescription.CellAttributes %>>
<div<%= Products.fDescription.ViewAttributes %>><%= Products.fDescription.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.fImage.Visible Then ' fImage %>
	<tr id="r_fImage"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fImage.FldCaption %></td>
		<td<%= Products.fImage.CellAttributes %>>
<% If Products.fImage.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.fImage.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage.UploadPath) & Products.fImage.Upload.DbValue %>" border=0<%= Products.fImage.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.fImage.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage.UploadPath) & Products.fImage.Upload.DbValue %>" border=0<%= Products.fImage.ViewAttributes %>>
<% End If %>
<% End If %>
</td>
	</tr>
<% End If %>
<% If Products.fSizes.Visible Then ' fSizes %>
	<tr id="r_fSizes"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fSizes.FldCaption %></td>
		<td<%= Products.fSizes.CellAttributes %>>
<div<%= Products.fSizes.ViewAttributes %>><%= Products.fSizes.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Products.fImage_Thumb.Visible Then ' fImage_Thumb %>
	<tr id="r_fImage_Thumb"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fImage_Thumb.FldCaption %></td>
		<td<%= Products.fImage_Thumb.CellAttributes %>>
<% If Products.fImage_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage_Thumb.UploadPath) & Products.fImage_Thumb.Upload.DbValue %>" border=0<%= Products.fImage_Thumb.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage_Thumb.UploadPath) & Products.fImage_Thumb.Upload.DbValue %>" border=0<%= Products.fImage_Thumb.ViewAttributes %>>
<% End If %>
<% End If %>
</td>
	</tr>
<% End If %>
<% If Products.fProductName.Visible Then ' fProductName %>
	<tr id="r_fProductName"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fProductName.FldCaption %></td>
		<td<%= Products.fProductName.CellAttributes %>>
<div<%= Products.fProductName.ViewAttributes %>><%= Products.fProductName.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Products_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Products.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Products_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProducts_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Products"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Products_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Products.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Products.TableVar & "&" ' add page token
	End Property

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

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
		If Products.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Products.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Products.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Products) Then Set Products = New cProducts
		Set Table = Products

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("ItemId").Count > 0 Then
			ew_AddKey RecKey, "ItemId", Request.QueryString("ItemId")
			KeyUrl = KeyUrl & "&ItemId=" & Server.URLEncode(Request.QueryString("ItemId"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Products"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.Tag = "span"
		ExportOptions.Separator = "&nbsp;&nbsp;"
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

		' Page unload event, used in current page
		Call Page_Unload()

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set Products = Nothing
		Set ObjForm = Nothing

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
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False
		If IsPageRequest Then ' Validate request
			If Request.QueryString("ItemId").Count > 0 Then
				Products.ItemId.QueryStringValue = Request.QueryString("ItemId")
			Else
				sReturnUrl = "Productslist.asp" ' Return to list
			End If

			' Get action
			Products.CurrentAction = "I" ' Display form
			Select Case Products.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Productslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Productslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Products.RowType = EW_ROWTYPE_VIEW
		Call Products.ResetAttrs()
		Call RenderRow()
	End Sub
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
				Products.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Products.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Products.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Products.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Products.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Products.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Products.KeyFilter

		' Call Row Selecting event
		Call Products.Row_Selecting(sFilter)

		' Load sql based on filter
		Products.CurrentFilter = sFilter
		sSql = Products.SQL
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
		Call Products.Row_Selected(RsRow)
		Products.ItemId.DbValue = RsRow("ItemId")
		Products.Description.DbValue = RsRow("Description")
		Products.Price.DbValue = RsRow("Price")
		Products.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Products.Image.Upload.DbValue = RsRow("Image")
		Products.Sizes.DbValue = RsRow("Sizes")
		Products.Image_Thumb.Upload.DbValue = RsRow("Image_Thumb")
		Products.ProductName.DbValue = RsRow("ProductName")
		Products.ItemNo.DbValue = RsRow("ItemNo")
		Products.UPC.DbValue = RsRow("UPC")
		Products.Price_rebate.DbValue = RsRow("Price_rebate")
		Products.fDescription.DbValue = RsRow("fDescription")
		Products.fImage.Upload.DbValue = RsRow("fImage")
		Products.fSizes.DbValue = RsRow("fSizes")
		Products.fImage_Thumb.Upload.DbValue = RsRow("fImage_Thumb")
		Products.fProductName.DbValue = RsRow("fProductName")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Products.AddUrl
		EditUrl = Products.EditUrl("")
		CopyUrl = Products.CopyUrl("")
		DeleteUrl = Products.DeleteUrl
		ListUrl = Products.ListUrl

		' Call Row Rendering event
		Call Products.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
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
		' -----------
		'  View  Row
		' -----------

		If Products.RowType = EW_ROWTYPE_VIEW Then ' View row

			' ItemId
			Products.ItemId.ViewValue = Products.ItemId.CurrentValue
			Products.ItemId.ViewCustomAttributes = ""

			' Description
			Products.Description.ViewValue = Products.Description.CurrentValue
			Products.Description.ViewCustomAttributes = ""

			' Price
			Products.Price.ViewValue = Products.Price.CurrentValue
			Products.Price.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Products.Active.CurrentValue) Then
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(1) <> "", Products.Active.FldTagCaption(1), "Yes")
			Else
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(2) <> "", Products.Active.FldTagCaption(2), "No")
			End If
			Products.Active.ViewCustomAttributes = ""

			' Image
			If Not ew_Empty(Products.Image.Upload.DbValue) Then
				Products.Image.ViewValue = Products.Image.Upload.DbValue
				Products.Image.ImageAlt = Products.Image.FldAlt
			Else
				Products.Image.ViewValue = ""
			End If
			Products.Image.ViewCustomAttributes = ""

			' Sizes
			Products.Sizes.ViewValue = Products.Sizes.CurrentValue
			Products.Sizes.ViewCustomAttributes = ""

			' Image_Thumb
			If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then
				Products.Image_Thumb.ViewValue = Products.Image_Thumb.Upload.DbValue
				Products.Image_Thumb.ImageAlt = Products.Image_Thumb.FldAlt
			Else
				Products.Image_Thumb.ViewValue = ""
			End If
			Products.Image_Thumb.ViewCustomAttributes = ""

			' ProductName
			Products.ProductName.ViewValue = Products.ProductName.CurrentValue
			Products.ProductName.ViewCustomAttributes = ""

			' ItemNo
			Products.ItemNo.ViewValue = Products.ItemNo.CurrentValue
			Products.ItemNo.ViewCustomAttributes = ""

			' UPC
			Products.UPC.ViewValue = Products.UPC.CurrentValue
			Products.UPC.ViewCustomAttributes = ""

			' Price_rebate
			Products.Price_rebate.ViewValue = Products.Price_rebate.CurrentValue
			Products.Price_rebate.ViewCustomAttributes = ""

			' fDescription
			Products.fDescription.ViewValue = Products.fDescription.CurrentValue
			Products.fDescription.ViewCustomAttributes = ""

			' fImage
			If Not ew_Empty(Products.fImage.Upload.DbValue) Then
				Products.fImage.ViewValue = Products.fImage.Upload.DbValue
				Products.fImage.ImageAlt = Products.fImage.FldAlt
			Else
				Products.fImage.ViewValue = ""
			End If
			Products.fImage.ViewCustomAttributes = ""

			' fSizes
			Products.fSizes.ViewValue = Products.fSizes.CurrentValue
			Products.fSizes.ViewCustomAttributes = ""

			' fImage_Thumb
			If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then
				Products.fImage_Thumb.ViewValue = Products.fImage_Thumb.Upload.DbValue
				Products.fImage_Thumb.ImageAlt = Products.fImage_Thumb.FldAlt
			Else
				Products.fImage_Thumb.ViewValue = ""
			End If
			Products.fImage_Thumb.ViewCustomAttributes = ""

			' fProductName
			Products.fProductName.ViewValue = Products.fProductName.CurrentValue
			Products.fProductName.ViewCustomAttributes = ""

			' View refer script
			' ItemId

			Products.ItemId.LinkCustomAttributes = ""
			Products.ItemId.HrefValue = ""
			Products.ItemId.TooltipValue = ""

			' Description
			Products.Description.LinkCustomAttributes = ""
			Products.Description.HrefValue = ""
			Products.Description.TooltipValue = ""

			' Price
			Products.Price.LinkCustomAttributes = ""
			Products.Price.HrefValue = ""
			Products.Price.TooltipValue = ""

			' Active
			Products.Active.LinkCustomAttributes = ""
			Products.Active.HrefValue = ""
			Products.Active.TooltipValue = ""

			' Image
			Products.Image.LinkCustomAttributes = ""
			Products.Image.HrefValue = ""
			Products.Image.TooltipValue = ""

			' Sizes
			Products.Sizes.LinkCustomAttributes = ""
			Products.Sizes.HrefValue = ""
			Products.Sizes.TooltipValue = ""

			' Image_Thumb
			Products.Image_Thumb.LinkCustomAttributes = ""
			Products.Image_Thumb.HrefValue = ""
			Products.Image_Thumb.TooltipValue = ""

			' ProductName
			Products.ProductName.LinkCustomAttributes = ""
			Products.ProductName.HrefValue = ""
			Products.ProductName.TooltipValue = ""

			' ItemNo
			Products.ItemNo.LinkCustomAttributes = ""
			Products.ItemNo.HrefValue = ""
			Products.ItemNo.TooltipValue = ""

			' UPC
			Products.UPC.LinkCustomAttributes = ""
			Products.UPC.HrefValue = ""
			Products.UPC.TooltipValue = ""

			' Price_rebate
			Products.Price_rebate.LinkCustomAttributes = ""
			Products.Price_rebate.HrefValue = ""
			Products.Price_rebate.TooltipValue = ""

			' fDescription
			Products.fDescription.LinkCustomAttributes = ""
			Products.fDescription.HrefValue = ""
			Products.fDescription.TooltipValue = ""

			' fImage
			Products.fImage.LinkCustomAttributes = ""
			Products.fImage.HrefValue = ""
			Products.fImage.TooltipValue = ""

			' fSizes
			Products.fSizes.LinkCustomAttributes = ""
			Products.fSizes.HrefValue = ""
			Products.fSizes.TooltipValue = ""

			' fImage_Thumb
			Products.fImage_Thumb.LinkCustomAttributes = ""
			Products.fImage_Thumb.HrefValue = ""
			Products.fImage_Thumb.TooltipValue = ""

			' fProductName
			Products.fProductName.LinkCustomAttributes = ""
			Products.fProductName.HrefValue = ""
			Products.fProductName.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Products.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Products.Row_Rendered()
		End If
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
End Class
%>
