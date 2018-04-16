<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Provinceinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Province_view
Set Province_view = New cProvince_view
Set Page = Province_view

' Page init processing
Call Province_view.Page_Init()

' Page main processing
Call Province_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Province.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Province_view = new ew_Page("Province_view");
// page properties
Province_view.PageID = "view"; // page ID
Province_view.FormID = "fProvinceview"; // form ID
var EW_PAGE_ID = Province_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Province_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Province_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Province_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Province_view.ValidateRequired = false; // no JavaScript validation
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
<% Province_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Province.TableCaption %>
&nbsp;&nbsp;<% Province_view.ExportOptions.Render "body", "" %>
</p>
<% If Province.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Province_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Province_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Province_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Province.Prov.Visible Then ' Prov %>
	<tr id="r_Prov"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.Prov.FldCaption %></td>
		<td<%= Province.Prov.CellAttributes %>>
<div<%= Province.Prov.ViewAttributes %>><%= Province.Prov.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Province.Province_1.Visible Then ' Province %>
	<tr id="r_Province_1"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.Province_1.FldCaption %></td>
		<td<%= Province.Province_1.CellAttributes %>>
<div<%= Province.Province_1.ViewAttributes %>><%= Province.Province_1.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Province.fProvince.Visible Then ' fProvince %>
	<tr id="r_fProvince"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.fProvince.FldCaption %></td>
		<td<%= Province.fProvince.CellAttributes %>>
<div<%= Province.fProvince.ViewAttributes %>><%= Province.fProvince.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Province.TaxRate.Visible Then ' TaxRate %>
	<tr id="r_TaxRate"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.TaxRate.FldCaption %></td>
		<td<%= Province.TaxRate.CellAttributes %>>
<div<%= Province.TaxRate.ViewAttributes %>><%= Province.TaxRate.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Province.ShipRate_first.Visible Then ' ShipRate_first %>
	<tr id="r_ShipRate_first"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.ShipRate_first.FldCaption %></td>
		<td<%= Province.ShipRate_first.CellAttributes %>>
<div<%= Province.ShipRate_first.ViewAttributes %>><%= Province.ShipRate_first.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Province.ShipRate_Rest.Visible Then ' ShipRate_Rest %>
	<tr id="r_ShipRate_Rest"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.ShipRate_Rest.FldCaption %></td>
		<td<%= Province.ShipRate_Rest.CellAttributes %>>
<div<%= Province.ShipRate_Rest.ViewAttributes %>><%= Province.ShipRate_Rest.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Province_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Province.Export = "" Then %>
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
Set Province_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProvince_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Province"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Province_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Province.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Province.TableVar & "&" ' add page token
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
		If Province.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Province.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Province.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Province) Then Set Province = New cProvince
		Set Table = Province

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("Prov").Count > 0 Then
			ew_AddKey RecKey, "Prov", Request.QueryString("Prov")
			KeyUrl = KeyUrl & "&Prov=" & Server.URLEncode(Request.QueryString("Prov"))
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
		EW_TABLE_NAME = "Province"

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
		Set Province = Nothing
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
			If Request.QueryString("Prov").Count > 0 Then
				Province.Prov.QueryStringValue = Request.QueryString("Prov")
			Else
				sReturnUrl = "Provincelist.asp" ' Return to list
			End If

			' Get action
			Province.CurrentAction = "I" ' Display form
			Select Case Province.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Provincelist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Provincelist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Province.RowType = EW_ROWTYPE_VIEW
		Call Province.ResetAttrs()
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
				Province.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Province.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Province.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Province.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Province.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Province.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Province.KeyFilter

		' Call Row Selecting event
		Call Province.Row_Selecting(sFilter)

		' Load sql based on filter
		Province.CurrentFilter = sFilter
		sSql = Province.SQL
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
		Call Province.Row_Selected(RsRow)
		Province.Prov.DbValue = RsRow("Prov")
		Province.Province_1.DbValue = RsRow("Province")
		Province.fProvince.DbValue = RsRow("fProvince")
		Province.TaxRate.DbValue = RsRow("TaxRate")
		Province.ShipRate_first.DbValue = RsRow("ShipRate_first")
		Province.ShipRate_Rest.DbValue = RsRow("ShipRate_Rest")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Province.AddUrl
		EditUrl = Province.EditUrl("")
		CopyUrl = Province.CopyUrl("")
		DeleteUrl = Province.DeleteUrl
		ListUrl = Province.ListUrl

		' Call Row Rendering event
		Call Province.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Prov
		' Province
		' fProvince
		' TaxRate
		' ShipRate_first
		' ShipRate_Rest
		' -----------
		'  View  Row
		' -----------

		If Province.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Prov
			Province.Prov.ViewValue = Province.Prov.CurrentValue
			Province.Prov.ViewCustomAttributes = ""

			' Province
			Province.Province_1.ViewValue = Province.Province_1.CurrentValue
			Province.Province_1.ViewCustomAttributes = ""

			' fProvince
			Province.fProvince.ViewValue = Province.fProvince.CurrentValue
			Province.fProvince.ViewCustomAttributes = ""

			' TaxRate
			Province.TaxRate.ViewValue = Province.TaxRate.CurrentValue
			Province.TaxRate.ViewCustomAttributes = ""

			' ShipRate_first
			Province.ShipRate_first.ViewValue = Province.ShipRate_first.CurrentValue
			Province.ShipRate_first.ViewCustomAttributes = ""

			' ShipRate_Rest
			Province.ShipRate_Rest.ViewValue = Province.ShipRate_Rest.CurrentValue
			Province.ShipRate_Rest.ViewCustomAttributes = ""

			' View refer script
			' Prov

			Province.Prov.LinkCustomAttributes = ""
			Province.Prov.HrefValue = ""
			Province.Prov.TooltipValue = ""

			' Province
			Province.Province_1.LinkCustomAttributes = ""
			Province.Province_1.HrefValue = ""
			Province.Province_1.TooltipValue = ""

			' fProvince
			Province.fProvince.LinkCustomAttributes = ""
			Province.fProvince.HrefValue = ""
			Province.fProvince.TooltipValue = ""

			' TaxRate
			Province.TaxRate.LinkCustomAttributes = ""
			Province.TaxRate.HrefValue = ""
			Province.TaxRate.TooltipValue = ""

			' ShipRate_first
			Province.ShipRate_first.LinkCustomAttributes = ""
			Province.ShipRate_first.HrefValue = ""
			Province.ShipRate_first.TooltipValue = ""

			' ShipRate_Rest
			Province.ShipRate_Rest.LinkCustomAttributes = ""
			Province.ShipRate_Rest.HrefValue = ""
			Province.ShipRate_Rest.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Province.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Province.Row_Rendered()
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
