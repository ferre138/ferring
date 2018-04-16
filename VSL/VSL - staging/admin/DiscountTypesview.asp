<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="DiscountTypesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim DiscountTypes_view
Set DiscountTypes_view = New cDiscountTypes_view
Set Page = DiscountTypes_view

' Page init processing
Call DiscountTypes_view.Page_Init()

' Page main processing
Call DiscountTypes_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If DiscountTypes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var DiscountTypes_view = new ew_Page("DiscountTypes_view");
// page properties
DiscountTypes_view.PageID = "view"; // page ID
DiscountTypes_view.FormID = "fDiscountTypesview"; // form ID
var EW_PAGE_ID = DiscountTypes_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
DiscountTypes_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
DiscountTypes_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
DiscountTypes_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
DiscountTypes_view.ValidateRequired = false; // no JavaScript validation
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
<% DiscountTypes_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= DiscountTypes.TableCaption %>
&nbsp;&nbsp;<% DiscountTypes_view.ExportOptions.Render "body", "" %>
</p>
<% If DiscountTypes.Export = "" Then %>
<p class="aspmaker">
<a href="<%= DiscountTypes_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= DiscountTypes_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= DiscountTypes_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= DiscountTypes_view.CopyUrl %>"><%= Language.Phrase("ViewPageCopyLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= DiscountTypes_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% DiscountTypes_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If DiscountTypes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<tr id="r_DiscountTypeId"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountTypeId.FldCaption %></td>
		<td<%= DiscountTypes.DiscountTypeId.CellAttributes %>>
<div<%= DiscountTypes.DiscountTypeId.ViewAttributes %>><%= DiscountTypes.DiscountTypeId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
	<tr id="r_DiscountType"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountType.FldCaption %></td>
		<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
	<tr id="r_DiscountTitle"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountTitle.FldCaption %></td>
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
	<tr id="r_freeShipping"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.freeShipping.FldCaption %></td>
		<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
	<tr id="r_FreePerQty"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.FreePerQty.FldCaption %></td>
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
	<tr id="r_SpecialPrice"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.SpecialPrice.FldCaption %></td>
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
DiscountTypes_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If DiscountTypes.Export = "" Then %>
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
Set DiscountTypes_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountTypes_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "DiscountTypes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "DiscountTypes_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If DiscountTypes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & DiscountTypes.TableVar & "&" ' add page token
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
		If DiscountTypes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (DiscountTypes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (DiscountTypes.TableVar = Request.QueryString("t"))
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
		If IsEmpty(DiscountTypes) Then Set DiscountTypes = New cDiscountTypes
		Set Table = DiscountTypes

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("DiscountTypeId").Count > 0 Then
			ew_AddKey RecKey, "DiscountTypeId", Request.QueryString("DiscountTypeId")
			KeyUrl = KeyUrl & "&DiscountTypeId=" & Server.URLEncode(Request.QueryString("DiscountTypeId"))
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
		EW_TABLE_NAME = "DiscountTypes"

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
		Set DiscountTypes = Nothing
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
			If Request.QueryString("DiscountTypeId").Count > 0 Then
				DiscountTypes.DiscountTypeId.QueryStringValue = Request.QueryString("DiscountTypeId")
			Else
				sReturnUrl = "DiscountTypeslist.asp" ' Return to list
			End If

			' Get action
			DiscountTypes.CurrentAction = "I" ' Display form
			Select Case DiscountTypes.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "DiscountTypeslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "DiscountTypeslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		DiscountTypes.RowType = EW_ROWTYPE_VIEW
		Call DiscountTypes.ResetAttrs()
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
				DiscountTypes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					DiscountTypes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = DiscountTypes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			DiscountTypes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			DiscountTypes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			DiscountTypes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = DiscountTypes.KeyFilter

		' Call Row Selecting event
		Call DiscountTypes.Row_Selecting(sFilter)

		' Load sql based on filter
		DiscountTypes.CurrentFilter = sFilter
		sSql = DiscountTypes.SQL
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
		Call DiscountTypes.Row_Selected(RsRow)
		DiscountTypes.DiscountTypeId.DbValue = RsRow("DiscountTypeId")
		DiscountTypes.DiscountType.DbValue = RsRow("DiscountType")
		DiscountTypes.DiscountTitle.DbValue = RsRow("DiscountTitle")
		DiscountTypes.freeShipping.DbValue = ew_IIf(RsRow("freeShipping"), "1", "0")
		DiscountTypes.FreePerQty.DbValue = RsRow("FreePerQty")
		DiscountTypes.SpecialPrice.DbValue = RsRow("SpecialPrice")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = DiscountTypes.AddUrl
		EditUrl = DiscountTypes.EditUrl("")
		CopyUrl = DiscountTypes.CopyUrl("")
		DeleteUrl = DiscountTypes.DeleteUrl
		ListUrl = DiscountTypes.ListUrl

		' Call Row Rendering event
		Call DiscountTypes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' DiscountTypeId
		' DiscountType
		' DiscountTitle
		' freeShipping
		' FreePerQty
		' SpecialPrice
		' -----------
		'  View  Row
		' -----------

		If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' DiscountTypeId
			DiscountTypes.DiscountTypeId.ViewValue = DiscountTypes.DiscountTypeId.CurrentValue
			DiscountTypes.DiscountTypeId.ViewCustomAttributes = ""

			' DiscountType
			DiscountTypes.DiscountType.ViewValue = DiscountTypes.DiscountType.CurrentValue
			DiscountTypes.DiscountType.ViewCustomAttributes = ""

			' DiscountTitle
			DiscountTypes.DiscountTitle.ViewValue = DiscountTypes.DiscountTitle.CurrentValue
			DiscountTypes.DiscountTitle.ViewCustomAttributes = ""

			' freeShipping
			If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then
				DiscountTypes.freeShipping.ViewValue = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(1) <> "", DiscountTypes.freeShipping.FldTagCaption(1), "Yes")
			Else
				DiscountTypes.freeShipping.ViewValue = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(2) <> "", DiscountTypes.freeShipping.FldTagCaption(2), "No")
			End If
			DiscountTypes.freeShipping.ViewCustomAttributes = ""

			' FreePerQty
			DiscountTypes.FreePerQty.ViewValue = DiscountTypes.FreePerQty.CurrentValue
			DiscountTypes.FreePerQty.ViewCustomAttributes = ""

			' SpecialPrice
			DiscountTypes.SpecialPrice.ViewValue = DiscountTypes.SpecialPrice.CurrentValue
			DiscountTypes.SpecialPrice.ViewCustomAttributes = ""

			' View refer script
			' DiscountTypeId

			DiscountTypes.DiscountTypeId.LinkCustomAttributes = ""
			DiscountTypes.DiscountTypeId.HrefValue = ""
			DiscountTypes.DiscountTypeId.TooltipValue = ""

			' DiscountType
			DiscountTypes.DiscountType.LinkCustomAttributes = ""
			DiscountTypes.DiscountType.HrefValue = ""
			DiscountTypes.DiscountType.TooltipValue = ""

			' DiscountTitle
			DiscountTypes.DiscountTitle.LinkCustomAttributes = ""
			DiscountTypes.DiscountTitle.HrefValue = ""
			DiscountTypes.DiscountTitle.TooltipValue = ""

			' freeShipping
			DiscountTypes.freeShipping.LinkCustomAttributes = ""
			DiscountTypes.freeShipping.HrefValue = ""
			DiscountTypes.freeShipping.TooltipValue = ""

			' FreePerQty
			DiscountTypes.FreePerQty.LinkCustomAttributes = ""
			DiscountTypes.FreePerQty.HrefValue = ""
			DiscountTypes.FreePerQty.TooltipValue = ""

			' SpecialPrice
			DiscountTypes.SpecialPrice.LinkCustomAttributes = ""
			DiscountTypes.SpecialPrice.HrefValue = ""
			DiscountTypes.SpecialPrice.TooltipValue = ""
		End If

		' Call Row Rendered event
		If DiscountTypes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call DiscountTypes.Row_Rendered()
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
