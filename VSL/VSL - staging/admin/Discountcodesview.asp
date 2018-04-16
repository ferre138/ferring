<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Discountcodesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Discountcodes_view
Set Discountcodes_view = New cDiscountcodes_view
Set Page = Discountcodes_view

' Page init processing
Call Discountcodes_view.Page_Init()

' Page main processing
Call Discountcodes_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Discountcodes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Discountcodes_view = new ew_Page("Discountcodes_view");
// page properties
Discountcodes_view.PageID = "view"; // page ID
Discountcodes_view.FormID = "fDiscountcodesview"; // form ID
var EW_PAGE_ID = Discountcodes_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Discountcodes_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Discountcodes_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Discountcodes_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Discountcodes_view.ValidateRequired = false; // no JavaScript validation
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
<% Discountcodes_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Discountcodes.TableCaption %>
&nbsp;&nbsp;<% Discountcodes_view.ExportOptions.Render "body", "" %>
</p>
<% If Discountcodes.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Discountcodes_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Discountcodes_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Discountcodes_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Discountcodes_view.CopyUrl %>"><%= Language.Phrase("ViewPageCopyLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Discountcodes_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Discountcodes_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Discountcodes.Discountid.Visible Then ' Discountid %>
	<tr id="r_Discountid"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Discountid.FldCaption %></td>
		<td<%= Discountcodes.Discountid.CellAttributes %>>
<div<%= Discountcodes.Discountid.ViewAttributes %>><%= Discountcodes.Discountid.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
	<tr id="r_DiscountCode"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountCode.FldCaption %></td>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>>
<div<%= Discountcodes.DiscountCode.ViewAttributes %>><%= Discountcodes.DiscountCode.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Discountcodes.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Active.FldCaption %></td>
		<td<%= Discountcodes.Active.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.Active.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.Active.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
<% If Discountcodes.used.Visible Then ' used %>
	<tr id="r_used"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.used.FldCaption %></td>
		<td<%= Discountcodes.used.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.used.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.used.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
<% If Discountcodes.OrderId.Visible Then ' OrderId %>
	<tr id="r_OrderId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.OrderId.FldCaption %></td>
		<td<%= Discountcodes.OrderId.CellAttributes %>>
<div<%= Discountcodes.OrderId.ViewAttributes %>><%= Discountcodes.OrderId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Discountcodes.Use_date.Visible Then ' Use_date %>
	<tr id="r_Use_date"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Use_date.FldCaption %></td>
		<td<%= Discountcodes.Use_date.CellAttributes %>>
<div<%= Discountcodes.Use_date.ViewAttributes %>><%= Discountcodes.Use_date.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<tr id="r_DiscountTypeId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountTypeId.FldCaption %></td>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Discountcodes_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Discountcodes.Export = "" Then %>
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
Set Discountcodes_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountcodes_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Discountcodes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Discountcodes_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Discountcodes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Discountcodes.TableVar & "&" ' add page token
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
		If Discountcodes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Discountcodes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Discountcodes.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes
		Set Table = Discountcodes

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("Discountid").Count > 0 Then
			ew_AddKey RecKey, "Discountid", Request.QueryString("Discountid")
			KeyUrl = KeyUrl & "&Discountid=" & Server.URLEncode(Request.QueryString("Discountid"))
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
		EW_TABLE_NAME = "Discountcodes"

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
		Set Discountcodes = Nothing
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
			If Request.QueryString("Discountid").Count > 0 Then
				Discountcodes.Discountid.QueryStringValue = Request.QueryString("Discountid")
			Else
				sReturnUrl = "Discountcodeslist.asp" ' Return to list
			End If

			' Get action
			Discountcodes.CurrentAction = "I" ' Display form
			Select Case Discountcodes.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Discountcodeslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Discountcodeslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Discountcodes.RowType = EW_ROWTYPE_VIEW
		Call Discountcodes.ResetAttrs()
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
				Discountcodes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Discountcodes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Discountcodes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Discountcodes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Discountcodes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Discountcodes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Discountcodes.KeyFilter

		' Call Row Selecting event
		Call Discountcodes.Row_Selecting(sFilter)

		' Load sql based on filter
		Discountcodes.CurrentFilter = sFilter
		sSql = Discountcodes.SQL
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
		Call Discountcodes.Row_Selected(RsRow)
		Discountcodes.Discountid.DbValue = RsRow("Discountid")
		Discountcodes.DiscountCode.DbValue = RsRow("DiscountCode")
		Discountcodes.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Discountcodes.used.DbValue = ew_IIf(RsRow("used"), "1", "0")
		Discountcodes.OrderId.DbValue = RsRow("OrderId")
		Discountcodes.Use_date.DbValue = RsRow("Use_date")
		Discountcodes.DiscountTypeId.DbValue = RsRow("DiscountTypeId")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Discountcodes.AddUrl
		EditUrl = Discountcodes.EditUrl("")
		CopyUrl = Discountcodes.CopyUrl("")
		DeleteUrl = Discountcodes.DeleteUrl
		ListUrl = Discountcodes.ListUrl

		' Call Row Rendering event
		Call Discountcodes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Discountid
		' DiscountCode
		' Active
		' used
		' OrderId
		' Use_date
		' DiscountTypeId
		' -----------
		'  View  Row
		' -----------

		If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Discountid
			Discountcodes.Discountid.ViewValue = Discountcodes.Discountid.CurrentValue
			Discountcodes.Discountid.ViewCustomAttributes = ""

			' DiscountCode
			Discountcodes.DiscountCode.ViewValue = Discountcodes.DiscountCode.CurrentValue
			Discountcodes.DiscountCode.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			Else
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			End If
			Discountcodes.Active.ViewCustomAttributes = ""

			' used
			If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			Else
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			End If
			Discountcodes.used.ViewCustomAttributes = ""

			' OrderId
			Discountcodes.OrderId.ViewValue = Discountcodes.OrderId.CurrentValue
			Discountcodes.OrderId.ViewCustomAttributes = ""

			' Use_date
			Discountcodes.Use_date.ViewValue = Discountcodes.Use_date.CurrentValue
			Discountcodes.Use_date.ViewCustomAttributes = ""

			' DiscountTypeId
			If Discountcodes.DiscountTypeId.CurrentValue & "" <> "" Then
				sFilterWrk = "[DiscountTypeId] = " & ew_AdjustSql(Discountcodes.DiscountTypeId.CurrentValue) & ""
			sSqlWrk = "SELECT [DiscountType] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Discountcodes.DiscountTypeId.ViewValue = RsWrk("DiscountType")
				Else
					Discountcodes.DiscountTypeId.ViewValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Discountcodes.DiscountTypeId.ViewValue = Null
			End If
			Discountcodes.DiscountTypeId.ViewCustomAttributes = ""

			' View refer script
			' Discountid

			Discountcodes.Discountid.LinkCustomAttributes = ""
			Discountcodes.Discountid.HrefValue = ""
			Discountcodes.Discountid.TooltipValue = ""

			' DiscountCode
			Discountcodes.DiscountCode.LinkCustomAttributes = ""
			Discountcodes.DiscountCode.HrefValue = ""
			Discountcodes.DiscountCode.TooltipValue = ""

			' Active
			Discountcodes.Active.LinkCustomAttributes = ""
			Discountcodes.Active.HrefValue = ""
			Discountcodes.Active.TooltipValue = ""

			' used
			Discountcodes.used.LinkCustomAttributes = ""
			Discountcodes.used.HrefValue = ""
			Discountcodes.used.TooltipValue = ""

			' OrderId
			Discountcodes.OrderId.LinkCustomAttributes = ""
			Discountcodes.OrderId.HrefValue = ""
			Discountcodes.OrderId.TooltipValue = ""

			' Use_date
			Discountcodes.Use_date.LinkCustomAttributes = ""
			Discountcodes.Use_date.HrefValue = ""
			Discountcodes.Use_date.TooltipValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.LinkCustomAttributes = ""
			Discountcodes.DiscountTypeId.HrefValue = ""
			Discountcodes.DiscountTypeId.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Discountcodes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Discountcodes.Row_Rendered()
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
