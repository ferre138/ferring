<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Discountcodesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="DiscountTypesinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Discountcodes_edit
Set Discountcodes_edit = New cDiscountcodes_edit
Set Page = Discountcodes_edit

' Page init processing
Call Discountcodes_edit.Page_Init()

' Page main processing
Call Discountcodes_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Discountcodes_edit = new ew_Page("Discountcodes_edit");
// page properties
Discountcodes_edit.PageID = "edit"; // page ID
Discountcodes_edit.FormID = "fDiscountcodesedit"; // form ID
var EW_PAGE_ID = Discountcodes_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Discountcodes_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_OrderId"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.OrderId.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Use_date"];
		if (elm && !ew_CheckDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.Use_date.FldErrMsg) %>");
		// Set up row object
		var row = {};
		row["index"] = infix;
		for (var j = 0; j < fobj.elements.length; j++) {
			var el = fobj.elements[j];
			var len = infix.length + 2;
			if (el.name.substr(0, len) == "x" + infix + "_") {
				var elname = "x_" + el.name.substr(len);
				if (ewLang.isObject(row[elname])) { // already exists
					if (ewLang.isArray(row[elname])) {
						row[elname][row[elname].length] = el; // add to array
					} else {
						row[elname] = [row[elname], el]; // convert to array
					}
				} else {
					row[elname] = el;
				}
			}
		}
		fobj.row = row;
		// Call Form Custom Validate event
		if (!this.Form_CustomValidate(fobj)) return false;
	}
	// Process detail page
	var detailpage = (fobj.detailpage) ? fobj.detailpage.value : "";
	if (detailpage != "") {
		return eval(detailpage+".ValidateForm(fobj)");
	}
	return true;
}
// extend page with Form_CustomValidate function
Discountcodes_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Discountcodes_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Discountcodes_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Discountcodes_edit.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-1">
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Discountcodes_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Discountcodes.TableCaption %></p>
<p class="aspmaker"><a href="<%= Discountcodes.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Discountcodes_edit.ShowMessage %>
<form name="fDiscountcodesedit" id="fDiscountcodesedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Discountcodes_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Discountcodes">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Discountcodes.Discountid.Visible Then ' Discountid %>
	<tr id="r_Discountid"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Discountid.FldCaption %></td>
		<td<%= Discountcodes.Discountid.CellAttributes %>><span id="el_Discountid">
<div<%= Discountcodes.Discountid.ViewAttributes %>><%= Discountcodes.Discountid.EditValue %></div>
<input type="hidden" name="x_Discountid" id="x_Discountid" value="<%= Server.HTMLEncode(Discountcodes.Discountid.CurrentValue&"") %>">
</span><%= Discountcodes.Discountid.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
	<tr id="r_DiscountCode"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountCode.FldCaption %></td>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>><span id="el_DiscountCode">
<input type="text" name="x_DiscountCode" id="x_DiscountCode" size="30" maxlength="6" value="<%= Discountcodes.DiscountCode.EditValue %>"<%= Discountcodes.DiscountCode.EditAttributes %>>
</span><%= Discountcodes.DiscountCode.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Active.FldCaption %></td>
		<td<%= Discountcodes.Active.CellAttributes %>><span id="el_Active">
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Active" id="x_Active" value="1"<%= selwrk %><%= Discountcodes.Active.EditAttributes %>>
</span><%= Discountcodes.Active.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.used.Visible Then ' used %>
	<tr id="r_used"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.used.FldCaption %></td>
		<td<%= Discountcodes.used.CellAttributes %>><span id="el_used">
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.used.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_used" id="x_used" value="1"<%= selwrk %><%= Discountcodes.used.EditAttributes %>>
</span><%= Discountcodes.used.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.OrderId.Visible Then ' OrderId %>
	<tr id="r_OrderId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.OrderId.FldCaption %></td>
		<td<%= Discountcodes.OrderId.CellAttributes %>><span id="el_OrderId">
<input type="text" name="x_OrderId" id="x_OrderId" size="30" value="<%= Discountcodes.OrderId.EditValue %>"<%= Discountcodes.OrderId.EditAttributes %>>
</span><%= Discountcodes.OrderId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.Use_date.Visible Then ' Use_date %>
	<tr id="r_Use_date"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Use_date.FldCaption %></td>
		<td<%= Discountcodes.Use_date.CellAttributes %>><span id="el_Use_date">
<input type="text" name="x_Use_date" id="x_Use_date" value="<%= Discountcodes.Use_date.EditValue %>"<%= Discountcodes.Use_date.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Use_date" name="cal_x_Use_date" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Use_date", // input field id
	ifFormat: "%Y/%m/%d", // date format
	button: "cal_x_Use_date" // button id
});
</script>
</span><%= Discountcodes.Use_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<tr id="r_DiscountTypeId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountTypeId.FldCaption %></td>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>><span id="el_DiscountTypeId">
<% If Discountcodes.DiscountTypeId.SessionValue <> "" Then %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ViewValue %></div>
<input type="hidden" id="x_DiscountTypeId" name="x_DiscountTypeId" value="<%= ew_HtmlEncode(Discountcodes.DiscountTypeId.CurrentValue) %>">
<% Else %>
<select id="x_DiscountTypeId" name="x_DiscountTypeId"<%= Discountcodes.DiscountTypeId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Discountcodes.DiscountTypeId.EditValue) Then
	arwrk = Discountcodes.DiscountTypeId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Discountcodes.DiscountTypeId.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
</option>
<%
	Next
End If
%>
</select>
<% End If %>
</span><%= Discountcodes.DiscountTypeId.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Discountcodes_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Discountcodes_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountcodes_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Discountcodes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Discountcodes_edit"
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
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(DiscountTypes) Then Set DiscountTypes = New cDiscountTypes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Discountcodes"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
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

	' Create form object
	Set ObjForm = New cFormObj

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("Discountid").Count > 0 Then
			Discountcodes.Discountid.QueryStringValue = Request.QueryString("Discountid")
		End If

		' Set up master detail parameters
		SetUpMasterParms()
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Discountcodes.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Discountcodes.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Discountcodes.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Discountcodes.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Discountcodes.Discountid.CurrentValue = "" Then Call Page_Terminate("Discountcodeslist.asp") ' Invalid key, return to list
		Select Case Discountcodes.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Discountcodeslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Discountcodes.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Discountcodes.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Discountcodes.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Discountcodes.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Discountcodes.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
		Dim index, confirmPage
		index = ObjForm.Index ' Save form index
		ObjForm.Index = 0
		confirmPage = (ObjForm.GetValue("a_confirm") & "" <> "")
		ObjForm.Index = index ' Restore form index
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not Discountcodes.Discountid.FldIsDetailKey Then Discountcodes.Discountid.FormValue = ObjForm.GetValue("x_Discountid")
		If Not Discountcodes.DiscountCode.FldIsDetailKey Then Discountcodes.DiscountCode.FormValue = ObjForm.GetValue("x_DiscountCode")
		If Not Discountcodes.Active.FldIsDetailKey Then Discountcodes.Active.FormValue = ObjForm.GetValue("x_Active")
		If Not Discountcodes.used.FldIsDetailKey Then Discountcodes.used.FormValue = ObjForm.GetValue("x_used")
		If Not Discountcodes.OrderId.FldIsDetailKey Then Discountcodes.OrderId.FormValue = ObjForm.GetValue("x_OrderId")
		If Not Discountcodes.Use_date.FldIsDetailKey Then Discountcodes.Use_date.FormValue = ObjForm.GetValue("x_Use_date")
		If Not Discountcodes.Use_date.FldIsDetailKey Then Discountcodes.Use_date.CurrentValue = ew_UnFormatDateTime(Discountcodes.Use_date.CurrentValue, 8)
		If Not Discountcodes.DiscountTypeId.FldIsDetailKey Then Discountcodes.DiscountTypeId.FormValue = ObjForm.GetValue("x_DiscountTypeId")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Discountcodes.Discountid.CurrentValue = Discountcodes.Discountid.FormValue
		Discountcodes.DiscountCode.CurrentValue = Discountcodes.DiscountCode.FormValue
		Discountcodes.Active.CurrentValue = Discountcodes.Active.FormValue
		Discountcodes.used.CurrentValue = Discountcodes.used.FormValue
		Discountcodes.OrderId.CurrentValue = Discountcodes.OrderId.FormValue
		Discountcodes.Use_date.CurrentValue = Discountcodes.Use_date.FormValue
		Discountcodes.Use_date.CurrentValue = ew_UnFormatDateTime(Discountcodes.Use_date.CurrentValue, 8)
		Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.FormValue
	End Function

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
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.ViewValue<>"", Discountcodes.OrderId.ViewValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If
			Discountcodes.OrderId.TooltipValue = ""

			' Use_date
			Discountcodes.Use_date.LinkCustomAttributes = ""
			Discountcodes.Use_date.HrefValue = ""
			Discountcodes.Use_date.TooltipValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.LinkCustomAttributes = ""
			Discountcodes.DiscountTypeId.HrefValue = ""
			Discountcodes.DiscountTypeId.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Discountid
			Discountcodes.Discountid.EditCustomAttributes = ""
			Discountcodes.Discountid.EditValue = Discountcodes.Discountid.CurrentValue
			Discountcodes.Discountid.ViewCustomAttributes = ""

			' DiscountCode
			Discountcodes.DiscountCode.EditCustomAttributes = ""
			Discountcodes.DiscountCode.EditValue = ew_HtmlEncode(Discountcodes.DiscountCode.CurrentValue)

			' Active
			Discountcodes.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			Discountcodes.Active.EditValue = arwrk

			' used
			Discountcodes.used.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			Discountcodes.used.EditValue = arwrk

			' OrderId
			Discountcodes.OrderId.EditCustomAttributes = ""
			Discountcodes.OrderId.EditValue = ew_HtmlEncode(Discountcodes.OrderId.CurrentValue)

			' Use_date
			Discountcodes.Use_date.EditCustomAttributes = ""
			Discountcodes.Use_date.EditValue = Discountcodes.Use_date.CurrentValue

			' DiscountTypeId
			Discountcodes.DiscountTypeId.EditCustomAttributes = ""
			If Discountcodes.DiscountTypeId.SessionValue <> "" Then
				Discountcodes.DiscountTypeId.CurrentValue = Discountcodes.DiscountTypeId.SessionValue
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
			Else
				sFilterWrk = ""
			sSqlWrk = "SELECT [DiscountTypeId], [DiscountType] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
			Set RsWrk = Server.CreateObject("ADODB.Recordset")
			RsWrk.Open sSqlWrk, Conn
			If Not RsWrk.Eof Then
				arwrk = RsWrk.GetRows
			Else
				arwrk = ""
			End If
			RsWrk.Close
			Set RsWrk = Nothing
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect")))
			Discountcodes.DiscountTypeId.EditValue = arwrk
			End If

			' Edit refer script
			' Discountid

			Discountcodes.Discountid.HrefValue = ""

			' DiscountCode
			Discountcodes.DiscountCode.HrefValue = ""

			' Active
			Discountcodes.Active.HrefValue = ""

			' used
			Discountcodes.used.HrefValue = ""

			' OrderId
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.EditValue<>"", Discountcodes.OrderId.EditValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If

			' Use_date
			Discountcodes.Use_date.HrefValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.HrefValue = ""
		End If
		If Discountcodes.RowType = EW_ROWTYPE_ADD Or Discountcodes.RowType = EW_ROWTYPE_EDIT Or Discountcodes.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Discountcodes.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Discountcodes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Discountcodes.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(Discountcodes.OrderId.FormValue) Then
			Call ew_AddMessage(gsFormError, Discountcodes.OrderId.FldErrMsg)
		End If
		If Not ew_CheckDate(Discountcodes.Use_date.FormValue) Then
			Call ew_AddMessage(gsFormError, Discountcodes.Use_date.FldErrMsg)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = Discountcodes.KeyFilter
		Discountcodes.CurrentFilter  = sFilter
		sSql = Discountcodes.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field DiscountCode
			Call Discountcodes.DiscountCode.SetDbValue(Rs, Discountcodes.DiscountCode.CurrentValue, Null, Discountcodes.DiscountCode.ReadOnly)

			' Field Active
			boolwrk = Discountcodes.Active.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Discountcodes.Active.SetDbValue(Rs, boolwrk, Null, Discountcodes.Active.ReadOnly)

			' Field used
			boolwrk = Discountcodes.used.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call Discountcodes.used.SetDbValue(Rs, boolwrk, Null, Discountcodes.used.ReadOnly)

			' Field OrderId
			Call Discountcodes.OrderId.SetDbValue(Rs, Discountcodes.OrderId.CurrentValue, Null, Discountcodes.OrderId.ReadOnly)

			' Field Use_date
			Call Discountcodes.Use_date.SetDbValue(Rs, Discountcodes.Use_date.CurrentValue, Null, Discountcodes.Use_date.ReadOnly)

			' Field DiscountTypeId
			Call Discountcodes.DiscountTypeId.SetDbValue(Rs, Discountcodes.DiscountTypeId.CurrentValue, Null, Discountcodes.DiscountTypeId.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Discountcodes.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				Rs.Update
				If Err.Number <> 0 Then
					FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
			Else
				Rs.CancelUpdate
				If Discountcodes.CancelMessage <> "" Then
					FailureMessage = Discountcodes.CancelMessage
					Discountcodes.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Discountcodes.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' -----------------------------------------------------------------
	' Set up Master Detail based on querystring parameter
	'
	Sub SetUpMasterParms()
		Dim bValidMaster, sMasterTblVar
		bValidMaster = False

		' Get the keys for master table
		If Request.QueryString(EW_TABLE_SHOW_MASTER).Count > 0 Then
			sMasterTblVar = Request.QueryString(EW_TABLE_SHOW_MASTER)
			If sMasterTblVar = "" Then
				bValidMaster = True
				DbMasterFilter = ""
				DbDetailFilter = ""
			End If
			If sMasterTblVar = "DiscountTypes" Then
				bValidMaster = True
				If Request.QueryString("DiscountTypeId").Count > 0 Then
					DiscountTypes.DiscountTypeId.QueryStringValue = Request.QueryString("DiscountTypeId")
					Discountcodes.DiscountTypeId.QueryStringValue = DiscountTypes.DiscountTypeId.QueryStringValue
					Discountcodes.DiscountTypeId.SessionValue = Discountcodes.DiscountTypeId.QueryStringValue
					If Not IsNumeric(DiscountTypes.DiscountTypeId.QueryStringValue) Then bValidMaster = False
				Else
					bValidMaster = False
				End If
			End If
		End If
		If bValidMaster Then

			' Save current master table
			Discountcodes.CurrentMasterTable = sMasterTblVar

			' Reset start record counter (new master key)
			StartRec = 1
			Discountcodes.StartRecordNumber = StartRec

			' Clear previous master session values
			If sMasterTblVar <> "DiscountTypes" Then
				If Discountcodes.DiscountTypeId.QueryStringValue = "" Then Discountcodes.DiscountTypeId.SessionValue = ""
			End If
		End If
		DbMasterFilter = Discountcodes.MasterFilter '  Get master filter
		DbDetailFilter = Discountcodes.DetailFilter ' Get detail filter
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

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
