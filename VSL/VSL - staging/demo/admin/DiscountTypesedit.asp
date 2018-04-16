<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="DiscountTypesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="Discountcodesinfo.asp"-->
<!--#include file="Discountcodesgridcls.asp" -->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim DiscountTypes_edit
Set DiscountTypes_edit = New cDiscountTypes_edit
Set Page = DiscountTypes_edit

' Page init processing
Call DiscountTypes_edit.Page_Init()

' Page main processing
Call DiscountTypes_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var DiscountTypes_edit = new ew_Page("DiscountTypes_edit");
// page properties
DiscountTypes_edit.PageID = "edit"; // page ID
DiscountTypes_edit.FormID = "fDiscountTypesedit"; // form ID
var EW_PAGE_ID = DiscountTypes_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
DiscountTypes_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_FreePerQty"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.FreePerQty.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_SpecialPrice"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.SpecialPrice.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_StartDate"];
		if (elm && !ew_CheckUSDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.StartDate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_EndDate"];
		if (elm && !ew_CheckUSDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.EndDate.FldErrMsg) %>");
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
DiscountTypes_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
DiscountTypes_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
DiscountTypes_edit.ValidateRequired = false; // no JavaScript validation
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
<% DiscountTypes_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= DiscountTypes.TableCaption %></p>
<p class="aspmaker"><a href="<%= DiscountTypes.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% DiscountTypes_edit.ShowMessage %>
<form name="fDiscountTypesedit" id="fDiscountTypesedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return DiscountTypes_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="DiscountTypes">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If DiscountTypes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<tr id="r_DiscountTypeId"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountTypeId.FldCaption %></td>
		<td<%= DiscountTypes.DiscountTypeId.CellAttributes %>><span id="el_DiscountTypeId">
<div<%= DiscountTypes.DiscountTypeId.ViewAttributes %>><%= DiscountTypes.DiscountTypeId.EditValue %></div>
<input type="hidden" name="x_DiscountTypeId" id="x_DiscountTypeId" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTypeId.CurrentValue&"") %>">
</span><%= DiscountTypes.DiscountTypeId.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
	<tr id="r_DiscountType"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountType.FldCaption %></td>
		<td<%= DiscountTypes.DiscountType.CellAttributes %>><span id="el_DiscountType">
<input type="text" name="x_DiscountType" id="x_DiscountType" size="30" maxlength="255" value="<%= DiscountTypes.DiscountType.EditValue %>"<%= DiscountTypes.DiscountType.EditAttributes %>>
</span><%= DiscountTypes.DiscountType.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
	<tr id="r_DiscountTitle"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.DiscountTitle.FldCaption %></td>
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>><span id="el_DiscountTitle">
<input type="text" name="x_DiscountTitle" id="x_DiscountTitle" size="30" maxlength="255" value="<%= DiscountTypes.DiscountTitle.EditValue %>"<%= DiscountTypes.DiscountTitle.EditAttributes %>>
</span><%= DiscountTypes.DiscountTitle.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
	<tr id="r_freeShipping"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.freeShipping.FldCaption %></td>
		<td<%= DiscountTypes.freeShipping.CellAttributes %>><span id="el_freeShipping">
<% selwrk = ew_IIf(ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_freeShipping" id="x_freeShipping" value="1"<%= selwrk %><%= DiscountTypes.freeShipping.EditAttributes %>>
</span><%= DiscountTypes.freeShipping.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
	<tr id="r_FreePerQty"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.FreePerQty.FldCaption %></td>
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>><span id="el_FreePerQty">
<input type="text" name="x_FreePerQty" id="x_FreePerQty" size="30" value="<%= DiscountTypes.FreePerQty.EditValue %>"<%= DiscountTypes.FreePerQty.EditAttributes %>>
</span><%= DiscountTypes.FreePerQty.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
	<tr id="r_SpecialPrice"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.SpecialPrice.FldCaption %></td>
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>><span id="el_SpecialPrice">
<input type="text" name="x_SpecialPrice" id="x_SpecialPrice" size="30" value="<%= DiscountTypes.SpecialPrice.EditValue %>"<%= DiscountTypes.SpecialPrice.EditAttributes %>>
</span><%= DiscountTypes.SpecialPrice.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.fDiscountTitle.Visible Then ' fDiscountTitle %>
	<tr id="r_fDiscountTitle"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.fDiscountTitle.FldCaption %></td>
		<td<%= DiscountTypes.fDiscountTitle.CellAttributes %>><span id="el_fDiscountTitle">
<input type="text" name="x_fDiscountTitle" id="x_fDiscountTitle" size="30" maxlength="255" value="<%= DiscountTypes.fDiscountTitle.EditValue %>"<%= DiscountTypes.fDiscountTitle.EditAttributes %>>
</span><%= DiscountTypes.fDiscountTitle.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.StartDate.Visible Then ' StartDate %>
	<tr id="r_StartDate"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.StartDate.FldCaption %></td>
		<td<%= DiscountTypes.StartDate.CellAttributes %>><span id="el_StartDate">
<input type="text" name="x_StartDate" id="x_StartDate" value="<%= DiscountTypes.StartDate.EditValue %>"<%= DiscountTypes.StartDate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_StartDate" name="cal_x_StartDate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_StartDate", // input field id
	ifFormat: "%m/%d/%Y", // date format
	button: "cal_x_StartDate" // button id
});
</script>
</span><%= DiscountTypes.StartDate.CustomMsg %></td>
	</tr>
<% End If %>
<% If DiscountTypes.EndDate.Visible Then ' EndDate %>
	<tr id="r_EndDate"<%= DiscountTypes.RowAttributes %>>
		<td class="ewTableHeader"><%= DiscountTypes.EndDate.FldCaption %></td>
		<td<%= DiscountTypes.EndDate.CellAttributes %>><span id="el_EndDate">
<input type="text" name="x_EndDate" id="x_EndDate" value="<%= DiscountTypes.EndDate.EditValue %>"<%= DiscountTypes.EndDate.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_EndDate" name="cal_x_EndDate" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_EndDate", // input field id
	ifFormat: "%m/%d/%Y", // date format
	button: "cal_x_EndDate" // button id
});
</script>
</span><%= DiscountTypes.EndDate.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<% If DiscountTypes.CurrentDetailTable = "Discountcodes" And Discountcodes.DetailEdit Then %>
<br>
<!--#include file="Discountcodesgrid.asp" -->
<br>
<% End If %>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
DiscountTypes_edit.ShowPageFooter()
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
Set DiscountTypes_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountTypes_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "DiscountTypes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "DiscountTypes_edit"
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
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "DiscountTypes"

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("DiscountTypeId").Count > 0 Then
			DiscountTypes.DiscountTypeId.QueryStringValue = Request.QueryString("DiscountTypeId")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			DiscountTypes.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Set up detail parameters
			SetUpDetailParms()

			' Validate Form
			If Not ValidateForm() Then
				DiscountTypes.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				DiscountTypes.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			DiscountTypes.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If DiscountTypes.DiscountTypeId.CurrentValue = "" Then Call Page_Terminate("DiscountTypeslist.asp") ' Invalid key, return to list

		' Set up detail parameters
		SetUpDetailParms()
		Select Case DiscountTypes.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("DiscountTypeslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				DiscountTypes.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					If DiscountTypes.CurrentDetailTable <> "" Then ' Master/detail edit
						sReturnUrl = DiscountTypes.DetailUrl
					Else
						sReturnUrl = DiscountTypes.ReturnUrl
					End If
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					DiscountTypes.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		DiscountTypes.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call DiscountTypes.ResetAttrs()
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
		If Not DiscountTypes.DiscountTypeId.FldIsDetailKey Then DiscountTypes.DiscountTypeId.FormValue = ObjForm.GetValue("x_DiscountTypeId")
		If Not DiscountTypes.DiscountType.FldIsDetailKey Then DiscountTypes.DiscountType.FormValue = ObjForm.GetValue("x_DiscountType")
		If Not DiscountTypes.DiscountTitle.FldIsDetailKey Then DiscountTypes.DiscountTitle.FormValue = ObjForm.GetValue("x_DiscountTitle")
		If Not DiscountTypes.freeShipping.FldIsDetailKey Then DiscountTypes.freeShipping.FormValue = ObjForm.GetValue("x_freeShipping")
		If Not DiscountTypes.FreePerQty.FldIsDetailKey Then DiscountTypes.FreePerQty.FormValue = ObjForm.GetValue("x_FreePerQty")
		If Not DiscountTypes.SpecialPrice.FldIsDetailKey Then DiscountTypes.SpecialPrice.FormValue = ObjForm.GetValue("x_SpecialPrice")
		If Not DiscountTypes.fDiscountTitle.FldIsDetailKey Then DiscountTypes.fDiscountTitle.FormValue = ObjForm.GetValue("x_fDiscountTitle")
		If Not DiscountTypes.StartDate.FldIsDetailKey Then DiscountTypes.StartDate.FormValue = ObjForm.GetValue("x_StartDate")
		If Not DiscountTypes.StartDate.FldIsDetailKey Then DiscountTypes.StartDate.CurrentValue = ew_UnFormatDateTime(DiscountTypes.StartDate.CurrentValue, 8)
		If Not DiscountTypes.EndDate.FldIsDetailKey Then DiscountTypes.EndDate.FormValue = ObjForm.GetValue("x_EndDate")
		If Not DiscountTypes.EndDate.FldIsDetailKey Then DiscountTypes.EndDate.CurrentValue = ew_UnFormatDateTime(DiscountTypes.EndDate.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		DiscountTypes.DiscountTypeId.CurrentValue = DiscountTypes.DiscountTypeId.FormValue
		DiscountTypes.DiscountType.CurrentValue = DiscountTypes.DiscountType.FormValue
		DiscountTypes.DiscountTitle.CurrentValue = DiscountTypes.DiscountTitle.FormValue
		DiscountTypes.freeShipping.CurrentValue = DiscountTypes.freeShipping.FormValue
		DiscountTypes.FreePerQty.CurrentValue = DiscountTypes.FreePerQty.FormValue
		DiscountTypes.SpecialPrice.CurrentValue = DiscountTypes.SpecialPrice.FormValue
		DiscountTypes.fDiscountTitle.CurrentValue = DiscountTypes.fDiscountTitle.FormValue
		DiscountTypes.StartDate.CurrentValue = DiscountTypes.StartDate.FormValue
		DiscountTypes.StartDate.CurrentValue = ew_UnFormatDateTime(DiscountTypes.StartDate.CurrentValue, 8)
		DiscountTypes.EndDate.CurrentValue = DiscountTypes.EndDate.FormValue
		DiscountTypes.EndDate.CurrentValue = ew_UnFormatDateTime(DiscountTypes.EndDate.CurrentValue, 8)
	End Function

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
		DiscountTypes.fDiscountTitle.DbValue = RsRow("fDiscountTitle")
		DiscountTypes.StartDate.DbValue = RsRow("StartDate")
		DiscountTypes.EndDate.DbValue = RsRow("EndDate")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
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
		' fDiscountTitle
		' StartDate
		' EndDate
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

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.ViewValue = DiscountTypes.fDiscountTitle.CurrentValue
			DiscountTypes.fDiscountTitle.ViewCustomAttributes = ""

			' StartDate
			DiscountTypes.StartDate.ViewValue = DiscountTypes.StartDate.CurrentValue
			DiscountTypes.StartDate.ViewCustomAttributes = ""

			' EndDate
			DiscountTypes.EndDate.ViewValue = DiscountTypes.EndDate.CurrentValue
			DiscountTypes.EndDate.ViewCustomAttributes = ""

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

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.LinkCustomAttributes = ""
			DiscountTypes.fDiscountTitle.HrefValue = ""
			DiscountTypes.fDiscountTitle.TooltipValue = ""

			' StartDate
			DiscountTypes.StartDate.LinkCustomAttributes = ""
			DiscountTypes.StartDate.HrefValue = ""
			DiscountTypes.StartDate.TooltipValue = ""

			' EndDate
			DiscountTypes.EndDate.LinkCustomAttributes = ""
			DiscountTypes.EndDate.HrefValue = ""
			DiscountTypes.EndDate.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' DiscountTypeId
			DiscountTypes.DiscountTypeId.EditCustomAttributes = ""
			DiscountTypes.DiscountTypeId.EditValue = DiscountTypes.DiscountTypeId.CurrentValue
			DiscountTypes.DiscountTypeId.ViewCustomAttributes = ""

			' DiscountType
			DiscountTypes.DiscountType.EditCustomAttributes = ""
			DiscountTypes.DiscountType.EditValue = ew_HtmlEncode(DiscountTypes.DiscountType.CurrentValue)

			' DiscountTitle
			DiscountTypes.DiscountTitle.EditCustomAttributes = ""
			DiscountTypes.DiscountTitle.EditValue = ew_HtmlEncode(DiscountTypes.DiscountTitle.CurrentValue)

			' freeShipping
			DiscountTypes.freeShipping.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(1) <> "", DiscountTypes.freeShipping.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(2) <> "", DiscountTypes.freeShipping.FldTagCaption(2), "No")
			DiscountTypes.freeShipping.EditValue = arwrk

			' FreePerQty
			DiscountTypes.FreePerQty.EditCustomAttributes = ""
			DiscountTypes.FreePerQty.EditValue = ew_HtmlEncode(DiscountTypes.FreePerQty.CurrentValue)

			' SpecialPrice
			DiscountTypes.SpecialPrice.EditCustomAttributes = ""
			DiscountTypes.SpecialPrice.EditValue = ew_HtmlEncode(DiscountTypes.SpecialPrice.CurrentValue)

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.EditCustomAttributes = ""
			DiscountTypes.fDiscountTitle.EditValue = ew_HtmlEncode(DiscountTypes.fDiscountTitle.CurrentValue)

			' StartDate
			DiscountTypes.StartDate.EditCustomAttributes = ""
			DiscountTypes.StartDate.EditValue = DiscountTypes.StartDate.CurrentValue

			' EndDate
			DiscountTypes.EndDate.EditCustomAttributes = ""
			DiscountTypes.EndDate.EditValue = DiscountTypes.EndDate.CurrentValue

			' Edit refer script
			' DiscountTypeId

			DiscountTypes.DiscountTypeId.HrefValue = ""

			' DiscountType
			DiscountTypes.DiscountType.HrefValue = ""

			' DiscountTitle
			DiscountTypes.DiscountTitle.HrefValue = ""

			' freeShipping
			DiscountTypes.freeShipping.HrefValue = ""

			' FreePerQty
			DiscountTypes.FreePerQty.HrefValue = ""

			' SpecialPrice
			DiscountTypes.SpecialPrice.HrefValue = ""

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.HrefValue = ""

			' StartDate
			DiscountTypes.StartDate.HrefValue = ""

			' EndDate
			DiscountTypes.EndDate.HrefValue = ""
		End If
		If DiscountTypes.RowType = EW_ROWTYPE_ADD Or DiscountTypes.RowType = EW_ROWTYPE_EDIT Or DiscountTypes.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call DiscountTypes.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If DiscountTypes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call DiscountTypes.Row_Rendered()
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
		If Not ew_CheckInteger(DiscountTypes.FreePerQty.FormValue) Then
			Call ew_AddMessage(gsFormError, DiscountTypes.FreePerQty.FldErrMsg)
		End If
		If Not ew_CheckNumber(DiscountTypes.SpecialPrice.FormValue) Then
			Call ew_AddMessage(gsFormError, DiscountTypes.SpecialPrice.FldErrMsg)
		End If
		If Not ew_CheckUSDate(DiscountTypes.StartDate.FormValue) Then
			Call ew_AddMessage(gsFormError, DiscountTypes.StartDate.FldErrMsg)
		End If
		If Not ew_CheckUSDate(DiscountTypes.EndDate.FormValue) Then
			Call ew_AddMessage(gsFormError, DiscountTypes.EndDate.FldErrMsg)
		End If

		' Validate detail grid
		If DiscountTypes.CurrentDetailTable = "Discountcodes" And Discountcodes.DetailEdit Then
			Dim Discountcodes_grid
			Set Discountcodes_grid = new cDiscountcodes_grid ' get detail page object
			Call Discountcodes_grid.ValidateGridForm()
			Set Discountcodes_grid = Nothing
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
		sFilter = DiscountTypes.KeyFilter
		DiscountTypes.CurrentFilter  = sFilter
		sSql = DiscountTypes.SQL
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

			' Begin transaction
			If DiscountTypes.CurrentDetailTable <> "" Then Conn.BeginTrans

			' Field DiscountType
			Call DiscountTypes.DiscountType.SetDbValue(Rs, DiscountTypes.DiscountType.CurrentValue, Null, DiscountTypes.DiscountType.ReadOnly)

			' Field DiscountTitle
			Call DiscountTypes.DiscountTitle.SetDbValue(Rs, DiscountTypes.DiscountTitle.CurrentValue, Null, DiscountTypes.DiscountTitle.ReadOnly)

			' Field freeShipping
			boolwrk = DiscountTypes.freeShipping.CurrentValue
			If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
			Call DiscountTypes.freeShipping.SetDbValue(Rs, boolwrk, Null, DiscountTypes.freeShipping.ReadOnly)

			' Field FreePerQty
			Call DiscountTypes.FreePerQty.SetDbValue(Rs, DiscountTypes.FreePerQty.CurrentValue, Null, DiscountTypes.FreePerQty.ReadOnly)

			' Field SpecialPrice
			Call DiscountTypes.SpecialPrice.SetDbValue(Rs, DiscountTypes.SpecialPrice.CurrentValue, Null, DiscountTypes.SpecialPrice.ReadOnly)

			' Field fDiscountTitle
			Call DiscountTypes.fDiscountTitle.SetDbValue(Rs, DiscountTypes.fDiscountTitle.CurrentValue, Null, DiscountTypes.fDiscountTitle.ReadOnly)

			' Field StartDate
			Call DiscountTypes.StartDate.SetDbValue(Rs, DiscountTypes.StartDate.CurrentValue, Null, DiscountTypes.StartDate.ReadOnly)

			' Field EndDate
			Call DiscountTypes.EndDate.SetDbValue(Rs, DiscountTypes.EndDate.CurrentValue, Null, DiscountTypes.EndDate.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = DiscountTypes.Row_Updating(RsOld, Rs)
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

				' Update detail records
				If EditRow Then
					If DiscountTypes.CurrentDetailTable = "Discountcodes" And Discountcodes.DetailEdit Then
						Dim Discountcodes_grid
						Set Discountcodes_grid = New cDiscountcodes_grid ' get detail page object
						EditRow = Discountcodes_grid.GridUpdate
						Set Discountcodes_grid = Nothing
					End If
				End If

				' Commit/Rollback transaction
				If DiscountTypes.CurrentDetailTable <> "" Then
					If EditRow Then
						Conn.CommitTrans ' Commit transaction
					Else
						Conn.RollbackTrans ' Rollback transaction
					End If
				End If
			Else
				Rs.CancelUpdate
				If DiscountTypes.CancelMessage <> "" Then
					FailureMessage = DiscountTypes.CancelMessage
					DiscountTypes.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call DiscountTypes.Row_Updated(RsOld, RsNew)
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

	' Set up detail parms based on QueryString
	Sub SetUpDetailParms()
		Dim sDetailTblVar, bValidDetail
		bValidDetail = False

		' Get the keys for master table
		If Request.QueryString(EW_TABLE_SHOW_DETAIL).Count > 0 Then
			sDetailTblVar = Request.QueryString(EW_TABLE_SHOW_DETAIL)
			DiscountTypes.CurrentDetailTable = sDetailTblVar
		Else
			sDetailTblVar = DiscountTypes.CurrentDetailTable
		End If
		If sDetailTblVar <> "" Then
			If sDetailTblVar = "Discountcodes" Then
				If IsEmpty(Discountcodes) Then
					Set Discountcodes = New cDiscountcodes
				End If
				If Discountcodes.DetailEdit Then
					Discountcodes.CurrentMode = "edit"
					Discountcodes.CurrentAction = "gridedit"

					' Save current master table to detail table
					Discountcodes.CurrentMasterTable = DiscountTypes.TableVar
					Discountcodes.StartRecordNumber = 1
					Discountcodes.DiscountTypeId.FldIsDetailKey = True
					Discountcodes.DiscountTypeId.CurrentValue = DiscountTypes.DiscountTypeId.CurrentValue
					Discountcodes.DiscountTypeId.SessionValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
			End If
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

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
