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
Dim Province_edit
Set Province_edit = New cProvince_edit
Set Page = Province_edit

' Page init processing
Call Province_edit.Page_Init()

' Page main processing
Call Province_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Province_edit = new ew_Page("Province_edit");
// page properties
Province_edit.PageID = "edit"; // page ID
Province_edit.FormID = "fProvinceedit"; // form ID
var EW_PAGE_ID = Province_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Province_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_Prov"];
		if (elm && !ew_HasValue(elm))
			return ew_OnError(this, elm, ewLanguage.Phrase("EnterRequiredField") + " - <%= ew_JsEncode2(Province.Prov.FldCaption) %>");
		elm = fobj.elements["x" + infix + "_TaxRate"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Province.TaxRate.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_ShipRate_first"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Province.ShipRate_first.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_ShipRate_Rest"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Province.ShipRate_Rest.FldErrMsg) %>");
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
Province_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Province_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Province_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Province_edit.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Province_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Province.TableCaption %></p>
<p class="aspmaker"><a href="<%= Province.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Province_edit.ShowMessage %>
<form name="fProvinceedit" id="fProvinceedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Province_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Province">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Province.Prov.Visible Then ' Prov %>
	<tr id="r_Prov"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.Prov.FldCaption %><%= Language.Phrase("FieldRequiredIndicator") %></td>
		<td<%= Province.Prov.CellAttributes %>><span id="el_Prov">
<div<%= Province.Prov.ViewAttributes %>><%= Province.Prov.EditValue %></div>
<input type="hidden" name="x_Prov" id="x_Prov" value="<%= Server.HTMLEncode(Province.Prov.CurrentValue&"") %>">
</span><%= Province.Prov.CustomMsg %></td>
	</tr>
<% End If %>
<% If Province.Province_1.Visible Then ' Province %>
	<tr id="r_Province_1"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.Province_1.FldCaption %></td>
		<td<%= Province.Province_1.CellAttributes %>><span id="el_Province_1">
<input type="text" name="x_Province_1" id="x_Province_1" size="30" maxlength="30" value="<%= Province.Province_1.EditValue %>"<%= Province.Province_1.EditAttributes %>>
</span><%= Province.Province_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If Province.fProvince.Visible Then ' fProvince %>
	<tr id="r_fProvince"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.fProvince.FldCaption %></td>
		<td<%= Province.fProvince.CellAttributes %>><span id="el_fProvince">
<input type="text" name="x_fProvince" id="x_fProvince" size="30" maxlength="255" value="<%= Province.fProvince.EditValue %>"<%= Province.fProvince.EditAttributes %>>
</span><%= Province.fProvince.CustomMsg %></td>
	</tr>
<% End If %>
<% If Province.TaxRate.Visible Then ' TaxRate %>
	<tr id="r_TaxRate"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.TaxRate.FldCaption %></td>
		<td<%= Province.TaxRate.CellAttributes %>><span id="el_TaxRate">
<input type="text" name="x_TaxRate" id="x_TaxRate" size="30" value="<%= Province.TaxRate.EditValue %>"<%= Province.TaxRate.EditAttributes %>>
</span><%= Province.TaxRate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Province.ShipRate_first.Visible Then ' ShipRate_first %>
	<tr id="r_ShipRate_first"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.ShipRate_first.FldCaption %></td>
		<td<%= Province.ShipRate_first.CellAttributes %>><span id="el_ShipRate_first">
<input type="text" name="x_ShipRate_first" id="x_ShipRate_first" size="30" value="<%= Province.ShipRate_first.EditValue %>"<%= Province.ShipRate_first.EditAttributes %>>
</span><%= Province.ShipRate_first.CustomMsg %></td>
	</tr>
<% End If %>
<% If Province.ShipRate_Rest.Visible Then ' ShipRate_Rest %>
	<tr id="r_ShipRate_Rest"<%= Province.RowAttributes %>>
		<td class="ewTableHeader"><%= Province.ShipRate_Rest.FldCaption %></td>
		<td<%= Province.ShipRate_Rest.CellAttributes %>><span id="el_ShipRate_Rest">
<input type="text" name="x_ShipRate_Rest" id="x_ShipRate_Rest" size="30" value="<%= Province.ShipRate_Rest.EditValue %>"<%= Province.ShipRate_Rest.EditAttributes %>>
</span><%= Province.ShipRate_Rest.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Province_edit.ShowPageFooter()
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
Set Province_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProvince_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Province"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Province_edit"
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
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Province"

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("Prov").Count > 0 Then
			Province.Prov.QueryStringValue = Request.QueryString("Prov")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Province.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Province.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Province.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Province.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Province.Prov.CurrentValue = "" Then Call Page_Terminate("Provincelist.asp") ' Invalid key, return to list
		Select Case Province.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Provincelist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Province.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Province.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Province.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Province.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Province.ResetAttrs()
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
		If Not Province.Prov.FldIsDetailKey Then Province.Prov.FormValue = ObjForm.GetValue("x_Prov")
		If Not Province.Province_1.FldIsDetailKey Then Province.Province_1.FormValue = ObjForm.GetValue("x_Province_1")
		If Not Province.fProvince.FldIsDetailKey Then Province.fProvince.FormValue = ObjForm.GetValue("x_fProvince")
		If Not Province.TaxRate.FldIsDetailKey Then Province.TaxRate.FormValue = ObjForm.GetValue("x_TaxRate")
		If Not Province.ShipRate_first.FldIsDetailKey Then Province.ShipRate_first.FormValue = ObjForm.GetValue("x_ShipRate_first")
		If Not Province.ShipRate_Rest.FldIsDetailKey Then Province.ShipRate_Rest.FormValue = ObjForm.GetValue("x_ShipRate_Rest")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Province.Prov.CurrentValue = Province.Prov.FormValue
		Province.Province_1.CurrentValue = Province.Province_1.FormValue
		Province.fProvince.CurrentValue = Province.fProvince.FormValue
		Province.TaxRate.CurrentValue = Province.TaxRate.FormValue
		Province.ShipRate_first.CurrentValue = Province.ShipRate_first.FormValue
		Province.ShipRate_Rest.CurrentValue = Province.ShipRate_Rest.FormValue
	End Function

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

		' ----------
		'  Edit Row
		' ----------

		ElseIf Province.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' Prov
			Province.Prov.EditCustomAttributes = ""
			Province.Prov.EditValue = Province.Prov.CurrentValue
			Province.Prov.ViewCustomAttributes = ""

			' Province
			Province.Province_1.EditCustomAttributes = ""
			Province.Province_1.EditValue = ew_HtmlEncode(Province.Province_1.CurrentValue)

			' fProvince
			Province.fProvince.EditCustomAttributes = ""
			Province.fProvince.EditValue = ew_HtmlEncode(Province.fProvince.CurrentValue)

			' TaxRate
			Province.TaxRate.EditCustomAttributes = ""
			Province.TaxRate.EditValue = ew_HtmlEncode(Province.TaxRate.CurrentValue)

			' ShipRate_first
			Province.ShipRate_first.EditCustomAttributes = ""
			Province.ShipRate_first.EditValue = ew_HtmlEncode(Province.ShipRate_first.CurrentValue)

			' ShipRate_Rest
			Province.ShipRate_Rest.EditCustomAttributes = ""
			Province.ShipRate_Rest.EditValue = ew_HtmlEncode(Province.ShipRate_Rest.CurrentValue)

			' Edit refer script
			' Prov

			Province.Prov.HrefValue = ""

			' Province
			Province.Province_1.HrefValue = ""

			' fProvince
			Province.fProvince.HrefValue = ""

			' TaxRate
			Province.TaxRate.HrefValue = ""

			' ShipRate_first
			Province.ShipRate_first.HrefValue = ""

			' ShipRate_Rest
			Province.ShipRate_Rest.HrefValue = ""
		End If
		If Province.RowType = EW_ROWTYPE_ADD Or Province.RowType = EW_ROWTYPE_EDIT Or Province.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Province.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Province.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Province.Row_Rendered()
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
		If Not IsNull(Province.Prov.FormValue) And Province.Prov.FormValue&"" = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterRequiredField") & " - " & Province.Prov.FldCaption)
		End If
		If Not ew_CheckNumber(Province.TaxRate.FormValue) Then
			Call ew_AddMessage(gsFormError, Province.TaxRate.FldErrMsg)
		End If
		If Not ew_CheckNumber(Province.ShipRate_first.FormValue) Then
			Call ew_AddMessage(gsFormError, Province.ShipRate_first.FldErrMsg)
		End If
		If Not ew_CheckNumber(Province.ShipRate_Rest.FormValue) Then
			Call ew_AddMessage(gsFormError, Province.ShipRate_Rest.FldErrMsg)
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
		sFilter = Province.KeyFilter
		Province.CurrentFilter  = sFilter
		sSql = Province.SQL
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

			' Field Prov
			' Field Province

			Call Province.Province_1.SetDbValue(Rs, Province.Province_1.CurrentValue, Null, Province.Province_1.ReadOnly)

			' Field fProvince
			Call Province.fProvince.SetDbValue(Rs, Province.fProvince.CurrentValue, Null, Province.fProvince.ReadOnly)

			' Field TaxRate
			Call Province.TaxRate.SetDbValue(Rs, Province.TaxRate.CurrentValue, Null, Province.TaxRate.ReadOnly)

			' Field ShipRate_first
			Call Province.ShipRate_first.SetDbValue(Rs, Province.ShipRate_first.CurrentValue, Null, Province.ShipRate_first.ReadOnly)

			' Field ShipRate_Rest
			Call Province.ShipRate_Rest.SetDbValue(Rs, Province.ShipRate_Rest.CurrentValue, Null, Province.ShipRate_Rest.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Province.Row_Updating(RsOld, Rs)
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
				If Province.CancelMessage <> "" Then
					FailureMessage = Province.CancelMessage
					Province.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Province.Row_Updated(RsOld, RsNew)
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
