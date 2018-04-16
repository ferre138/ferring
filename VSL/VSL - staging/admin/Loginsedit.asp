<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Logins_edit
Set Logins_edit = New cLogins_edit
Set Page = Logins_edit

' Page init processing
Call Logins_edit.Page_Init()

' Page main processing
Call Logins_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Logins_edit = new ew_Page("Logins_edit");
// page properties
Logins_edit.PageID = "edit"; // page ID
Logins_edit.FormID = "fLoginsedit"; // form ID
var EW_PAGE_ID = Logins_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Logins_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
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
Logins_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Logins_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Logins_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Logins_edit.ValidateRequired = false; // no JavaScript validation
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
<% Logins_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Logins.TableCaption %></p>
<p class="aspmaker"><a href="<%= Logins.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Logins_edit.ShowMessage %>
<form name="fLoginsedit" id="fLoginsedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Logins_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Logins">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Logins.zUserId.Visible Then ' UserId %>
	<tr id="r_zUserId"<%= Logins.RowAttributes %>>
		<td class="ewTableHeader"><%= Logins.zUserId.FldCaption %></td>
		<td<%= Logins.zUserId.CellAttributes %>><span id="el_zUserId">
<div<%= Logins.zUserId.ViewAttributes %>><%= Logins.zUserId.EditValue %></div>
<input type="hidden" name="x_zUserId" id="x_zUserId" value="<%= Server.HTMLEncode(Logins.zUserId.CurrentValue&"") %>">
</span><%= Logins.zUserId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Logins.Loginname.Visible Then ' Loginname %>
	<tr id="r_Loginname"<%= Logins.RowAttributes %>>
		<td class="ewTableHeader"><%= Logins.Loginname.FldCaption %></td>
		<td<%= Logins.Loginname.CellAttributes %>><span id="el_Loginname">
<input type="text" name="x_Loginname" id="x_Loginname" size="30" maxlength="15" value="<%= Logins.Loginname.EditValue %>"<%= Logins.Loginname.EditAttributes %>>
</span><%= Logins.Loginname.CustomMsg %></td>
	</tr>
<% End If %>
<% If Logins.Loginpass.Visible Then ' Loginpass %>
	<tr id="r_Loginpass"<%= Logins.RowAttributes %>>
		<td class="ewTableHeader"><%= Logins.Loginpass.FldCaption %></td>
		<td<%= Logins.Loginpass.CellAttributes %>><span id="el_Loginpass">
<input type="text" name="x_Loginpass" id="x_Loginpass" size="30" maxlength="15" value="<%= Logins.Loginpass.EditValue %>"<%= Logins.Loginpass.EditAttributes %>>
</span><%= Logins.Loginpass.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Logins_edit.ShowPageFooter()
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
Set Logins_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cLogins_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Logins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Logins_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Logins.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Logins.TableVar & "&" ' add page token
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
		If Logins.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Logins.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Logins.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Logins) Then Set Logins = New cLogins
		Set Table = Logins

		' Initialize urls
		' Initialize form object

		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Logins"

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
		Set Logins = Nothing
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
		If Request.QueryString("zUserId").Count > 0 Then
			Logins.zUserId.QueryStringValue = Request.QueryString("zUserId")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Logins.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Logins.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Logins.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Logins.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Logins.zUserId.CurrentValue = "" Then Call Page_Terminate("Loginslist.asp") ' Invalid key, return to list
		Select Case Logins.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Loginslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Logins.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Logins.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Logins.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Logins.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Logins.ResetAttrs()
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
		If Not Logins.zUserId.FldIsDetailKey Then Logins.zUserId.FormValue = ObjForm.GetValue("x_zUserId")
		If Not Logins.Loginname.FldIsDetailKey Then Logins.Loginname.FormValue = ObjForm.GetValue("x_Loginname")
		If Not Logins.Loginpass.FldIsDetailKey Then Logins.Loginpass.FormValue = ObjForm.GetValue("x_Loginpass")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Logins.zUserId.CurrentValue = Logins.zUserId.FormValue
		Logins.Loginname.CurrentValue = Logins.Loginname.FormValue
		Logins.Loginpass.CurrentValue = Logins.Loginpass.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Logins.KeyFilter

		' Call Row Selecting event
		Call Logins.Row_Selecting(sFilter)

		' Load sql based on filter
		Logins.CurrentFilter = sFilter
		sSql = Logins.SQL
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
		Call Logins.Row_Selected(RsRow)
		Logins.zUserId.DbValue = RsRow("UserId")
		Logins.Loginname.DbValue = RsRow("Loginname")
		Logins.Loginpass.DbValue = RsRow("Loginpass")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Logins.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' UserId
		' Loginname
		' Loginpass
		' -----------
		'  View  Row
		' -----------

		If Logins.RowType = EW_ROWTYPE_VIEW Then ' View row

			' UserId
			Logins.zUserId.ViewValue = Logins.zUserId.CurrentValue
			Logins.zUserId.ViewCustomAttributes = ""

			' Loginname
			Logins.Loginname.ViewValue = Logins.Loginname.CurrentValue
			Logins.Loginname.ViewCustomAttributes = ""

			' Loginpass
			Logins.Loginpass.ViewValue = Logins.Loginpass.CurrentValue
			Logins.Loginpass.ViewCustomAttributes = ""

			' View refer script
			' UserId

			Logins.zUserId.LinkCustomAttributes = ""
			Logins.zUserId.HrefValue = ""
			Logins.zUserId.TooltipValue = ""

			' Loginname
			Logins.Loginname.LinkCustomAttributes = ""
			Logins.Loginname.HrefValue = ""
			Logins.Loginname.TooltipValue = ""

			' Loginpass
			Logins.Loginpass.LinkCustomAttributes = ""
			Logins.Loginpass.HrefValue = ""
			Logins.Loginpass.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Logins.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' UserId
			Logins.zUserId.EditCustomAttributes = ""
			Logins.zUserId.EditValue = Logins.zUserId.CurrentValue
			Logins.zUserId.ViewCustomAttributes = ""

			' Loginname
			Logins.Loginname.EditCustomAttributes = ""
			Logins.Loginname.EditValue = ew_HtmlEncode(Logins.Loginname.CurrentValue)

			' Loginpass
			Logins.Loginpass.EditCustomAttributes = ""
			Logins.Loginpass.EditValue = ew_HtmlEncode(Logins.Loginpass.CurrentValue)

			' Edit refer script
			' UserId

			Logins.zUserId.HrefValue = ""

			' Loginname
			Logins.Loginname.HrefValue = ""

			' Loginpass
			Logins.Loginpass.HrefValue = ""
		End If
		If Logins.RowType = EW_ROWTYPE_ADD Or Logins.RowType = EW_ROWTYPE_EDIT Or Logins.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Logins.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Logins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Logins.Row_Rendered()
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
		sFilter = Logins.KeyFilter
		Logins.CurrentFilter  = sFilter
		sSql = Logins.SQL
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

			' Field Loginname
			Call Logins.Loginname.SetDbValue(Rs, Logins.Loginname.CurrentValue, Null, Logins.Loginname.ReadOnly)

			' Field Loginpass
			If Not EW_CASE_SENSITIVE_PASSWORD And Not IsNull(Logins.Loginpass.CurrentValue) Then Logins.Loginpass.CurrentValue = LCase(Logins.Loginpass.CurrentValue)
			If EW_ENCRYPTED_PASSWORD And Not IsNull(Logins.Loginpass.CurrentValue) And (Logins.Loginpass.CurrentValue <> Rs("Loginpass")) Then Logins.Loginpass.CurrentValue = MD5(Logins.Loginpass.CurrentValue)
			Call Logins.Loginpass.SetDbValue(Rs, Logins.Loginpass.CurrentValue, Null, Logins.Loginpass.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Logins.Row_Updating(RsOld, Rs)
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
				If Logins.CancelMessage <> "" Then
					FailureMessage = Logins.CancelMessage
					Logins.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Logins.Row_Updated(RsOld, RsNew)
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
