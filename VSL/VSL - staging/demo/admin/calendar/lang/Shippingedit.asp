<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Shippinginfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Shipping_edit
Set Shipping_edit = New cShipping_edit
Set Page = Shipping_edit

' Page init processing
Call Shipping_edit.Page_Init()

' Page main processing
Call Shipping_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Shipping_edit = new ew_Page("Shipping_edit");
// page properties
Shipping_edit.PageID = "edit"; // page ID
Shipping_edit.FormID = "fShippingedit"; // form ID
var EW_PAGE_ID = Shipping_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Shipping_edit.ValidateForm = function(fobj) {
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
Shipping_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Shipping_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Shipping_edit.ValidateRequired = false; // no JavaScript validation
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
<% Shipping_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Shipping.TableCaption %></p>
<p class="aspmaker"><a href="<%= Shipping.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Shipping_edit.ShowMessage %>
<form name="fShippingedit" id="fShippingedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Shipping_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Shipping">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Shipping.AddressID.Visible Then ' AddressID %>
	<tr id="r_AddressID"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.AddressID.FldCaption %></td>
		<td<%= Shipping.AddressID.CellAttributes %>><span id="el_AddressID">
<div<%= Shipping.AddressID.ViewAttributes %>><%= Shipping.AddressID.EditValue %></div>
<input type="hidden" name="x_AddressID" id="x_AddressID" value="<%= Server.HTMLEncode(Shipping.AddressID.CurrentValue&"") %>">
</span><%= Shipping.AddressID.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.CustomerId.Visible Then ' CustomerId %>
	<tr id="r_CustomerId"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.CustomerId.FldCaption %></td>
		<td<%= Shipping.CustomerId.CellAttributes %>><span id="el_CustomerId">
<select id="x_CustomerId" name="x_CustomerId"<%= Shipping.CustomerId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Shipping.CustomerId.EditValue) Then
	arwrk = Shipping.CustomerId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Shipping.CustomerId.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
<% If arwrk(2, rowcntwrk) <> "" Then %>
<%= ew_ValueSeparator(rowcntwrk,1,Shipping.CustomerId) %><%= arwrk(2, rowcntwrk) %>
<% End If %>
</option>
<%
	Next
End If
%>
</select>
</span><%= Shipping.CustomerId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_FirstName.Visible Then ' ship_FirstName %>
	<tr id="r_ship_FirstName"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_FirstName.FldCaption %></td>
		<td<%= Shipping.ship_FirstName.CellAttributes %>><span id="el_ship_FirstName">
<input type="text" name="x_ship_FirstName" id="x_ship_FirstName" size="30" maxlength="50" value="<%= Shipping.ship_FirstName.EditValue %>"<%= Shipping.ship_FirstName.EditAttributes %>>
</span><%= Shipping.ship_FirstName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_LastName.Visible Then ' ship_LastName %>
	<tr id="r_ship_LastName"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_LastName.FldCaption %></td>
		<td<%= Shipping.ship_LastName.CellAttributes %>><span id="el_ship_LastName">
<input type="text" name="x_ship_LastName" id="x_ship_LastName" size="30" maxlength="50" value="<%= Shipping.ship_LastName.EditValue %>"<%= Shipping.ship_LastName.EditAttributes %>>
</span><%= Shipping.ship_LastName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_Address.Visible Then ' ship_Address %>
	<tr id="r_ship_Address"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Address.FldCaption %></td>
		<td<%= Shipping.ship_Address.CellAttributes %>><span id="el_ship_Address">
<input type="text" name="x_ship_Address" id="x_ship_Address" size="30" maxlength="255" value="<%= Shipping.ship_Address.EditValue %>"<%= Shipping.ship_Address.EditAttributes %>>
</span><%= Shipping.ship_Address.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_City.Visible Then ' ship_City %>
	<tr id="r_ship_City"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_City.FldCaption %></td>
		<td<%= Shipping.ship_City.CellAttributes %>><span id="el_ship_City">
<input type="text" name="x_ship_City" id="x_ship_City" size="30" maxlength="50" value="<%= Shipping.ship_City.EditValue %>"<%= Shipping.ship_City.EditAttributes %>>
</span><%= Shipping.ship_City.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_Province.Visible Then ' ship_Province %>
	<tr id="r_ship_Province"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Province.FldCaption %></td>
		<td<%= Shipping.ship_Province.CellAttributes %>><span id="el_ship_Province">
<input type="text" name="x_ship_Province" id="x_ship_Province" size="30" maxlength="2" value="<%= Shipping.ship_Province.EditValue %>"<%= Shipping.ship_Province.EditAttributes %>>
</span><%= Shipping.ship_Province.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_PostalCode.Visible Then ' ship_PostalCode %>
	<tr id="r_ship_PostalCode"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_PostalCode.FldCaption %></td>
		<td<%= Shipping.ship_PostalCode.CellAttributes %>><span id="el_ship_PostalCode">
<input type="text" name="x_ship_PostalCode" id="x_ship_PostalCode" size="30" maxlength="20" value="<%= Shipping.ship_PostalCode.EditValue %>"<%= Shipping.ship_PostalCode.EditAttributes %>>
</span><%= Shipping.ship_PostalCode.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_Country.Visible Then ' ship_Country %>
	<tr id="r_ship_Country"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Country.FldCaption %></td>
		<td<%= Shipping.ship_Country.CellAttributes %>><span id="el_ship_Country">
<input type="text" name="x_ship_Country" id="x_ship_Country" size="30" maxlength="50" value="<%= Shipping.ship_Country.EditValue %>"<%= Shipping.ship_Country.EditAttributes %>>
</span><%= Shipping.ship_Country.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
	<tr id="r_ship_EmailAddress"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_EmailAddress.FldCaption %></td>
		<td<%= Shipping.ship_EmailAddress.CellAttributes %>><span id="el_ship_EmailAddress">
<input type="text" name="x_ship_EmailAddress" id="x_ship_EmailAddress" size="30" maxlength="50" value="<%= Shipping.ship_EmailAddress.EditValue %>"<%= Shipping.ship_EmailAddress.EditAttributes %>>
</span><%= Shipping.ship_EmailAddress.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.HomePhone.Visible Then ' HomePhone %>
	<tr id="r_HomePhone"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.HomePhone.FldCaption %></td>
		<td<%= Shipping.HomePhone.CellAttributes %>><span id="el_HomePhone">
<input type="text" name="x_HomePhone" id="x_HomePhone" size="30" maxlength="30" value="<%= Shipping.HomePhone.EditValue %>"<%= Shipping.HomePhone.EditAttributes %>>
</span><%= Shipping.HomePhone.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.WorkPhone.Visible Then ' WorkPhone %>
	<tr id="r_WorkPhone"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.WorkPhone.FldCaption %></td>
		<td<%= Shipping.WorkPhone.CellAttributes %>><span id="el_WorkPhone">
<input type="text" name="x_WorkPhone" id="x_WorkPhone" size="30" maxlength="30" value="<%= Shipping.WorkPhone.EditValue %>"<%= Shipping.WorkPhone.EditAttributes %>>
</span><%= Shipping.WorkPhone.CustomMsg %></td>
	</tr>
<% End If %>
<% If Shipping.ship_Address2.Visible Then ' ship_Address2 %>
	<tr id="r_ship_Address2"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Address2.FldCaption %></td>
		<td<%= Shipping.ship_Address2.CellAttributes %>><span id="el_ship_Address2">
<input type="text" name="x_ship_Address2" id="x_ship_Address2" size="30" maxlength="50" value="<%= Shipping.ship_Address2.EditValue %>"<%= Shipping.ship_Address2.EditAttributes %>>
</span><%= Shipping.ship_Address2.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Shipping_edit.ShowPageFooter()
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
Set Shipping_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cShipping_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Shipping"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Shipping_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Shipping.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Shipping.TableVar & "&" ' add page token
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
		If Shipping.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Shipping.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Shipping.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Shipping) Then Set Shipping = New cShipping
		Set Table = Shipping

		' Initialize urls
		' Initialize form object

		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Shipping"

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
		Set Shipping = Nothing
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
		If Request.QueryString("AddressID").Count > 0 Then
			Shipping.AddressID.QueryStringValue = Request.QueryString("AddressID")
		End If
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Shipping.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				Shipping.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Shipping.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Shipping.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Shipping.AddressID.CurrentValue = "" Then Call Page_Terminate("Shippinglist.asp") ' Invalid key, return to list
		Select Case Shipping.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Shippinglist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Shipping.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = Shipping.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Shipping.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Shipping.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Shipping.ResetAttrs()
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
		If Not Shipping.AddressID.FldIsDetailKey Then Shipping.AddressID.FormValue = ObjForm.GetValue("x_AddressID")
		If Not Shipping.CustomerId.FldIsDetailKey Then Shipping.CustomerId.FormValue = ObjForm.GetValue("x_CustomerId")
		If Not Shipping.ship_FirstName.FldIsDetailKey Then Shipping.ship_FirstName.FormValue = ObjForm.GetValue("x_ship_FirstName")
		If Not Shipping.ship_LastName.FldIsDetailKey Then Shipping.ship_LastName.FormValue = ObjForm.GetValue("x_ship_LastName")
		If Not Shipping.ship_Address.FldIsDetailKey Then Shipping.ship_Address.FormValue = ObjForm.GetValue("x_ship_Address")
		If Not Shipping.ship_City.FldIsDetailKey Then Shipping.ship_City.FormValue = ObjForm.GetValue("x_ship_City")
		If Not Shipping.ship_Province.FldIsDetailKey Then Shipping.ship_Province.FormValue = ObjForm.GetValue("x_ship_Province")
		If Not Shipping.ship_PostalCode.FldIsDetailKey Then Shipping.ship_PostalCode.FormValue = ObjForm.GetValue("x_ship_PostalCode")
		If Not Shipping.ship_Country.FldIsDetailKey Then Shipping.ship_Country.FormValue = ObjForm.GetValue("x_ship_Country")
		If Not Shipping.ship_EmailAddress.FldIsDetailKey Then Shipping.ship_EmailAddress.FormValue = ObjForm.GetValue("x_ship_EmailAddress")
		If Not Shipping.HomePhone.FldIsDetailKey Then Shipping.HomePhone.FormValue = ObjForm.GetValue("x_HomePhone")
		If Not Shipping.WorkPhone.FldIsDetailKey Then Shipping.WorkPhone.FormValue = ObjForm.GetValue("x_WorkPhone")
		If Not Shipping.ship_Address2.FldIsDetailKey Then Shipping.ship_Address2.FormValue = ObjForm.GetValue("x_ship_Address2")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Shipping.AddressID.CurrentValue = Shipping.AddressID.FormValue
		Shipping.CustomerId.CurrentValue = Shipping.CustomerId.FormValue
		Shipping.ship_FirstName.CurrentValue = Shipping.ship_FirstName.FormValue
		Shipping.ship_LastName.CurrentValue = Shipping.ship_LastName.FormValue
		Shipping.ship_Address.CurrentValue = Shipping.ship_Address.FormValue
		Shipping.ship_City.CurrentValue = Shipping.ship_City.FormValue
		Shipping.ship_Province.CurrentValue = Shipping.ship_Province.FormValue
		Shipping.ship_PostalCode.CurrentValue = Shipping.ship_PostalCode.FormValue
		Shipping.ship_Country.CurrentValue = Shipping.ship_Country.FormValue
		Shipping.ship_EmailAddress.CurrentValue = Shipping.ship_EmailAddress.FormValue
		Shipping.HomePhone.CurrentValue = Shipping.HomePhone.FormValue
		Shipping.WorkPhone.CurrentValue = Shipping.WorkPhone.FormValue
		Shipping.ship_Address2.CurrentValue = Shipping.ship_Address2.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Shipping.KeyFilter

		' Call Row Selecting event
		Call Shipping.Row_Selecting(sFilter)

		' Load sql based on filter
		Shipping.CurrentFilter = sFilter
		sSql = Shipping.SQL
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
		Call Shipping.Row_Selected(RsRow)
		Shipping.AddressID.DbValue = RsRow("AddressID")
		Shipping.CustomerId.DbValue = RsRow("CustomerId")
		Shipping.ship_FirstName.DbValue = RsRow("ship_FirstName")
		Shipping.ship_LastName.DbValue = RsRow("ship_LastName")
		Shipping.ship_Address.DbValue = RsRow("ship_Address")
		Shipping.ship_City.DbValue = RsRow("ship_City")
		Shipping.ship_Province.DbValue = RsRow("ship_Province")
		Shipping.ship_PostalCode.DbValue = RsRow("ship_PostalCode")
		Shipping.ship_Country.DbValue = RsRow("ship_Country")
		Shipping.ship_EmailAddress.DbValue = RsRow("ship_EmailAddress")
		Shipping.HomePhone.DbValue = RsRow("HomePhone")
		Shipping.WorkPhone.DbValue = RsRow("WorkPhone")
		Shipping.ship_Address2.DbValue = RsRow("ship_Address2")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Shipping.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' AddressID
		' CustomerId
		' ship_FirstName
		' ship_LastName
		' ship_Address
		' ship_City
		' ship_Province
		' ship_PostalCode
		' ship_Country
		' ship_EmailAddress
		' HomePhone
		' WorkPhone
		' ship_Address2
		' -----------
		'  View  Row
		' -----------

		If Shipping.RowType = EW_ROWTYPE_VIEW Then ' View row

			' AddressID
			Shipping.AddressID.ViewValue = Shipping.AddressID.CurrentValue
			Shipping.AddressID.ViewCustomAttributes = ""

			' CustomerId
			If Shipping.CustomerId.CurrentValue & "" <> "" Then
				sFilterWrk = "[CustomerID] = " & ew_AdjustSql(Shipping.CustomerId.CurrentValue) & ""
			sSqlWrk = "SELECT [Inv_FirstName], [Inv_LastName] FROM [Customers]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Shipping.CustomerId.ViewValue = RsWrk("Inv_FirstName")
					Shipping.CustomerId.ViewValue = Shipping.CustomerId.ViewValue & ew_ValueSeparator(0,1,Shipping.CustomerId) & RsWrk("Inv_LastName")
				Else
					Shipping.CustomerId.ViewValue = Shipping.CustomerId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Shipping.CustomerId.ViewValue = Null
			End If
			Shipping.CustomerId.ViewCustomAttributes = ""

			' ship_FirstName
			Shipping.ship_FirstName.ViewValue = Shipping.ship_FirstName.CurrentValue
			Shipping.ship_FirstName.ViewCustomAttributes = ""

			' ship_LastName
			Shipping.ship_LastName.ViewValue = Shipping.ship_LastName.CurrentValue
			Shipping.ship_LastName.ViewCustomAttributes = ""

			' ship_Address
			Shipping.ship_Address.ViewValue = Shipping.ship_Address.CurrentValue
			Shipping.ship_Address.ViewCustomAttributes = ""

			' ship_City
			Shipping.ship_City.ViewValue = Shipping.ship_City.CurrentValue
			Shipping.ship_City.ViewCustomAttributes = ""

			' ship_Province
			Shipping.ship_Province.ViewValue = Shipping.ship_Province.CurrentValue
			Shipping.ship_Province.ViewCustomAttributes = ""

			' ship_PostalCode
			Shipping.ship_PostalCode.ViewValue = Shipping.ship_PostalCode.CurrentValue
			Shipping.ship_PostalCode.ViewCustomAttributes = ""

			' ship_Country
			Shipping.ship_Country.ViewValue = Shipping.ship_Country.CurrentValue
			Shipping.ship_Country.ViewCustomAttributes = ""

			' ship_EmailAddress
			Shipping.ship_EmailAddress.ViewValue = Shipping.ship_EmailAddress.CurrentValue
			Shipping.ship_EmailAddress.ViewCustomAttributes = ""

			' HomePhone
			Shipping.HomePhone.ViewValue = Shipping.HomePhone.CurrentValue
			Shipping.HomePhone.ViewCustomAttributes = ""

			' WorkPhone
			Shipping.WorkPhone.ViewValue = Shipping.WorkPhone.CurrentValue
			Shipping.WorkPhone.ViewCustomAttributes = ""

			' ship_Address2
			Shipping.ship_Address2.ViewValue = Shipping.ship_Address2.CurrentValue
			Shipping.ship_Address2.ViewCustomAttributes = ""

			' View refer script
			' AddressID

			Shipping.AddressID.LinkCustomAttributes = ""
			Shipping.AddressID.HrefValue = ""
			Shipping.AddressID.TooltipValue = ""

			' CustomerId
			Shipping.CustomerId.LinkCustomAttributes = ""
			Shipping.CustomerId.HrefValue = ""
			Shipping.CustomerId.TooltipValue = ""

			' ship_FirstName
			Shipping.ship_FirstName.LinkCustomAttributes = ""
			Shipping.ship_FirstName.HrefValue = ""
			Shipping.ship_FirstName.TooltipValue = ""

			' ship_LastName
			Shipping.ship_LastName.LinkCustomAttributes = ""
			Shipping.ship_LastName.HrefValue = ""
			Shipping.ship_LastName.TooltipValue = ""

			' ship_Address
			Shipping.ship_Address.LinkCustomAttributes = ""
			Shipping.ship_Address.HrefValue = ""
			Shipping.ship_Address.TooltipValue = ""

			' ship_City
			Shipping.ship_City.LinkCustomAttributes = ""
			Shipping.ship_City.HrefValue = ""
			Shipping.ship_City.TooltipValue = ""

			' ship_Province
			Shipping.ship_Province.LinkCustomAttributes = ""
			Shipping.ship_Province.HrefValue = ""
			Shipping.ship_Province.TooltipValue = ""

			' ship_PostalCode
			Shipping.ship_PostalCode.LinkCustomAttributes = ""
			Shipping.ship_PostalCode.HrefValue = ""
			Shipping.ship_PostalCode.TooltipValue = ""

			' ship_Country
			Shipping.ship_Country.LinkCustomAttributes = ""
			Shipping.ship_Country.HrefValue = ""
			Shipping.ship_Country.TooltipValue = ""

			' ship_EmailAddress
			Shipping.ship_EmailAddress.LinkCustomAttributes = ""
			Shipping.ship_EmailAddress.HrefValue = ""
			Shipping.ship_EmailAddress.TooltipValue = ""

			' HomePhone
			Shipping.HomePhone.LinkCustomAttributes = ""
			Shipping.HomePhone.HrefValue = ""
			Shipping.HomePhone.TooltipValue = ""

			' WorkPhone
			Shipping.WorkPhone.LinkCustomAttributes = ""
			Shipping.WorkPhone.HrefValue = ""
			Shipping.WorkPhone.TooltipValue = ""

			' ship_Address2
			Shipping.ship_Address2.LinkCustomAttributes = ""
			Shipping.ship_Address2.HrefValue = ""
			Shipping.ship_Address2.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf Shipping.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' AddressID
			Shipping.AddressID.EditCustomAttributes = ""
			Shipping.AddressID.EditValue = Shipping.AddressID.CurrentValue
			Shipping.AddressID.ViewCustomAttributes = ""

			' CustomerId
			Shipping.CustomerId.EditCustomAttributes = ""
				sFilterWrk = ""
			sSqlWrk = "SELECT [CustomerID], [Inv_FirstName] AS [DispFld], [Inv_LastName] AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [Customers]"
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
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect"), ""))
			Shipping.CustomerId.EditValue = arwrk

			' ship_FirstName
			Shipping.ship_FirstName.EditCustomAttributes = ""
			Shipping.ship_FirstName.EditValue = ew_HtmlEncode(Shipping.ship_FirstName.CurrentValue)

			' ship_LastName
			Shipping.ship_LastName.EditCustomAttributes = ""
			Shipping.ship_LastName.EditValue = ew_HtmlEncode(Shipping.ship_LastName.CurrentValue)

			' ship_Address
			Shipping.ship_Address.EditCustomAttributes = ""
			Shipping.ship_Address.EditValue = ew_HtmlEncode(Shipping.ship_Address.CurrentValue)

			' ship_City
			Shipping.ship_City.EditCustomAttributes = ""
			Shipping.ship_City.EditValue = ew_HtmlEncode(Shipping.ship_City.CurrentValue)

			' ship_Province
			Shipping.ship_Province.EditCustomAttributes = ""
			Shipping.ship_Province.EditValue = ew_HtmlEncode(Shipping.ship_Province.CurrentValue)

			' ship_PostalCode
			Shipping.ship_PostalCode.EditCustomAttributes = ""
			Shipping.ship_PostalCode.EditValue = ew_HtmlEncode(Shipping.ship_PostalCode.CurrentValue)

			' ship_Country
			Shipping.ship_Country.EditCustomAttributes = ""
			Shipping.ship_Country.EditValue = ew_HtmlEncode(Shipping.ship_Country.CurrentValue)

			' ship_EmailAddress
			Shipping.ship_EmailAddress.EditCustomAttributes = ""
			Shipping.ship_EmailAddress.EditValue = ew_HtmlEncode(Shipping.ship_EmailAddress.CurrentValue)

			' HomePhone
			Shipping.HomePhone.EditCustomAttributes = ""
			Shipping.HomePhone.EditValue = ew_HtmlEncode(Shipping.HomePhone.CurrentValue)

			' WorkPhone
			Shipping.WorkPhone.EditCustomAttributes = ""
			Shipping.WorkPhone.EditValue = ew_HtmlEncode(Shipping.WorkPhone.CurrentValue)

			' ship_Address2
			Shipping.ship_Address2.EditCustomAttributes = ""
			Shipping.ship_Address2.EditValue = ew_HtmlEncode(Shipping.ship_Address2.CurrentValue)

			' Edit refer script
			' AddressID

			Shipping.AddressID.HrefValue = ""

			' CustomerId
			Shipping.CustomerId.HrefValue = ""

			' ship_FirstName
			Shipping.ship_FirstName.HrefValue = ""

			' ship_LastName
			Shipping.ship_LastName.HrefValue = ""

			' ship_Address
			Shipping.ship_Address.HrefValue = ""

			' ship_City
			Shipping.ship_City.HrefValue = ""

			' ship_Province
			Shipping.ship_Province.HrefValue = ""

			' ship_PostalCode
			Shipping.ship_PostalCode.HrefValue = ""

			' ship_Country
			Shipping.ship_Country.HrefValue = ""

			' ship_EmailAddress
			Shipping.ship_EmailAddress.HrefValue = ""

			' HomePhone
			Shipping.HomePhone.HrefValue = ""

			' WorkPhone
			Shipping.WorkPhone.HrefValue = ""

			' ship_Address2
			Shipping.ship_Address2.HrefValue = ""
		End If
		If Shipping.RowType = EW_ROWTYPE_ADD Or Shipping.RowType = EW_ROWTYPE_EDIT Or Shipping.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Shipping.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Shipping.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Shipping.Row_Rendered()
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
		sFilter = Shipping.KeyFilter
		Shipping.CurrentFilter  = sFilter
		sSql = Shipping.SQL
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

			' Field CustomerId
			Call Shipping.CustomerId.SetDbValue(Rs, Shipping.CustomerId.CurrentValue, Null, Shipping.CustomerId.ReadOnly)

			' Field ship_FirstName
			Call Shipping.ship_FirstName.SetDbValue(Rs, Shipping.ship_FirstName.CurrentValue, Null, Shipping.ship_FirstName.ReadOnly)

			' Field ship_LastName
			Call Shipping.ship_LastName.SetDbValue(Rs, Shipping.ship_LastName.CurrentValue, Null, Shipping.ship_LastName.ReadOnly)

			' Field ship_Address
			Call Shipping.ship_Address.SetDbValue(Rs, Shipping.ship_Address.CurrentValue, Null, Shipping.ship_Address.ReadOnly)

			' Field ship_City
			Call Shipping.ship_City.SetDbValue(Rs, Shipping.ship_City.CurrentValue, Null, Shipping.ship_City.ReadOnly)

			' Field ship_Province
			Call Shipping.ship_Province.SetDbValue(Rs, Shipping.ship_Province.CurrentValue, Null, Shipping.ship_Province.ReadOnly)

			' Field ship_PostalCode
			Call Shipping.ship_PostalCode.SetDbValue(Rs, Shipping.ship_PostalCode.CurrentValue, Null, Shipping.ship_PostalCode.ReadOnly)

			' Field ship_Country
			Call Shipping.ship_Country.SetDbValue(Rs, Shipping.ship_Country.CurrentValue, Null, Shipping.ship_Country.ReadOnly)

			' Field ship_EmailAddress
			Call Shipping.ship_EmailAddress.SetDbValue(Rs, Shipping.ship_EmailAddress.CurrentValue, Null, Shipping.ship_EmailAddress.ReadOnly)

			' Field HomePhone
			Call Shipping.HomePhone.SetDbValue(Rs, Shipping.HomePhone.CurrentValue, Null, Shipping.HomePhone.ReadOnly)

			' Field WorkPhone
			Call Shipping.WorkPhone.SetDbValue(Rs, Shipping.WorkPhone.CurrentValue, Null, Shipping.WorkPhone.ReadOnly)

			' Field ship_Address2
			Call Shipping.ship_Address2.SetDbValue(Rs, Shipping.ship_Address2.CurrentValue, Null, Shipping.ship_Address2.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Shipping.Row_Updating(RsOld, Rs)
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
				If Shipping.CancelMessage <> "" Then
					FailureMessage = Shipping.CancelMessage
					Shipping.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Shipping.Row_Updated(RsOld, RsNew)
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
