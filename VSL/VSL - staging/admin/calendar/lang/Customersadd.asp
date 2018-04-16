<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Customers_add
Set Customers_add = New cCustomers_add
Set Page = Customers_add

' Page init processing
Call Customers_add.Page_Init()

' Page main processing
Call Customers_add.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Customers_add = new ew_Page("Customers_add");
// page properties
Customers_add.PageID = "add"; // page ID
Customers_add.FormID = "fCustomersadd"; // form ID
var EW_PAGE_ID = Customers_add.PageID; // for backward compatibility
// extend page with ValidateForm function
Customers_add.ValidateForm = function(fobj) {
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
Customers_add.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Customers_add.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Customers_add.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Customers_add.ValidateRequired = false; // no JavaScript validation
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
<% Customers_add.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Add") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Customers.TableCaption %></p>
<p class="aspmaker"><a href="<%= Customers.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Customers_add.ShowMessage %>
<form name="fCustomersadd" id="fCustomersadd" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Customers_add.ValidateForm(this);">
<p>
<input type="hidden" name="t" id="t" value="Customers">
<input type="hidden" name="a_add" id="a_add" value="A">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Customers.Inv_FirstName.Visible Then ' Inv_FirstName %>
	<tr id="r_Inv_FirstName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_FirstName.FldCaption %></td>
		<td<%= Customers.Inv_FirstName.CellAttributes %>><span id="el_Inv_FirstName">
<input type="text" name="x_Inv_FirstName" id="x_Inv_FirstName" size="30" maxlength="30" value="<%= Customers.Inv_FirstName.EditValue %>"<%= Customers.Inv_FirstName.EditAttributes %>>
</span><%= Customers.Inv_FirstName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.Inv_LastName.Visible Then ' Inv_LastName %>
	<tr id="r_Inv_LastName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_LastName.FldCaption %></td>
		<td<%= Customers.Inv_LastName.CellAttributes %>><span id="el_Inv_LastName">
<input type="text" name="x_Inv_LastName" id="x_Inv_LastName" size="30" maxlength="50" value="<%= Customers.Inv_LastName.EditValue %>"<%= Customers.Inv_LastName.EditAttributes %>>
</span><%= Customers.Inv_LastName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.Inv_Address.Visible Then ' Inv_Address %>
	<tr id="r_Inv_Address"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_Address.FldCaption %></td>
		<td<%= Customers.Inv_Address.CellAttributes %>><span id="el_Inv_Address">
<input type="text" name="x_Inv_Address" id="x_Inv_Address" size="30" maxlength="255" value="<%= Customers.Inv_Address.EditValue %>"<%= Customers.Inv_Address.EditAttributes %>>
</span><%= Customers.Inv_Address.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_City.Visible Then ' inv_City %>
	<tr id="r_inv_City"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_City.FldCaption %></td>
		<td<%= Customers.inv_City.CellAttributes %>><span id="el_inv_City">
<input type="text" name="x_inv_City" id="x_inv_City" size="30" maxlength="50" value="<%= Customers.inv_City.EditValue %>"<%= Customers.inv_City.EditAttributes %>>
</span><%= Customers.inv_City.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_Province.Visible Then ' inv_Province %>
	<tr id="r_inv_Province"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Province.FldCaption %></td>
		<td<%= Customers.inv_Province.CellAttributes %>><span id="el_inv_Province">
<input type="text" name="x_inv_Province" id="x_inv_Province" size="30" maxlength="2" value="<%= Customers.inv_Province.EditValue %>"<%= Customers.inv_Province.EditAttributes %>>
</span><%= Customers.inv_Province.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_PostalCode.Visible Then ' inv_PostalCode %>
	<tr id="r_inv_PostalCode"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_PostalCode.FldCaption %></td>
		<td<%= Customers.inv_PostalCode.CellAttributes %>><span id="el_inv_PostalCode">
<input type="text" name="x_inv_PostalCode" id="x_inv_PostalCode" size="30" maxlength="20" value="<%= Customers.inv_PostalCode.EditValue %>"<%= Customers.inv_PostalCode.EditAttributes %>>
</span><%= Customers.inv_PostalCode.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_Country.Visible Then ' inv_Country %>
	<tr id="r_inv_Country"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Country.FldCaption %></td>
		<td<%= Customers.inv_Country.CellAttributes %>><span id="el_inv_Country">
<input type="text" name="x_inv_Country" id="x_inv_Country" size="30" maxlength="50" value="<%= Customers.inv_Country.EditValue %>"<%= Customers.inv_Country.EditAttributes %>>
</span><%= Customers.inv_Country.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_PhoneNumber.Visible Then ' inv_PhoneNumber %>
	<tr id="r_inv_PhoneNumber"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_PhoneNumber.FldCaption %></td>
		<td<%= Customers.inv_PhoneNumber.CellAttributes %>><span id="el_inv_PhoneNumber">
<input type="text" name="x_inv_PhoneNumber" id="x_inv_PhoneNumber" size="30" maxlength="30" value="<%= Customers.inv_PhoneNumber.EditValue %>"<%= Customers.inv_PhoneNumber.EditAttributes %>>
</span><%= Customers.inv_PhoneNumber.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
	<tr id="r_inv_EmailAddress"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_EmailAddress.FldCaption %></td>
		<td<%= Customers.inv_EmailAddress.CellAttributes %>><span id="el_inv_EmailAddress">
<input type="text" name="x_inv_EmailAddress" id="x_inv_EmailAddress" size="30" maxlength="50" value="<%= Customers.inv_EmailAddress.EditValue %>"<%= Customers.inv_EmailAddress.EditAttributes %>>
</span><%= Customers.inv_EmailAddress.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.Notes.Visible Then ' Notes %>
	<tr id="r_Notes"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Notes.FldCaption %></td>
		<td<%= Customers.Notes.CellAttributes %>><span id="el_Notes">
<textarea name="x_Notes" id="x_Notes" cols="35" rows="4"<%= Customers.Notes.EditAttributes %>><%= Customers.Notes.EditValue %></textarea>
</span><%= Customers.Notes.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.inv_Fax.Visible Then ' inv_Fax %>
	<tr id="r_inv_Fax"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Fax.FldCaption %></td>
		<td<%= Customers.inv_Fax.CellAttributes %>><span id="el_inv_Fax">
<input type="text" name="x_inv_Fax" id="x_inv_Fax" size="30" maxlength="30" value="<%= Customers.inv_Fax.EditValue %>"<%= Customers.inv_Fax.EditAttributes %>>
</span><%= Customers.inv_Fax.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.Inv_Address2.Visible Then ' Inv_Address2 %>
	<tr id="r_Inv_Address2"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_Address2.FldCaption %></td>
		<td<%= Customers.Inv_Address2.CellAttributes %>><span id="el_Inv_Address2">
<input type="text" name="x_Inv_Address2" id="x_Inv_Address2" size="30" maxlength="255" value="<%= Customers.Inv_Address2.EditValue %>"<%= Customers.Inv_Address2.EditAttributes %>>
</span><%= Customers.Inv_Address2.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.UserName.Visible Then ' UserName %>
	<tr id="r_UserName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.UserName.FldCaption %></td>
		<td<%= Customers.UserName.CellAttributes %>><span id="el_UserName">
<input type="text" name="x_UserName" id="x_UserName" size="30" maxlength="15" value="<%= Customers.UserName.EditValue %>"<%= Customers.UserName.EditAttributes %>>
</span><%= Customers.UserName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.passwrd.Visible Then ' passwrd %>
	<tr id="r_passwrd"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.passwrd.FldCaption %></td>
		<td<%= Customers.passwrd.CellAttributes %>><span id="el_passwrd">
<input type="text" name="x_passwrd" id="x_passwrd" size="30" maxlength="15" value="<%= Customers.passwrd.EditValue %>"<%= Customers.passwrd.EditAttributes %>>
</span><%= Customers.passwrd.CustomMsg %></td>
	</tr>
<% End If %>
<% If Customers.NewCustomer.Visible Then ' NewCustomer %>
	<tr id="r_NewCustomer"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.NewCustomer.FldCaption %></td>
		<td<%= Customers.NewCustomer.CellAttributes %>><span id="el_NewCustomer">
<% selwrk = ew_IIf(ew_ConvertToBool(Customers.NewCustomer.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_NewCustomer" id="x_NewCustomer" value="1"<%= selwrk %><%= Customers.NewCustomer.EditAttributes %>>
</span><%= Customers.NewCustomer.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("AddBtn")) %>">
</form>
<%
Customers_add.ShowPageFooter()
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
Set Customers_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cCustomers_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Customers"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Customers_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Customers.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Customers.TableVar & "&" ' add page token
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
		If Customers.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Customers.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Customers.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Customers) Then Set Customers = New cCustomers
		Set Table = Customers

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Customers"

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
		Set Customers = Nothing
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
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			Customers.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

			' Validate Form
			If Not ValidateForm() Then
				Customers.CurrentAction = "I" ' Form error, reset action
				Customers.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("CustomerID").Count > 0 Then
				Customers.CustomerID.QueryStringValue = Request.QueryString("CustomerID")
				Call Customers.SetKey("CustomerID", Customers.CustomerID.CurrentValue) ' Set up key
			Else
				Call Customers.SetKey("CustomerID", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				Customers.CurrentAction = "C" ' Copy Record
			Else
				Customers.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Perform action based on action code
		Select Case Customers.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Customerslist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				Customers.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = Customers.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "Customersview.asp" Then sReturnUrl = Customers.ViewUrl ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					Customers.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		Customers.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call Customers.ResetAttrs()
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
	' Load default values
	'
	Function LoadDefaultValues()
		Customers.Inv_FirstName.CurrentValue = Null
		Customers.Inv_FirstName.OldValue = Customers.Inv_FirstName.CurrentValue
		Customers.Inv_LastName.CurrentValue = Null
		Customers.Inv_LastName.OldValue = Customers.Inv_LastName.CurrentValue
		Customers.Inv_Address.CurrentValue = Null
		Customers.Inv_Address.OldValue = Customers.Inv_Address.CurrentValue
		Customers.inv_City.CurrentValue = Null
		Customers.inv_City.OldValue = Customers.inv_City.CurrentValue
		Customers.inv_Province.CurrentValue = Null
		Customers.inv_Province.OldValue = Customers.inv_Province.CurrentValue
		Customers.inv_PostalCode.CurrentValue = Null
		Customers.inv_PostalCode.OldValue = Customers.inv_PostalCode.CurrentValue
		Customers.inv_Country.CurrentValue = Null
		Customers.inv_Country.OldValue = Customers.inv_Country.CurrentValue
		Customers.inv_PhoneNumber.CurrentValue = Null
		Customers.inv_PhoneNumber.OldValue = Customers.inv_PhoneNumber.CurrentValue
		Customers.inv_EmailAddress.CurrentValue = Null
		Customers.inv_EmailAddress.OldValue = Customers.inv_EmailAddress.CurrentValue
		Customers.Notes.CurrentValue = Null
		Customers.Notes.OldValue = Customers.Notes.CurrentValue
		Customers.inv_Fax.CurrentValue = Null
		Customers.inv_Fax.OldValue = Customers.inv_Fax.CurrentValue
		Customers.Inv_Address2.CurrentValue = Null
		Customers.Inv_Address2.OldValue = Customers.Inv_Address2.CurrentValue
		Customers.UserName.CurrentValue = Null
		Customers.UserName.OldValue = Customers.UserName.CurrentValue
		Customers.passwrd.CurrentValue = Null
		Customers.passwrd.OldValue = Customers.passwrd.CurrentValue
		Customers.NewCustomer.CurrentValue = "1"
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not Customers.Inv_FirstName.FldIsDetailKey Then Customers.Inv_FirstName.FormValue = ObjForm.GetValue("x_Inv_FirstName")
		If Not Customers.Inv_LastName.FldIsDetailKey Then Customers.Inv_LastName.FormValue = ObjForm.GetValue("x_Inv_LastName")
		If Not Customers.Inv_Address.FldIsDetailKey Then Customers.Inv_Address.FormValue = ObjForm.GetValue("x_Inv_Address")
		If Not Customers.inv_City.FldIsDetailKey Then Customers.inv_City.FormValue = ObjForm.GetValue("x_inv_City")
		If Not Customers.inv_Province.FldIsDetailKey Then Customers.inv_Province.FormValue = ObjForm.GetValue("x_inv_Province")
		If Not Customers.inv_PostalCode.FldIsDetailKey Then Customers.inv_PostalCode.FormValue = ObjForm.GetValue("x_inv_PostalCode")
		If Not Customers.inv_Country.FldIsDetailKey Then Customers.inv_Country.FormValue = ObjForm.GetValue("x_inv_Country")
		If Not Customers.inv_PhoneNumber.FldIsDetailKey Then Customers.inv_PhoneNumber.FormValue = ObjForm.GetValue("x_inv_PhoneNumber")
		If Not Customers.inv_EmailAddress.FldIsDetailKey Then Customers.inv_EmailAddress.FormValue = ObjForm.GetValue("x_inv_EmailAddress")
		If Not Customers.Notes.FldIsDetailKey Then Customers.Notes.FormValue = ObjForm.GetValue("x_Notes")
		If Not Customers.inv_Fax.FldIsDetailKey Then Customers.inv_Fax.FormValue = ObjForm.GetValue("x_inv_Fax")
		If Not Customers.Inv_Address2.FldIsDetailKey Then Customers.Inv_Address2.FormValue = ObjForm.GetValue("x_Inv_Address2")
		If Not Customers.UserName.FldIsDetailKey Then Customers.UserName.FormValue = ObjForm.GetValue("x_UserName")
		If Not Customers.passwrd.FldIsDetailKey Then Customers.passwrd.FormValue = ObjForm.GetValue("x_passwrd")
		If Not Customers.NewCustomer.FldIsDetailKey Then Customers.NewCustomer.FormValue = ObjForm.GetValue("x_NewCustomer")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		Customers.Inv_FirstName.CurrentValue = Customers.Inv_FirstName.FormValue
		Customers.Inv_LastName.CurrentValue = Customers.Inv_LastName.FormValue
		Customers.Inv_Address.CurrentValue = Customers.Inv_Address.FormValue
		Customers.inv_City.CurrentValue = Customers.inv_City.FormValue
		Customers.inv_Province.CurrentValue = Customers.inv_Province.FormValue
		Customers.inv_PostalCode.CurrentValue = Customers.inv_PostalCode.FormValue
		Customers.inv_Country.CurrentValue = Customers.inv_Country.FormValue
		Customers.inv_PhoneNumber.CurrentValue = Customers.inv_PhoneNumber.FormValue
		Customers.inv_EmailAddress.CurrentValue = Customers.inv_EmailAddress.FormValue
		Customers.Notes.CurrentValue = Customers.Notes.FormValue
		Customers.inv_Fax.CurrentValue = Customers.inv_Fax.FormValue
		Customers.Inv_Address2.CurrentValue = Customers.Inv_Address2.FormValue
		Customers.UserName.CurrentValue = Customers.UserName.FormValue
		Customers.passwrd.CurrentValue = Customers.passwrd.FormValue
		Customers.NewCustomer.CurrentValue = Customers.NewCustomer.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Customers.KeyFilter

		' Call Row Selecting event
		Call Customers.Row_Selecting(sFilter)

		' Load sql based on filter
		Customers.CurrentFilter = sFilter
		sSql = Customers.SQL
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
		Call Customers.Row_Selected(RsRow)
		Customers.CustomerID.DbValue = RsRow("CustomerID")
		Customers.Inv_FirstName.DbValue = RsRow("Inv_FirstName")
		Customers.Inv_LastName.DbValue = RsRow("Inv_LastName")
		Customers.Inv_Address.DbValue = RsRow("Inv_Address")
		Customers.inv_City.DbValue = RsRow("inv_City")
		Customers.inv_Province.DbValue = RsRow("inv_Province")
		Customers.inv_PostalCode.DbValue = RsRow("inv_PostalCode")
		Customers.inv_Country.DbValue = RsRow("inv_Country")
		Customers.inv_PhoneNumber.DbValue = RsRow("inv_PhoneNumber")
		Customers.inv_EmailAddress.DbValue = RsRow("inv_EmailAddress")
		Customers.Notes.DbValue = RsRow("Notes")
		Customers.inv_Fax.DbValue = RsRow("inv_Fax")
		Customers.Inv_Address2.DbValue = RsRow("Inv_Address2")
		Customers.UserName.DbValue = RsRow("UserName")
		Customers.passwrd.DbValue = RsRow("passwrd")
		Customers.NewCustomer.DbValue = ew_IIf(RsRow("NewCustomer"), "1", "0")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Customers.GetKey("CustomerID")&"" <> "" Then
			Customers.CustomerID.CurrentValue = Customers.GetKey("CustomerID") ' CustomerID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Customers.CurrentFilter = Customers.KeyFilter
			Dim sSql
			sSql = Customers.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Customers.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' CustomerID
		' Inv_FirstName
		' Inv_LastName
		' Inv_Address
		' inv_City
		' inv_Province
		' inv_PostalCode
		' inv_Country
		' inv_PhoneNumber
		' inv_EmailAddress
		' Notes
		' inv_Fax
		' Inv_Address2
		' UserName
		' passwrd
		' NewCustomer
		' -----------
		'  View  Row
		' -----------

		If Customers.RowType = EW_ROWTYPE_VIEW Then ' View row

			' CustomerID
			Customers.CustomerID.ViewValue = Customers.CustomerID.CurrentValue
			Customers.CustomerID.ViewCustomAttributes = ""

			' Inv_FirstName
			Customers.Inv_FirstName.ViewValue = Customers.Inv_FirstName.CurrentValue
			Customers.Inv_FirstName.ViewCustomAttributes = ""

			' Inv_LastName
			Customers.Inv_LastName.ViewValue = Customers.Inv_LastName.CurrentValue
			Customers.Inv_LastName.ViewCustomAttributes = ""

			' Inv_Address
			Customers.Inv_Address.ViewValue = Customers.Inv_Address.CurrentValue
			Customers.Inv_Address.ViewCustomAttributes = ""

			' inv_City
			Customers.inv_City.ViewValue = Customers.inv_City.CurrentValue
			Customers.inv_City.ViewCustomAttributes = ""

			' inv_Province
			Customers.inv_Province.ViewValue = Customers.inv_Province.CurrentValue
			Customers.inv_Province.ViewCustomAttributes = ""

			' inv_PostalCode
			Customers.inv_PostalCode.ViewValue = Customers.inv_PostalCode.CurrentValue
			Customers.inv_PostalCode.ViewCustomAttributes = ""

			' inv_Country
			Customers.inv_Country.ViewValue = Customers.inv_Country.CurrentValue
			Customers.inv_Country.ViewCustomAttributes = ""

			' inv_PhoneNumber
			Customers.inv_PhoneNumber.ViewValue = Customers.inv_PhoneNumber.CurrentValue
			Customers.inv_PhoneNumber.ViewCustomAttributes = ""

			' inv_EmailAddress
			Customers.inv_EmailAddress.ViewValue = Customers.inv_EmailAddress.CurrentValue
			Customers.inv_EmailAddress.ViewCustomAttributes = ""

			' Notes
			Customers.Notes.ViewValue = Customers.Notes.CurrentValue
			Customers.Notes.ViewCustomAttributes = ""

			' inv_Fax
			Customers.inv_Fax.ViewValue = Customers.inv_Fax.CurrentValue
			Customers.inv_Fax.ViewCustomAttributes = ""

			' Inv_Address2
			Customers.Inv_Address2.ViewValue = Customers.Inv_Address2.CurrentValue
			Customers.Inv_Address2.ViewCustomAttributes = ""

			' UserName
			Customers.UserName.ViewValue = Customers.UserName.CurrentValue
			Customers.UserName.ViewCustomAttributes = ""

			' passwrd
			Customers.passwrd.ViewValue = Customers.passwrd.CurrentValue
			Customers.passwrd.ViewCustomAttributes = ""

			' NewCustomer
			If ew_ConvertToBool(Customers.NewCustomer.CurrentValue) Then
				Customers.NewCustomer.ViewValue = ew_IIf(Customers.NewCustomer.FldTagCaption(1) <> "", Customers.NewCustomer.FldTagCaption(1), "Yes")
			Else
				Customers.NewCustomer.ViewValue = ew_IIf(Customers.NewCustomer.FldTagCaption(2) <> "", Customers.NewCustomer.FldTagCaption(2), "No")
			End If
			Customers.NewCustomer.ViewCustomAttributes = ""

			' View refer script
			' Inv_FirstName

			Customers.Inv_FirstName.LinkCustomAttributes = ""
			Customers.Inv_FirstName.HrefValue = ""
			Customers.Inv_FirstName.TooltipValue = ""

			' Inv_LastName
			Customers.Inv_LastName.LinkCustomAttributes = ""
			Customers.Inv_LastName.HrefValue = ""
			Customers.Inv_LastName.TooltipValue = ""

			' Inv_Address
			Customers.Inv_Address.LinkCustomAttributes = ""
			Customers.Inv_Address.HrefValue = ""
			Customers.Inv_Address.TooltipValue = ""

			' inv_City
			Customers.inv_City.LinkCustomAttributes = ""
			Customers.inv_City.HrefValue = ""
			Customers.inv_City.TooltipValue = ""

			' inv_Province
			Customers.inv_Province.LinkCustomAttributes = ""
			Customers.inv_Province.HrefValue = ""
			Customers.inv_Province.TooltipValue = ""

			' inv_PostalCode
			Customers.inv_PostalCode.LinkCustomAttributes = ""
			Customers.inv_PostalCode.HrefValue = ""
			Customers.inv_PostalCode.TooltipValue = ""

			' inv_Country
			Customers.inv_Country.LinkCustomAttributes = ""
			Customers.inv_Country.HrefValue = ""
			Customers.inv_Country.TooltipValue = ""

			' inv_PhoneNumber
			Customers.inv_PhoneNumber.LinkCustomAttributes = ""
			Customers.inv_PhoneNumber.HrefValue = ""
			Customers.inv_PhoneNumber.TooltipValue = ""

			' inv_EmailAddress
			Customers.inv_EmailAddress.LinkCustomAttributes = ""
			Customers.inv_EmailAddress.HrefValue = ""
			Customers.inv_EmailAddress.TooltipValue = ""

			' Notes
			Customers.Notes.LinkCustomAttributes = ""
			Customers.Notes.HrefValue = ""
			Customers.Notes.TooltipValue = ""

			' inv_Fax
			Customers.inv_Fax.LinkCustomAttributes = ""
			Customers.inv_Fax.HrefValue = ""
			Customers.inv_Fax.TooltipValue = ""

			' Inv_Address2
			Customers.Inv_Address2.LinkCustomAttributes = ""
			Customers.Inv_Address2.HrefValue = ""
			Customers.Inv_Address2.TooltipValue = ""

			' UserName
			Customers.UserName.LinkCustomAttributes = ""
			Customers.UserName.HrefValue = ""
			Customers.UserName.TooltipValue = ""

			' passwrd
			Customers.passwrd.LinkCustomAttributes = ""
			Customers.passwrd.HrefValue = ""
			Customers.passwrd.TooltipValue = ""

			' NewCustomer
			Customers.NewCustomer.LinkCustomAttributes = ""
			Customers.NewCustomer.HrefValue = ""
			Customers.NewCustomer.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf Customers.RowType = EW_ROWTYPE_ADD Then ' Add row

			' Inv_FirstName
			Customers.Inv_FirstName.EditCustomAttributes = ""
			Customers.Inv_FirstName.EditValue = ew_HtmlEncode(Customers.Inv_FirstName.CurrentValue)

			' Inv_LastName
			Customers.Inv_LastName.EditCustomAttributes = ""
			Customers.Inv_LastName.EditValue = ew_HtmlEncode(Customers.Inv_LastName.CurrentValue)

			' Inv_Address
			Customers.Inv_Address.EditCustomAttributes = ""
			Customers.Inv_Address.EditValue = ew_HtmlEncode(Customers.Inv_Address.CurrentValue)

			' inv_City
			Customers.inv_City.EditCustomAttributes = ""
			Customers.inv_City.EditValue = ew_HtmlEncode(Customers.inv_City.CurrentValue)

			' inv_Province
			Customers.inv_Province.EditCustomAttributes = ""
			Customers.inv_Province.EditValue = ew_HtmlEncode(Customers.inv_Province.CurrentValue)

			' inv_PostalCode
			Customers.inv_PostalCode.EditCustomAttributes = ""
			Customers.inv_PostalCode.EditValue = ew_HtmlEncode(Customers.inv_PostalCode.CurrentValue)

			' inv_Country
			Customers.inv_Country.EditCustomAttributes = ""
			Customers.inv_Country.EditValue = ew_HtmlEncode(Customers.inv_Country.CurrentValue)

			' inv_PhoneNumber
			Customers.inv_PhoneNumber.EditCustomAttributes = ""
			Customers.inv_PhoneNumber.EditValue = ew_HtmlEncode(Customers.inv_PhoneNumber.CurrentValue)

			' inv_EmailAddress
			Customers.inv_EmailAddress.EditCustomAttributes = ""
			Customers.inv_EmailAddress.EditValue = ew_HtmlEncode(Customers.inv_EmailAddress.CurrentValue)

			' Notes
			Customers.Notes.EditCustomAttributes = ""
			Customers.Notes.EditValue = ew_HtmlEncode(Customers.Notes.CurrentValue)

			' inv_Fax
			Customers.inv_Fax.EditCustomAttributes = ""
			Customers.inv_Fax.EditValue = ew_HtmlEncode(Customers.inv_Fax.CurrentValue)

			' Inv_Address2
			Customers.Inv_Address2.EditCustomAttributes = ""
			Customers.Inv_Address2.EditValue = ew_HtmlEncode(Customers.Inv_Address2.CurrentValue)

			' UserName
			Customers.UserName.EditCustomAttributes = ""
			Customers.UserName.EditValue = ew_HtmlEncode(Customers.UserName.CurrentValue)

			' passwrd
			Customers.passwrd.EditCustomAttributes = ""
			Customers.passwrd.EditValue = ew_HtmlEncode(Customers.passwrd.CurrentValue)

			' NewCustomer
			Customers.NewCustomer.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Customers.NewCustomer.FldTagCaption(1) <> "", Customers.NewCustomer.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Customers.NewCustomer.FldTagCaption(2) <> "", Customers.NewCustomer.FldTagCaption(2), "No")
			Customers.NewCustomer.EditValue = arwrk

			' Edit refer script
			' Inv_FirstName

			Customers.Inv_FirstName.HrefValue = ""

			' Inv_LastName
			Customers.Inv_LastName.HrefValue = ""

			' Inv_Address
			Customers.Inv_Address.HrefValue = ""

			' inv_City
			Customers.inv_City.HrefValue = ""

			' inv_Province
			Customers.inv_Province.HrefValue = ""

			' inv_PostalCode
			Customers.inv_PostalCode.HrefValue = ""

			' inv_Country
			Customers.inv_Country.HrefValue = ""

			' inv_PhoneNumber
			Customers.inv_PhoneNumber.HrefValue = ""

			' inv_EmailAddress
			Customers.inv_EmailAddress.HrefValue = ""

			' Notes
			Customers.Notes.HrefValue = ""

			' inv_Fax
			Customers.inv_Fax.HrefValue = ""

			' Inv_Address2
			Customers.Inv_Address2.HrefValue = ""

			' UserName
			Customers.UserName.HrefValue = ""

			' passwrd
			Customers.passwrd.HrefValue = ""

			' NewCustomer
			Customers.NewCustomer.HrefValue = ""
		End If
		If Customers.RowType = EW_ROWTYPE_ADD Or Customers.RowType = EW_ROWTYPE_EDIT Or Customers.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Customers.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Customers.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Customers.Row_Rendered()
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
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear

		' Add new record
		sFilter = "(0 = 1)"
		Customers.CurrentFilter = sFilter
		sSql = Customers.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field Inv_FirstName
		Call Customers.Inv_FirstName.SetDbValue(Rs, Customers.Inv_FirstName.CurrentValue, Null, False)

		' Field Inv_LastName
		Call Customers.Inv_LastName.SetDbValue(Rs, Customers.Inv_LastName.CurrentValue, Null, False)

		' Field Inv_Address
		Call Customers.Inv_Address.SetDbValue(Rs, Customers.Inv_Address.CurrentValue, Null, False)

		' Field inv_City
		Call Customers.inv_City.SetDbValue(Rs, Customers.inv_City.CurrentValue, Null, False)

		' Field inv_Province
		Call Customers.inv_Province.SetDbValue(Rs, Customers.inv_Province.CurrentValue, Null, False)

		' Field inv_PostalCode
		Call Customers.inv_PostalCode.SetDbValue(Rs, Customers.inv_PostalCode.CurrentValue, Null, False)

		' Field inv_Country
		Call Customers.inv_Country.SetDbValue(Rs, Customers.inv_Country.CurrentValue, Null, False)

		' Field inv_PhoneNumber
		Call Customers.inv_PhoneNumber.SetDbValue(Rs, Customers.inv_PhoneNumber.CurrentValue, Null, False)

		' Field inv_EmailAddress
		Call Customers.inv_EmailAddress.SetDbValue(Rs, Customers.inv_EmailAddress.CurrentValue, Null, False)

		' Field Notes
		Call Customers.Notes.SetDbValue(Rs, Customers.Notes.CurrentValue, Null, False)

		' Field inv_Fax
		Call Customers.inv_Fax.SetDbValue(Rs, Customers.inv_Fax.CurrentValue, Null, False)

		' Field Inv_Address2
		Call Customers.Inv_Address2.SetDbValue(Rs, Customers.Inv_Address2.CurrentValue, Null, False)

		' Field UserName
		Call Customers.UserName.SetDbValue(Rs, Customers.UserName.CurrentValue, Null, False)

		' Field passwrd
		Call Customers.passwrd.SetDbValue(Rs, Customers.passwrd.CurrentValue, Null, False)

		' Field NewCustomer
		boolwrk = Customers.NewCustomer.CurrentValue
		If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
		Call Customers.NewCustomer.SetDbValue(Rs, boolwrk, Null, (Customers.NewCustomer.CurrentValue&"" = ""))

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = Customers.Row_Inserting(RsOld, Rs)
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
		Else
			Rs.CancelUpdate
			If Customers.CancelMessage <> "" Then
				FailureMessage = Customers.CancelMessage
				Customers.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			Customers.CustomerID.DbValue = RsNew("CustomerID")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call Customers.Row_Inserted(RsOld, RsNew)
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
