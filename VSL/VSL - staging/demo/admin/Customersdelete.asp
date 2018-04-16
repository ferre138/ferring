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
Dim Customers_delete
Set Customers_delete = New cCustomers_delete
Set Page = Customers_delete

' Page init processing
Call Customers_delete.Page_Init()

' Page main processing
Call Customers_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Customers_delete = new ew_Page("Customers_delete");
// page properties
Customers_delete.PageID = "delete"; // page ID
Customers_delete.FormID = "fCustomersdelete"; // form ID
var EW_PAGE_ID = Customers_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Customers_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Customers_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Customers_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Customers_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Customers_delete.ShowPageHeader() %>
<%

' Load records for display
Set Customers_delete.Recordset = Customers_delete.LoadRecordset()
Customers_delete.TotalRecs = Customers_delete.Recordset.RecordCount ' Get record count
If Customers_delete.TotalRecs <= 0 Then ' No record found, exit
	Customers_delete.Recordset.Close
	Set Customers_delete.Recordset = Nothing
	Call Customers_delete.Page_Terminate("Customerslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Customers.TableCaption %></p>
<p class="aspmaker"><a href="<%= Customers.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Customers_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Customers">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Customers_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Customers_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Customers.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Customers.Inv_FirstName.FldCaption %></td>
		<td valign="top"><%= Customers.Inv_LastName.FldCaption %></td>
		<td valign="top"><%= Customers.Inv_Address.FldCaption %></td>
		<td valign="top"><%= Customers.inv_City.FldCaption %></td>
		<td valign="top"><%= Customers.inv_EmailAddress.FldCaption %></td>
		<td valign="top"><%= Customers.UserName.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Customers_delete.RecCnt = 0
i = 0
Do While (Not Customers_delete.Recordset.Eof)
	Customers_delete.RecCnt = Customers_delete.RecCnt + 1

	' Set row properties
	Call Customers.ResetAttrs()
	Customers.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Customers_delete.LoadRowValues(Customers_delete.Recordset)

	' Render row
	Call Customers_delete.RenderRow()
%>
	<tr<%= Customers.RowAttributes %>>
		<td<%= Customers.Inv_FirstName.CellAttributes %>>
<div<%= Customers.Inv_FirstName.ViewAttributes %>><%= Customers.Inv_FirstName.ListViewValue %></div>
</td>
		<td<%= Customers.Inv_LastName.CellAttributes %>>
<div<%= Customers.Inv_LastName.ViewAttributes %>><%= Customers.Inv_LastName.ListViewValue %></div>
</td>
		<td<%= Customers.Inv_Address.CellAttributes %>>
<div<%= Customers.Inv_Address.ViewAttributes %>><%= Customers.Inv_Address.ListViewValue %></div>
</td>
		<td<%= Customers.inv_City.CellAttributes %>>
<div<%= Customers.inv_City.ViewAttributes %>><%= Customers.inv_City.ListViewValue %></div>
</td>
		<td<%= Customers.inv_EmailAddress.CellAttributes %>>
<div<%= Customers.inv_EmailAddress.ViewAttributes %>><%= Customers.inv_EmailAddress.ListViewValue %></div>
</td>
		<td<%= Customers.UserName.CellAttributes %>>
<div<%= Customers.UserName.ViewAttributes %>><%= Customers.UserName.ListViewValue %></div>
</td>
	</tr>
<%
	Customers_delete.Recordset.MoveNext
Loop
Customers_delete.Recordset.Close
Set Customers_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Customers_delete.ShowPageFooter()
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
Set Customers_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cCustomers_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Customers"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Customers_delete"
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
		EW_PAGE_ID = "delete"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = Customers.GetRecordKeys() ' Load record keys
		sFilter = Customers.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("Customerslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Customers class, Customersinfo.asp

		Customers.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Customers.CurrentAction = Request.Form("a_delete")
		Else
			Customers.CurrentAction = "I"	' Display record
		End If
		Select Case Customers.CurrentAction
			Case "D" ' Delete
				Customers.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Customers.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Customers.CurrentFilter
		Call Customers.Recordset_Selecting(sFilter)
		Customers.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Customers.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Customers.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
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

			' inv_EmailAddress
			Customers.inv_EmailAddress.LinkCustomAttributes = ""
			Customers.inv_EmailAddress.HrefValue = ""
			Customers.inv_EmailAddress.TooltipValue = ""

			' UserName
			Customers.UserName.LinkCustomAttributes = ""
			Customers.UserName.HrefValue = ""
			Customers.UserName.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Customers.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Customers.Row_Rendered()
		End If
	End Sub

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld
		DeleteRows = True
		sSql = Customers.SQL
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If
		Conn.BeginTrans

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = Customers.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("CustomerID")
				RsDelete.Delete
				If Err.Number <> 0 Then
					FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If Customers.CancelMessage <> "" Then
				FailureMessage = Customers.CancelMessage
				Customers.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
			Conn.CommitTrans ' Commit the changes
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				DeleteRows = False ' Delete failed
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			If Not RsOld.Eof Then RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call Customers.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
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
End Class
%>
