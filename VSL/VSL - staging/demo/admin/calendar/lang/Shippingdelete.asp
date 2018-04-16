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
Dim Shipping_delete
Set Shipping_delete = New cShipping_delete
Set Page = Shipping_delete

' Page init processing
Call Shipping_delete.Page_Init()

' Page main processing
Call Shipping_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Shipping_delete = new ew_Page("Shipping_delete");
// page properties
Shipping_delete.PageID = "delete"; // page ID
Shipping_delete.FormID = "fShippingdelete"; // form ID
var EW_PAGE_ID = Shipping_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Shipping_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Shipping_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Shipping_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Shipping_delete.ShowPageHeader() %>
<%

' Load records for display
Set Shipping_delete.Recordset = Shipping_delete.LoadRecordset()
Shipping_delete.TotalRecs = Shipping_delete.Recordset.RecordCount ' Get record count
If Shipping_delete.TotalRecs <= 0 Then ' No record found, exit
	Shipping_delete.Recordset.Close
	Set Shipping_delete.Recordset = Nothing
	Call Shipping_delete.Page_Terminate("Shippinglist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Shipping.TableCaption %></p>
<p class="aspmaker"><a href="<%= Shipping.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Shipping_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Shipping">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Shipping_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Shipping_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Shipping.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Shipping.CustomerId.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_FirstName.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_LastName.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_Address.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_City.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_Province.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_PostalCode.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_Country.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_EmailAddress.FldCaption %></td>
		<td valign="top"><%= Shipping.HomePhone.FldCaption %></td>
		<td valign="top"><%= Shipping.WorkPhone.FldCaption %></td>
		<td valign="top"><%= Shipping.ship_Address2.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Shipping_delete.RecCnt = 0
i = 0
Do While (Not Shipping_delete.Recordset.Eof)
	Shipping_delete.RecCnt = Shipping_delete.RecCnt + 1

	' Set row properties
	Call Shipping.ResetAttrs()
	Shipping.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Shipping_delete.LoadRowValues(Shipping_delete.Recordset)

	' Render row
	Call Shipping_delete.RenderRow()
%>
	<tr<%= Shipping.RowAttributes %>>
		<td<%= Shipping.CustomerId.CellAttributes %>>
<div<%= Shipping.CustomerId.ViewAttributes %>><%= Shipping.CustomerId.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_FirstName.CellAttributes %>>
<div<%= Shipping.ship_FirstName.ViewAttributes %>><%= Shipping.ship_FirstName.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_LastName.CellAttributes %>>
<div<%= Shipping.ship_LastName.ViewAttributes %>><%= Shipping.ship_LastName.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_Address.CellAttributes %>>
<div<%= Shipping.ship_Address.ViewAttributes %>><%= Shipping.ship_Address.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_City.CellAttributes %>>
<div<%= Shipping.ship_City.ViewAttributes %>><%= Shipping.ship_City.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_Province.CellAttributes %>>
<div<%= Shipping.ship_Province.ViewAttributes %>><%= Shipping.ship_Province.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_PostalCode.CellAttributes %>>
<div<%= Shipping.ship_PostalCode.ViewAttributes %>><%= Shipping.ship_PostalCode.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_Country.CellAttributes %>>
<div<%= Shipping.ship_Country.ViewAttributes %>><%= Shipping.ship_Country.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_EmailAddress.CellAttributes %>>
<div<%= Shipping.ship_EmailAddress.ViewAttributes %>><%= Shipping.ship_EmailAddress.ListViewValue %></div>
</td>
		<td<%= Shipping.HomePhone.CellAttributes %>>
<div<%= Shipping.HomePhone.ViewAttributes %>><%= Shipping.HomePhone.ListViewValue %></div>
</td>
		<td<%= Shipping.WorkPhone.CellAttributes %>>
<div<%= Shipping.WorkPhone.ViewAttributes %>><%= Shipping.WorkPhone.ListViewValue %></div>
</td>
		<td<%= Shipping.ship_Address2.CellAttributes %>>
<div<%= Shipping.ship_Address2.ViewAttributes %>><%= Shipping.ship_Address2.ListViewValue %></div>
</td>
	</tr>
<%
	Shipping_delete.Recordset.MoveNext
Loop
Shipping_delete.Recordset.Close
Set Shipping_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Shipping_delete.ShowPageFooter()
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
Set Shipping_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cShipping_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Shipping"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Shipping_delete"
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
		EW_PAGE_ID = "delete"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset

	' Page main processing
	Sub Page_Main()
		Dim sFilter

		' Load Key Parameters
		RecKeys = Shipping.GetRecordKeys() ' Load record keys
		sFilter = Shipping.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("Shippinglist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Shipping class, Shippinginfo.asp

		Shipping.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Shipping.CurrentAction = Request.Form("a_delete")
		Else
			Shipping.CurrentAction = "I"	' Display record
		End If
		Select Case Shipping.CurrentAction
			Case "D" ' Delete
				Shipping.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Shipping.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Shipping.CurrentFilter
		Call Shipping.Recordset_Selecting(sFilter)
		Shipping.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Shipping.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Shipping.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
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
		End If

		' Call Row Rendered event
		If Shipping.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Shipping.Row_Rendered()
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
		sSql = Shipping.SQL
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
				DeleteRows = Shipping.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("AddressID")
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
			If Shipping.CancelMessage <> "" Then
				FailureMessage = Shipping.CancelMessage
				Shipping.CancelMessage = ""
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
				Call Shipping.Row_Deleted(RsOld)
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
