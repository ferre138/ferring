<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="OrderDetailsinfo.asp"-->
<!--#include file="Ordersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim OrderDetails_delete
Set OrderDetails_delete = New cOrderDetails_delete
Set Page = OrderDetails_delete

' Page init processing
Call OrderDetails_delete.Page_Init()

' Page main processing
Call OrderDetails_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var OrderDetails_delete = new ew_Page("OrderDetails_delete");
// page properties
OrderDetails_delete.PageID = "delete"; // page ID
OrderDetails_delete.FormID = "fOrderDetailsdelete"; // form ID
var EW_PAGE_ID = OrderDetails_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
OrderDetails_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
OrderDetails_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
OrderDetails_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
OrderDetails_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% OrderDetails_delete.ShowPageHeader() %>
<%

' Load records for display
Set OrderDetails_delete.Recordset = OrderDetails_delete.LoadRecordset()
OrderDetails_delete.TotalRecs = OrderDetails_delete.Recordset.RecordCount ' Get record count
If OrderDetails_delete.TotalRecs <= 0 Then ' No record found, exit
	OrderDetails_delete.Recordset.Close
	Set OrderDetails_delete.Recordset = Nothing
	Call OrderDetails_delete.Page_Terminate("OrderDetailslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= OrderDetails.TableCaption %></p>
<p class="aspmaker"><a href="<%= OrderDetails.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% OrderDetails_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="OrderDetails">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(OrderDetails_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(OrderDetails_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= OrderDetails.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= OrderDetails.ProductId.FldCaption %></td>
		<td valign="top"><%= OrderDetails.Quantity.FldCaption %></td>
		<td valign="top"><%= OrderDetails.Price.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
OrderDetails_delete.RecCnt = 0
i = 0
Do While (Not OrderDetails_delete.Recordset.Eof)
	OrderDetails_delete.RecCnt = OrderDetails_delete.RecCnt + 1

	' Set row properties
	Call OrderDetails.ResetAttrs()
	OrderDetails.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call OrderDetails_delete.LoadRowValues(OrderDetails_delete.Recordset)

	' Render row
	Call OrderDetails_delete.RenderRow()
%>
	<tr<%= OrderDetails.RowAttributes %>>
		<td<%= OrderDetails.ProductId.CellAttributes %>>
<div<%= OrderDetails.ProductId.ViewAttributes %>><%= OrderDetails.ProductId.ListViewValue %></div>
</td>
		<td<%= OrderDetails.Quantity.CellAttributes %>>
<div<%= OrderDetails.Quantity.ViewAttributes %>><%= OrderDetails.Quantity.ListViewValue %></div>
</td>
		<td<%= OrderDetails.Price.CellAttributes %>>
<div<%= OrderDetails.Price.ViewAttributes %>><%= OrderDetails.Price.ListViewValue %></div>
</td>
	</tr>
<%
	OrderDetails_delete.Recordset.MoveNext
Loop
OrderDetails_delete.Recordset.Close
Set OrderDetails_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
OrderDetails_delete.ShowPageFooter()
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
Set OrderDetails_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrderDetails_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "OrderDetails"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "OrderDetails_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If OrderDetails.UseTokenInUrl Then PageUrl = PageUrl & "t=" & OrderDetails.TableVar & "&" ' add page token
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
		If OrderDetails.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (OrderDetails.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (OrderDetails.TableVar = Request.QueryString("t"))
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
		If IsEmpty(OrderDetails) Then Set OrderDetails = New cOrderDetails
		Set Table = OrderDetails

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Orders) Then Set Orders = New cOrders

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "OrderDetails"

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
		Set OrderDetails = Nothing
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
		RecKeys = OrderDetails.GetRecordKeys() ' Load record keys
		sFilter = OrderDetails.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("OrderDetailslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in OrderDetails class, OrderDetailsinfo.asp

		OrderDetails.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			OrderDetails.CurrentAction = Request.Form("a_delete")
		Else
			OrderDetails.CurrentAction = "I"	' Display record
		End If
		Select Case OrderDetails.CurrentAction
			Case "D" ' Delete
				OrderDetails.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(OrderDetails.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = OrderDetails.CurrentFilter
		Call OrderDetails.Recordset_Selecting(sFilter)
		OrderDetails.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = OrderDetails.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call OrderDetails.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = OrderDetails.KeyFilter

		' Call Row Selecting event
		Call OrderDetails.Row_Selecting(sFilter)

		' Load sql based on filter
		OrderDetails.CurrentFilter = sFilter
		sSql = OrderDetails.SQL
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
		Call OrderDetails.Row_Selected(RsRow)
		OrderDetails.OrderDetailsId.DbValue = RsRow("OrderDetailsId")
		OrderDetails.OrderId.DbValue = RsRow("OrderId")
		OrderDetails.ProductId.DbValue = RsRow("ProductId")
		OrderDetails.Quantity.DbValue = RsRow("Quantity")
		OrderDetails.Price.DbValue = RsRow("Price")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call OrderDetails.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' OrderDetailsId
		' OrderId
		' ProductId
		' Quantity
		' Price
		' -----------
		'  View  Row
		' -----------

		If OrderDetails.RowType = EW_ROWTYPE_VIEW Then ' View row

			' OrderDetailsId
			OrderDetails.OrderDetailsId.ViewValue = OrderDetails.OrderDetailsId.CurrentValue
			OrderDetails.OrderDetailsId.ViewCustomAttributes = ""

			' OrderId
			OrderDetails.OrderId.ViewValue = OrderDetails.OrderId.CurrentValue
			OrderDetails.OrderId.ViewCustomAttributes = ""

			' ProductId
			If OrderDetails.ProductId.CurrentValue & "" <> "" Then
				sFilterWrk = "[ItemId] = " & ew_AdjustSql(OrderDetails.ProductId.CurrentValue) & ""
			sSqlWrk = "SELECT [Description] FROM [Products]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					OrderDetails.ProductId.ViewValue = RsWrk("Description")
				Else
					OrderDetails.ProductId.ViewValue = OrderDetails.ProductId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				OrderDetails.ProductId.ViewValue = Null
			End If
			OrderDetails.ProductId.ViewCustomAttributes = ""

			' Quantity
			OrderDetails.Quantity.ViewValue = OrderDetails.Quantity.CurrentValue
			OrderDetails.Quantity.ViewCustomAttributes = ""

			' Price
			OrderDetails.Price.ViewValue = OrderDetails.Price.CurrentValue
			OrderDetails.Price.ViewCustomAttributes = ""

			' View refer script
			' ProductId

			OrderDetails.ProductId.LinkCustomAttributes = ""
			OrderDetails.ProductId.HrefValue = ""
			OrderDetails.ProductId.TooltipValue = ""

			' Quantity
			OrderDetails.Quantity.LinkCustomAttributes = ""
			OrderDetails.Quantity.HrefValue = ""
			OrderDetails.Quantity.TooltipValue = ""

			' Price
			OrderDetails.Price.LinkCustomAttributes = ""
			OrderDetails.Price.HrefValue = ""
			OrderDetails.Price.TooltipValue = ""
		End If

		' Call Row Rendered event
		If OrderDetails.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call OrderDetails.Row_Rendered()
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
		sSql = OrderDetails.SQL
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
				DeleteRows = OrderDetails.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("OrderDetailsId")
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
			If OrderDetails.CancelMessage <> "" Then
				FailureMessage = OrderDetails.CancelMessage
				OrderDetails.CancelMessage = ""
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
				Call OrderDetails.Row_Deleted(RsOld)
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
