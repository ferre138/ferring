<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Productsinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Products_delete
Set Products_delete = New cProducts_delete
Set Page = Products_delete

' Page init processing
Call Products_delete.Page_Init()

' Page main processing
Call Products_delete.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Products_delete = new ew_Page("Products_delete");
// page properties
Products_delete.PageID = "delete"; // page ID
Products_delete.FormID = "fProductsdelete"; // form ID
var EW_PAGE_ID = Products_delete.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Products_delete.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Products_delete.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Products_delete.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Products_delete.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% Products_delete.ShowPageHeader() %>
<%

' Load records for display
Set Products_delete.Recordset = Products_delete.LoadRecordset()
Products_delete.TotalRecs = Products_delete.Recordset.RecordCount ' Get record count
If Products_delete.TotalRecs <= 0 Then ' No record found, exit
	Products_delete.Recordset.Close
	Set Products_delete.Recordset = Nothing
	Call Products_delete.Page_Terminate("Productslist.asp") ' Return to list
End If
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("Delete") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Products.TableCaption %></p>
<p class="aspmaker"><a href="<%= Products.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Products_delete.ShowMessage %>
<form action="<%= ew_CurrentPage %>" method="post">
<p>
<input type="hidden" name="t" id="t" value="Products">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(Products_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(Products_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
<%= Products.TableCustomInnerHTML %>
	<thead>
	<tr class="ewTableHeader">
		<td valign="top"><%= Products.Description.FldCaption %></td>
		<td valign="top"><%= Products.Price.FldCaption %></td>
		<td valign="top"><%= Products.Active.FldCaption %></td>
		<td valign="top"><%= Products.Sizes.FldCaption %></td>
		<td valign="top"><%= Products.Image_Thumb.FldCaption %></td>
		<td valign="top"><%= Products.ProductName.FldCaption %></td>
		<td valign="top"><%= Products.ItemNo.FldCaption %></td>
		<td valign="top"><%= Products.UPC.FldCaption %></td>
	</tr>
	</thead>
	<tbody>
<%
Products_delete.RecCnt = 0
i = 0
Do While (Not Products_delete.Recordset.Eof)
	Products_delete.RecCnt = Products_delete.RecCnt + 1

	' Set row properties
	Call Products.ResetAttrs()
	Products.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call Products_delete.LoadRowValues(Products_delete.Recordset)

	' Render row
	Call Products_delete.RenderRow()
%>
	<tr<%= Products.RowAttributes %>>
		<td<%= Products.Description.CellAttributes %>>
<div<%= Products.Description.ViewAttributes %>><%= Products.Description.ListViewValue %></div>
</td>
		<td<%= Products.Price.CellAttributes %>>
<div<%= Products.Price.ViewAttributes %>><%= Products.Price.ListViewValue %></div>
</td>
		<td<%= Products.Active.CellAttributes %>>
<% If ew_ConvertToBool(Products.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Products.Active.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Products.Active.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
		<td<%= Products.Sizes.CellAttributes %>>
<div<%= Products.Sizes.ViewAttributes %>><%= Products.Sizes.ListViewValue %></div>
</td>
		<td<%= Products.Image_Thumb.CellAttributes %>>
<% If Products.Image_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% End If %>
</td>
		<td<%= Products.ProductName.CellAttributes %>>
<div<%= Products.ProductName.ViewAttributes %>><%= Products.ProductName.ListViewValue %></div>
</td>
		<td<%= Products.ItemNo.CellAttributes %>>
<div<%= Products.ItemNo.ViewAttributes %>><%= Products.ItemNo.ListViewValue %></div>
</td>
		<td<%= Products.UPC.CellAttributes %>>
<div<%= Products.UPC.ViewAttributes %>><%= Products.UPC.ListViewValue %></div>
</td>
	</tr>
<%
	Products_delete.Recordset.MoveNext
Loop
Products_delete.Recordset.Close
Set Products_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("DeleteBtn")) %>">
</form>
<%
Products_delete.ShowPageFooter()
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
Set Products_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProducts_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Products"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Products_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Products.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Products.TableVar & "&" ' add page token
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
		If Products.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Products.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Products.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Products) Then Set Products = New cProducts
		Set Table = Products

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Products"

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
		Set Products = Nothing
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
		RecKeys = Products.GetRecordKeys() ' Load record keys
		sFilter = Products.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("Productslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Products class, Productsinfo.asp

		Products.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			Products.CurrentAction = Request.Form("a_delete")
		Else
			Products.CurrentAction = "I"	' Display record
		End If
		Select Case Products.CurrentAction
			Case "D" ' Delete
				Products.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(Products.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Products.CurrentFilter
		Call Products.Recordset_Selecting(sFilter)
		Products.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Products.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Products.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Products.KeyFilter

		' Call Row Selecting event
		Call Products.Row_Selecting(sFilter)

		' Load sql based on filter
		Products.CurrentFilter = sFilter
		sSql = Products.SQL
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
		Call Products.Row_Selected(RsRow)
		Products.ItemId.DbValue = RsRow("ItemId")
		Products.Description.DbValue = RsRow("Description")
		Products.Price.DbValue = RsRow("Price")
		Products.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Products.Image.Upload.DbValue = RsRow("Image")
		Products.Sizes.DbValue = RsRow("Sizes")
		Products.Image_Thumb.Upload.DbValue = RsRow("Image_Thumb")
		Products.ProductName.DbValue = RsRow("ProductName")
		Products.ItemNo.DbValue = RsRow("ItemNo")
		Products.UPC.DbValue = RsRow("UPC")
		Products.Price_rebate.DbValue = RsRow("Price_rebate")
		Products.fDescription.DbValue = RsRow("fDescription")
		Products.fImage.Upload.DbValue = RsRow("fImage")
		Products.fSizes.DbValue = RsRow("fSizes")
		Products.fImage_Thumb.Upload.DbValue = RsRow("fImage_Thumb")
		Products.fProductName.DbValue = RsRow("fProductName")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Products.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' ItemId
		' Description
		' Price
		' Active
		' Image
		' Sizes
		' Image_Thumb
		' ProductName
		' ItemNo
		' UPC
		' Price_rebate
		' fDescription
		' fImage
		' fSizes
		' fImage_Thumb
		' fProductName
		' -----------
		'  View  Row
		' -----------

		If Products.RowType = EW_ROWTYPE_VIEW Then ' View row

			' ItemId
			Products.ItemId.ViewValue = Products.ItemId.CurrentValue
			Products.ItemId.ViewCustomAttributes = ""

			' Description
			Products.Description.ViewValue = Products.Description.CurrentValue
			Products.Description.ViewCustomAttributes = ""

			' Price
			Products.Price.ViewValue = Products.Price.CurrentValue
			Products.Price.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Products.Active.CurrentValue) Then
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(1) <> "", Products.Active.FldTagCaption(1), "Yes")
			Else
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(2) <> "", Products.Active.FldTagCaption(2), "No")
			End If
			Products.Active.ViewCustomAttributes = ""

			' Image
			If Not ew_Empty(Products.Image.Upload.DbValue) Then
				Products.Image.ViewValue = Products.Image.Upload.DbValue
				Products.Image.ImageAlt = Products.Image.FldAlt
			Else
				Products.Image.ViewValue = ""
			End If
			Products.Image.ViewCustomAttributes = ""

			' Sizes
			Products.Sizes.ViewValue = Products.Sizes.CurrentValue
			Products.Sizes.ViewCustomAttributes = ""

			' Image_Thumb
			If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then
				Products.Image_Thumb.ViewValue = Products.Image_Thumb.Upload.DbValue
				Products.Image_Thumb.ImageAlt = Products.Image_Thumb.FldAlt
			Else
				Products.Image_Thumb.ViewValue = ""
			End If
			Products.Image_Thumb.ViewCustomAttributes = ""

			' ProductName
			Products.ProductName.ViewValue = Products.ProductName.CurrentValue
			Products.ProductName.ViewCustomAttributes = ""

			' ItemNo
			Products.ItemNo.ViewValue = Products.ItemNo.CurrentValue
			Products.ItemNo.ViewCustomAttributes = ""

			' UPC
			Products.UPC.ViewValue = Products.UPC.CurrentValue
			Products.UPC.ViewCustomAttributes = ""

			' Price_rebate
			Products.Price_rebate.ViewValue = Products.Price_rebate.CurrentValue
			Products.Price_rebate.ViewCustomAttributes = ""

			' fDescription
			Products.fDescription.ViewValue = Products.fDescription.CurrentValue
			Products.fDescription.ViewCustomAttributes = ""

			' fImage
			If Not ew_Empty(Products.fImage.Upload.DbValue) Then
				Products.fImage.ViewValue = Products.fImage.Upload.DbValue
				Products.fImage.ImageAlt = Products.fImage.FldAlt
			Else
				Products.fImage.ViewValue = ""
			End If
			Products.fImage.ViewCustomAttributes = ""

			' fSizes
			Products.fSizes.ViewValue = Products.fSizes.CurrentValue
			Products.fSizes.ViewCustomAttributes = ""

			' fImage_Thumb
			If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then
				Products.fImage_Thumb.ViewValue = Products.fImage_Thumb.Upload.DbValue
				Products.fImage_Thumb.ImageAlt = Products.fImage_Thumb.FldAlt
			Else
				Products.fImage_Thumb.ViewValue = ""
			End If
			Products.fImage_Thumb.ViewCustomAttributes = ""

			' fProductName
			Products.fProductName.ViewValue = Products.fProductName.CurrentValue
			Products.fProductName.ViewCustomAttributes = ""

			' View refer script
			' Description

			Products.Description.LinkCustomAttributes = ""
			Products.Description.HrefValue = ""
			Products.Description.TooltipValue = ""

			' Price
			Products.Price.LinkCustomAttributes = ""
			Products.Price.HrefValue = ""
			Products.Price.TooltipValue = ""

			' Active
			Products.Active.LinkCustomAttributes = ""
			Products.Active.HrefValue = ""
			Products.Active.TooltipValue = ""

			' Sizes
			Products.Sizes.LinkCustomAttributes = ""
			Products.Sizes.HrefValue = ""
			Products.Sizes.TooltipValue = ""

			' Image_Thumb
			Products.Image_Thumb.LinkCustomAttributes = ""
			Products.Image_Thumb.HrefValue = ""
			Products.Image_Thumb.TooltipValue = ""

			' ProductName
			Products.ProductName.LinkCustomAttributes = ""
			Products.ProductName.HrefValue = ""
			Products.ProductName.TooltipValue = ""

			' ItemNo
			Products.ItemNo.LinkCustomAttributes = ""
			Products.ItemNo.HrefValue = ""
			Products.ItemNo.TooltipValue = ""

			' UPC
			Products.UPC.LinkCustomAttributes = ""
			Products.UPC.HrefValue = ""
			Products.UPC.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Products.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Products.Row_Rendered()
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
		sSql = Products.SQL
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
				DeleteRows = Products.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("ItemId")
				ew_DeleteFile ew_UploadPathEx(True, Products.Image.UploadPath) & RsDelete("Image")
				ew_DeleteFile ew_UploadPathEx(True, Products.Image_Thumb.UploadPath) & RsDelete("Image_Thumb")
				ew_DeleteFile ew_UploadPathEx(True, Products.fImage.UploadPath) & RsDelete("fImage")
				ew_DeleteFile ew_UploadPathEx(True, Products.fImage_Thumb.UploadPath) & RsDelete("fImage_Thumb")
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
			If Products.CancelMessage <> "" Then
				FailureMessage = Products.CancelMessage
				Products.CancelMessage = ""
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
				Call Products.Row_Deleted(RsOld)
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
