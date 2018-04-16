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
Dim OrderDetails_edit
Set OrderDetails_edit = New cOrderDetails_edit
Set Page = OrderDetails_edit

' Page init processing
Call OrderDetails_edit.Page_Init()

' Page main processing
Call OrderDetails_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var OrderDetails_edit = new ew_Page("OrderDetails_edit");
// page properties
OrderDetails_edit.PageID = "edit"; // page ID
OrderDetails_edit.FormID = "fOrderDetailsedit"; // form ID
var EW_PAGE_ID = OrderDetails_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
OrderDetails_edit.ValidateForm = function(fobj) {
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
			return ew_OnError(this, elm, "<%= ew_JsEncode2(OrderDetails.OrderId.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Quantity"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(OrderDetails.Quantity.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Price"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(OrderDetails.Price.FldErrMsg) %>");
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
OrderDetails_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
OrderDetails_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
OrderDetails_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
OrderDetails_edit.ValidateRequired = false; // no JavaScript validation
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
<% OrderDetails_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= OrderDetails.TableCaption %></p>
<p class="aspmaker"><a href="<%= OrderDetails.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% OrderDetails_edit.ShowMessage %>
<form name="fOrderDetailsedit" id="fOrderDetailsedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return OrderDetails_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="OrderDetails">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If OrderDetails.OrderDetailsId.Visible Then ' OrderDetailsId %>
	<tr id="r_OrderDetailsId"<%= OrderDetails.RowAttributes %>>
		<td class="ewTableHeader"><%= OrderDetails.OrderDetailsId.FldCaption %></td>
		<td<%= OrderDetails.OrderDetailsId.CellAttributes %>><span id="el_OrderDetailsId">
<div<%= OrderDetails.OrderDetailsId.ViewAttributes %>><%= OrderDetails.OrderDetailsId.EditValue %></div>
<input type="hidden" name="x_OrderDetailsId" id="x_OrderDetailsId" value="<%= Server.HTMLEncode(OrderDetails.OrderDetailsId.CurrentValue&"") %>">
</span><%= OrderDetails.OrderDetailsId.CustomMsg %></td>
	</tr>
<% End If %>
<% If OrderDetails.OrderId.Visible Then ' OrderId %>
	<tr id="r_OrderId"<%= OrderDetails.RowAttributes %>>
		<td class="ewTableHeader"><%= OrderDetails.OrderId.FldCaption %></td>
		<td<%= OrderDetails.OrderId.CellAttributes %>><span id="el_OrderId">
<% If OrderDetails.OrderId.SessionValue <> "" Then %>
<div<%= OrderDetails.OrderId.ViewAttributes %>><%= OrderDetails.OrderId.ViewValue %></div>
<input type="hidden" id="x_OrderId" name="x_OrderId" value="<%= ew_HtmlEncode(OrderDetails.OrderId.CurrentValue) %>">
<% Else %>
<input type="text" name="x_OrderId" id="x_OrderId" size="30" value="<%= OrderDetails.OrderId.EditValue %>"<%= OrderDetails.OrderId.EditAttributes %>>
<% End If %>
</span><%= OrderDetails.OrderId.CustomMsg %></td>
	</tr>
<% End If %>
<% If OrderDetails.ProductId.Visible Then ' ProductId %>
	<tr id="r_ProductId"<%= OrderDetails.RowAttributes %>>
		<td class="ewTableHeader"><%= OrderDetails.ProductId.FldCaption %></td>
		<td<%= OrderDetails.ProductId.CellAttributes %>><span id="el_ProductId">
<select id="x_ProductId" name="x_ProductId"<%= OrderDetails.ProductId.EditAttributes %>>
<%
emptywrk = True
If IsArray(OrderDetails.ProductId.EditValue) Then
	arwrk = OrderDetails.ProductId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = OrderDetails.ProductId.CurrentValue&"" Then
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
</span><%= OrderDetails.ProductId.CustomMsg %></td>
	</tr>
<% End If %>
<% If OrderDetails.Quantity.Visible Then ' Quantity %>
	<tr id="r_Quantity"<%= OrderDetails.RowAttributes %>>
		<td class="ewTableHeader"><%= OrderDetails.Quantity.FldCaption %></td>
		<td<%= OrderDetails.Quantity.CellAttributes %>><span id="el_Quantity">
<input type="text" name="x_Quantity" id="x_Quantity" size="30" value="<%= OrderDetails.Quantity.EditValue %>"<%= OrderDetails.Quantity.EditAttributes %>>
</span><%= OrderDetails.Quantity.CustomMsg %></td>
	</tr>
<% End If %>
<% If OrderDetails.Price.Visible Then ' Price %>
	<tr id="r_Price"<%= OrderDetails.RowAttributes %>>
		<td class="ewTableHeader"><%= OrderDetails.Price.FldCaption %></td>
		<td<%= OrderDetails.Price.CellAttributes %>><span id="el_Price">
<input type="text" name="x_Price" id="x_Price" size="30" value="<%= OrderDetails.Price.EditValue %>"<%= OrderDetails.Price.EditAttributes %>>
</span><%= OrderDetails.Price.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
OrderDetails_edit.ShowPageFooter()
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
Set OrderDetails_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrderDetails_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "OrderDetails"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "OrderDetails_edit"
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
		EW_PAGE_ID = "edit"

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("OrderDetailsId").Count > 0 Then
			OrderDetails.OrderDetailsId.QueryStringValue = Request.QueryString("OrderDetailsId")
		End If

		' Set up master detail parameters
		SetUpMasterParms()
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			OrderDetails.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Validate Form
			If Not ValidateForm() Then
				OrderDetails.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				OrderDetails.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			OrderDetails.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If OrderDetails.OrderDetailsId.CurrentValue = "" Then Call Page_Terminate("OrderDetailslist.asp") ' Invalid key, return to list
		Select Case OrderDetails.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("OrderDetailslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				OrderDetails.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					sReturnUrl = OrderDetails.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					OrderDetails.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		OrderDetails.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call OrderDetails.ResetAttrs()
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
		If Not OrderDetails.OrderDetailsId.FldIsDetailKey Then OrderDetails.OrderDetailsId.FormValue = ObjForm.GetValue("x_OrderDetailsId")
		If Not OrderDetails.OrderId.FldIsDetailKey Then OrderDetails.OrderId.FormValue = ObjForm.GetValue("x_OrderId")
		If Not OrderDetails.ProductId.FldIsDetailKey Then OrderDetails.ProductId.FormValue = ObjForm.GetValue("x_ProductId")
		If Not OrderDetails.Quantity.FldIsDetailKey Then OrderDetails.Quantity.FormValue = ObjForm.GetValue("x_Quantity")
		If Not OrderDetails.Price.FldIsDetailKey Then OrderDetails.Price.FormValue = ObjForm.GetValue("x_Price")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		OrderDetails.OrderDetailsId.CurrentValue = OrderDetails.OrderDetailsId.FormValue
		OrderDetails.OrderId.CurrentValue = OrderDetails.OrderId.FormValue
		OrderDetails.ProductId.CurrentValue = OrderDetails.ProductId.FormValue
		OrderDetails.Quantity.CurrentValue = OrderDetails.Quantity.FormValue
		OrderDetails.Price.CurrentValue = OrderDetails.Price.FormValue
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
			' OrderDetailsId

			OrderDetails.OrderDetailsId.LinkCustomAttributes = ""
			OrderDetails.OrderDetailsId.HrefValue = ""
			OrderDetails.OrderDetailsId.TooltipValue = ""

			' OrderId
			OrderDetails.OrderId.LinkCustomAttributes = ""
			OrderDetails.OrderId.HrefValue = ""
			OrderDetails.OrderId.TooltipValue = ""

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

		' ----------
		'  Edit Row
		' ----------

		ElseIf OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' OrderDetailsId
			OrderDetails.OrderDetailsId.EditCustomAttributes = ""
			OrderDetails.OrderDetailsId.EditValue = OrderDetails.OrderDetailsId.CurrentValue
			OrderDetails.OrderDetailsId.ViewCustomAttributes = ""

			' OrderId
			OrderDetails.OrderId.EditCustomAttributes = ""
			If OrderDetails.OrderId.SessionValue <> "" Then
				OrderDetails.OrderId.CurrentValue = OrderDetails.OrderId.SessionValue
			OrderDetails.OrderId.ViewValue = OrderDetails.OrderId.CurrentValue
			OrderDetails.OrderId.ViewCustomAttributes = ""
			Else
			OrderDetails.OrderId.EditValue = ew_HtmlEncode(OrderDetails.OrderId.CurrentValue)
			End If

			' ProductId
			OrderDetails.ProductId.EditCustomAttributes = ""
				sFilterWrk = ""
			sSqlWrk = "SELECT [ItemId], [Description] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [Products]"
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
			OrderDetails.ProductId.EditValue = arwrk

			' Quantity
			OrderDetails.Quantity.EditCustomAttributes = ""
			OrderDetails.Quantity.EditValue = ew_HtmlEncode(OrderDetails.Quantity.CurrentValue)

			' Price
			OrderDetails.Price.EditCustomAttributes = ""
			OrderDetails.Price.EditValue = ew_HtmlEncode(OrderDetails.Price.CurrentValue)

			' Edit refer script
			' OrderDetailsId

			OrderDetails.OrderDetailsId.HrefValue = ""

			' OrderId
			OrderDetails.OrderId.HrefValue = ""

			' ProductId
			OrderDetails.ProductId.HrefValue = ""

			' Quantity
			OrderDetails.Quantity.HrefValue = ""

			' Price
			OrderDetails.Price.HrefValue = ""
		End If
		If OrderDetails.RowType = EW_ROWTYPE_ADD Or OrderDetails.RowType = EW_ROWTYPE_EDIT Or OrderDetails.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call OrderDetails.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If OrderDetails.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call OrderDetails.Row_Rendered()
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
		If Not ew_CheckInteger(OrderDetails.OrderId.FormValue) Then
			Call ew_AddMessage(gsFormError, OrderDetails.OrderId.FldErrMsg)
		End If
		If Not ew_CheckInteger(OrderDetails.Quantity.FormValue) Then
			Call ew_AddMessage(gsFormError, OrderDetails.Quantity.FldErrMsg)
		End If
		If Not ew_CheckInteger(OrderDetails.Price.FormValue) Then
			Call ew_AddMessage(gsFormError, OrderDetails.Price.FldErrMsg)
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
		sFilter = OrderDetails.KeyFilter
		OrderDetails.CurrentFilter  = sFilter
		sSql = OrderDetails.SQL
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

			' Field OrderId
			Call OrderDetails.OrderId.SetDbValue(Rs, OrderDetails.OrderId.CurrentValue, Null, OrderDetails.OrderId.ReadOnly)

			' Field ProductId
			Call OrderDetails.ProductId.SetDbValue(Rs, OrderDetails.ProductId.CurrentValue, Null, OrderDetails.ProductId.ReadOnly)

			' Field Quantity
			Call OrderDetails.Quantity.SetDbValue(Rs, OrderDetails.Quantity.CurrentValue, Null, OrderDetails.Quantity.ReadOnly)

			' Field Price
			Call OrderDetails.Price.SetDbValue(Rs, OrderDetails.Price.CurrentValue, Null, OrderDetails.Price.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = OrderDetails.Row_Updating(RsOld, Rs)
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
				If OrderDetails.CancelMessage <> "" Then
					FailureMessage = OrderDetails.CancelMessage
					OrderDetails.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call OrderDetails.Row_Updated(RsOld, RsNew)
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
			If sMasterTblVar = "Orders" Then
				bValidMaster = True
				If Request.QueryString("OrderId").Count > 0 Then
					Orders.OrderId.QueryStringValue = Request.QueryString("OrderId")
					OrderDetails.OrderId.QueryStringValue = Orders.OrderId.QueryStringValue
					OrderDetails.OrderId.SessionValue = OrderDetails.OrderId.QueryStringValue
					If Not IsNumeric(Orders.OrderId.QueryStringValue) Then bValidMaster = False
				Else
					bValidMaster = False
				End If
			End If
		End If
		If bValidMaster Then

			' Save current master table
			OrderDetails.CurrentMasterTable = sMasterTblVar

			' Reset start record counter (new master key)
			StartRec = 1
			OrderDetails.StartRecordNumber = StartRec

			' Clear previous master session values
			If sMasterTblVar <> "Orders" Then
				If OrderDetails.OrderId.QueryStringValue = "" Then OrderDetails.OrderId.SessionValue = ""
			End If
		End If
		DbMasterFilter = OrderDetails.MasterFilter '  Get master filter
		DbDetailFilter = OrderDetails.DetailFilter ' Get detail filter
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
