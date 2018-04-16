<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Discountcodesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="DiscountTypesinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Discountcodes_search
Set Discountcodes_search = New cDiscountcodes_search
Set Page = Discountcodes_search

' Page init processing
Call Discountcodes_search.Page_Init()

' Page main processing
Call Discountcodes_search.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Discountcodes_search = new ew_Page("Discountcodes_search");
// page properties
Discountcodes_search.PageID = "search"; // page ID
Discountcodes_search.FormID = "fDiscountcodessearch"; // form ID
var EW_PAGE_ID = Discountcodes_search.PageID; // for backward compatibility
// extend page with validate function for search
Discountcodes_search.ValidateSearch = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (this.ValidateRequired) {
		var infix = "";
		elm = fobj.elements["x" + infix + "_Discountid"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.Discountid.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_OrderId"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.OrderId.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Use_date"];
		if (elm && !ew_CheckDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.Use_date.FldErrMsg) %>");
		// Call Form Custom Validate event
		if (!this.Form_CustomValidate(fobj)) return false;
	}
	for (var i=0;i<fobj.elements.length;i++) {
		var elem = fobj.elements[i];
		if (elem.name.substring(0,2) == "s_" || elem.name.substring(0,3) == "sv_")
			elem.value = "";
	}
	return true;
}
// extend page with Form_CustomValidate function
Discountcodes_search.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Discountcodes_search.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Discountcodes_search.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Discountcodes_search.ValidateRequired = false; // no JavaScript validation
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
<% Discountcodes_search.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Search") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Discountcodes.TableCaption %></p>
<p class="aspmaker"><a href="<%= Discountcodes.ReturnUrl %>"><%= Language.Phrase("BackToList") %></a></p>
<% Discountcodes_search.ShowMessage %>
<form name="fDiscountcodessearch" id="fDiscountcodessearch" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Discountcodes_search.ValidateSearch(this);">
<p>
<input type="hidden" name="t" id="t" value="Discountcodes">
<input type="hidden" name="a_search" id="a_search" value="S">
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
	<tr id="r_Discountid"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Discountid.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_Discountid" id="z_Discountid" value="="></td>
		<td<%= Discountcodes.Discountid.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<input type="text" name="x_Discountid" id="x_Discountid" value="<%= Discountcodes.Discountid.EditValue %>"<%= Discountcodes.Discountid.EditAttributes %>>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_DiscountCode"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountCode.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("LIKE") %><input type="hidden" name="z_DiscountCode" id="z_DiscountCode" value="LIKE"></td>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<input type="text" name="x_DiscountCode" id="x_DiscountCode" size="30" maxlength="6" value="<%= Discountcodes.DiscountCode.EditValue %>"<%= Discountcodes.DiscountCode.EditAttributes %>>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_Active"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Active.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_Active" id="z_Active" value="="></td>
		<td<%= Discountcodes.Active.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.Active.AdvancedSearch.SearchValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Active" id="x_Active" value="1"<%= selwrk %><%= Discountcodes.Active.EditAttributes %>>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_used"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.used.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_used" id="z_used" value="="></td>
		<td<%= Discountcodes.used.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.used.AdvancedSearch.SearchValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_used" id="x_used" value="1"<%= selwrk %><%= Discountcodes.used.EditAttributes %>>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_OrderId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.OrderId.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_OrderId" id="z_OrderId" value="="></td>
		<td<%= Discountcodes.OrderId.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<input type="text" name="x_OrderId" id="x_OrderId" size="30" value="<%= Discountcodes.OrderId.EditValue %>"<%= Discountcodes.OrderId.EditAttributes %>>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_Use_date"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.Use_date.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_Use_date" id="z_Use_date" value="="></td>
		<td<%= Discountcodes.Use_date.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<input type="text" name="x_Use_date" id="x_Use_date" value="<%= Discountcodes.Use_date.EditValue %>"<%= Discountcodes.Use_date.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x_Use_date" name="cal_x_Use_date" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x_Use_date", // input field id
	ifFormat: "%Y/%m/%d", // date format
	button: "cal_x_Use_date" // button id
});
</script>
</span>
			</div>
		</td>
	</tr>
	<tr id="r_DiscountTypeId"<%= Discountcodes.RowAttributes %>>
		<td class="ewTableHeader"><%= Discountcodes.DiscountTypeId.FldCaption %></td>
		<td class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_DiscountTypeId" id="z_DiscountTypeId" value="="></td>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
			<div style="white-space: nowrap;">
				<span class="aspmaker">
<select id="x_DiscountTypeId" name="x_DiscountTypeId"<%= Discountcodes.DiscountTypeId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Discountcodes.DiscountTypeId.EditValue) Then
	arwrk = Discountcodes.DiscountTypeId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue&"" Then
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
</span>
			</div>
		</td>
	</tr>
</table>
</div>
</td></tr></table>
<p>
<input type="submit" name="Action" id="Action" value="<%= ew_BtnCaption(Language.Phrase("Search")) %>">
<input type="button" name="Reset" id="Reset" value="<%= ew_BtnCaption(Language.Phrase("Reset")) %>" onclick="ew_ClearForm(this.form);">
</form>
<%
Discountcodes_search.ShowPageFooter()
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
Set Discountcodes_search = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountcodes_search

	' Page ID
	Public Property Get PageID()
		PageID = "search"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Discountcodes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Discountcodes_search"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Discountcodes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Discountcodes.TableVar & "&" ' add page token
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
		If Discountcodes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Discountcodes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Discountcodes.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes
		Set Table = Discountcodes

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(DiscountTypes) Then Set DiscountTypes = New cDiscountTypes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "search"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Discountcodes"

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
		Set Discountcodes = Nothing
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
	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		If IsPageRequest Then ' Validate request

			' Get action
			Discountcodes.CurrentAction = ObjForm.GetValue("a_search")
			Select Case Discountcodes.CurrentAction
				Case "S" ' Get Search Criteria

					' Build search string for advanced search, remove blank field
					Dim sSrchStr
					Call LoadSearchValues() ' Get search values
					If ValidateSearch() Then
						sSrchStr = BuildAdvancedSearch()
					Else
						sSrchStr = ""
						FailureMessage = gsSearchError
					End If
					If sSrchStr <> "" Then
						sSrchStr = Discountcodes.UrlParm(sSrchStr)
						Call Page_Terminate("Discountcodeslist.asp" & "?" & sSrchStr) ' Go to list page
					End If
			End Select
		End If

		' Restore search settings from Session
		If gsSearchError = "" Then
			Call LoadAdvancedSearch()
		End If

		' Render row for search
		Discountcodes.RowType = EW_ROWTYPE_SEARCH
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Build advanced search
	'
	Function BuildAdvancedSearch()
		Dim sSrchUrl
		sSrchUrl = ""

		' Field Discountid
		Call BuildSearchUrl(sSrchUrl, Discountcodes.Discountid)

		' Field DiscountCode
		Call BuildSearchUrl(sSrchUrl, Discountcodes.DiscountCode)

		' Field Active
		Call BuildSearchUrl(sSrchUrl, Discountcodes.Active)

		' Field used
		Call BuildSearchUrl(sSrchUrl, Discountcodes.used)

		' Field OrderId
		Call BuildSearchUrl(sSrchUrl, Discountcodes.OrderId)

		' Field Use_date
		Call BuildSearchUrl(sSrchUrl, Discountcodes.Use_date)

		' Field DiscountTypeId
		Call BuildSearchUrl(sSrchUrl, Discountcodes.DiscountTypeId)
		BuildAdvancedSearch = sSrchUrl
	End Function

	' -----------------------------------------------------------------
	' Function to build search URL
	'
	Sub BuildSearchUrl(Url, Fld)
		Dim FldVal, FldOpr, FldCond, FldVal2, FldOpr2
		Dim FldParm
		Dim IsValidValue, sWrk
		sWrk = ""
		FldParm = Mid(Fld.FldVar, 3)
		FldVal = ObjForm.GetValue("x_" & FldParm)
		FldOpr = ObjForm.GetValue("z_" & FldParm)
		FldCond = ObjForm.GetValue("v_" & FldParm)
		FldVal2 = ObjForm.GetValue("y_" & FldParm)
		FldOpr2 = ObjForm.GetValue("w_" & FldParm)
		FldOpr = UCase(Trim(FldOpr))
		Dim lFldDataType
		If Fld.FldIsVirtual Then
			lFldDataType = EW_DATATYPE_STRING
		Else
			lFldDataType = Fld.FldDataType
		End If
		If FldOpr = "BETWEEN" Then
			IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
				(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal) And IsNumeric(FldVal2))
			If FldVal <> "" And FldVal2 <> "" And IsValidValue Then
				sWrk = "x_" & FldParm & "=" & Server.URLEncode(FldVal) & _
					"&y_" & FldParm & "=" & Server.URLEncode(FldVal2) & _
					"&z_" & FldParm & "=" & Server.URLEncode(FldOpr)
			End If
		ElseIf FldOpr = "IS NULL" Or FldOpr = "IS NOT NULL" Then
			sWrk = "x_" & FldParm & "=" & Server.URLEncode(FldVal) & _
				"&z_" & FldParm & "=" & Server.URLEncode(FldOpr)
		Else
			IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
				(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal))
			If FldVal <> "" And IsValidValue And ew_IsValidOpr(FldOpr, lFldDataType) Then
				sWrk = "x_" & FldParm & "=" & Server.URLEncode(FldVal) & _
					"&z_" & FldParm & "=" & Server.URLEncode(FldOpr)
			End If
			IsValidValue = (lFldDataType <> EW_DATATYPE_NUMBER) Or _
				(lFldDataType = EW_DATATYPE_NUMBER And IsNumeric(FldVal2))
			If FldVal2 <> "" And IsValidValue And ew_IsValidOpr(FldOpr2, lFldDataType) Then
				If sWrk <> "" Then sWrk = sWrk & "&v_" & FldParm & "=" & FldCond & "&"
				sWrk = sWrk & "y_" & FldParm & "=" & Server.URLEncode(FldVal2) & _
					"&w_" & FldParm & "=" & Server.URLEncode(FldOpr2)
			End If
		End If
		If sWrk <> "" Then
			If Url <> "" Then Url = Url & "&"
			Url = Url & sWrk
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load search values for validation
	'
	Function LoadSearchValues()

		' Load search values
		Discountcodes.Discountid.AdvancedSearch.SearchValue = ObjForm.GetValue("x_Discountid")
		Discountcodes.Discountid.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_Discountid")
		Discountcodes.DiscountCode.AdvancedSearch.SearchValue = ObjForm.GetValue("x_DiscountCode")
		Discountcodes.DiscountCode.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_DiscountCode")
		Discountcodes.Active.AdvancedSearch.SearchValue = ObjForm.GetValue("x_Active")
		Discountcodes.Active.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_Active")
		Discountcodes.used.AdvancedSearch.SearchValue = ObjForm.GetValue("x_used")
		Discountcodes.used.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_used")
		Discountcodes.OrderId.AdvancedSearch.SearchValue = ObjForm.GetValue("x_OrderId")
		Discountcodes.OrderId.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_OrderId")
		Discountcodes.Use_date.AdvancedSearch.SearchValue = ObjForm.GetValue("x_Use_date")
		Discountcodes.Use_date.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_Use_date")
		Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue = ObjForm.GetValue("x_DiscountTypeId")
		Discountcodes.DiscountTypeId.AdvancedSearch.SearchOperator = ObjForm.GetValue("z_DiscountTypeId")
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Discountcodes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Discountid
		' DiscountCode
		' Active
		' used
		' OrderId
		' Use_date
		' DiscountTypeId
		' -----------
		'  View  Row
		' -----------

		If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Discountid
			Discountcodes.Discountid.ViewValue = Discountcodes.Discountid.CurrentValue
			Discountcodes.Discountid.ViewCustomAttributes = ""

			' DiscountCode
			Discountcodes.DiscountCode.ViewValue = Discountcodes.DiscountCode.CurrentValue
			Discountcodes.DiscountCode.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			Else
				Discountcodes.Active.ViewValue = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			End If
			Discountcodes.Active.ViewCustomAttributes = ""

			' used
			If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			Else
				Discountcodes.used.ViewValue = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			End If
			Discountcodes.used.ViewCustomAttributes = ""

			' OrderId
			Discountcodes.OrderId.ViewValue = Discountcodes.OrderId.CurrentValue
			Discountcodes.OrderId.ViewCustomAttributes = ""

			' Use_date
			Discountcodes.Use_date.ViewValue = Discountcodes.Use_date.CurrentValue
			Discountcodes.Use_date.ViewCustomAttributes = ""

			' DiscountTypeId
			If Discountcodes.DiscountTypeId.CurrentValue & "" <> "" Then
				sFilterWrk = "[DiscountTypeId] = " & ew_AdjustSql(Discountcodes.DiscountTypeId.CurrentValue) & ""
			sSqlWrk = "SELECT [DiscountType] FROM [DiscountTypes]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Discountcodes.DiscountTypeId.ViewValue = RsWrk("DiscountType")
				Else
					Discountcodes.DiscountTypeId.ViewValue = Discountcodes.DiscountTypeId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Discountcodes.DiscountTypeId.ViewValue = Null
			End If
			Discountcodes.DiscountTypeId.ViewCustomAttributes = ""

			' View refer script
			' Discountid

			Discountcodes.Discountid.LinkCustomAttributes = ""
			Discountcodes.Discountid.HrefValue = ""
			Discountcodes.Discountid.TooltipValue = ""

			' DiscountCode
			Discountcodes.DiscountCode.LinkCustomAttributes = ""
			Discountcodes.DiscountCode.HrefValue = ""
			Discountcodes.DiscountCode.TooltipValue = ""

			' Active
			Discountcodes.Active.LinkCustomAttributes = ""
			Discountcodes.Active.HrefValue = ""
			Discountcodes.Active.TooltipValue = ""

			' used
			Discountcodes.used.LinkCustomAttributes = ""
			Discountcodes.used.HrefValue = ""
			Discountcodes.used.TooltipValue = ""

			' OrderId
			Discountcodes.OrderId.LinkCustomAttributes = ""
			If Not ew_Empty(Discountcodes.OrderId.CurrentValue) Then
				Discountcodes.OrderId.HrefValue = "OrderDetailslist.asp?showmaster=Orders&OrderId=" & ew_IIf(Discountcodes.OrderId.ViewValue<>"", Discountcodes.OrderId.ViewValue, Discountcodes.OrderId.CurrentValue)
				Discountcodes.OrderId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Discountcodes.Export <> "" Then Discountcodes.OrderId.HrefValue = ew_ConvertFullUrl(Discountcodes.OrderId.HrefValue)
			Else
				Discountcodes.OrderId.HrefValue = ""
			End If
			Discountcodes.OrderId.TooltipValue = ""

			' Use_date
			Discountcodes.Use_date.LinkCustomAttributes = ""
			Discountcodes.Use_date.HrefValue = ""
			Discountcodes.Use_date.TooltipValue = ""

			' DiscountTypeId
			Discountcodes.DiscountTypeId.LinkCustomAttributes = ""
			Discountcodes.DiscountTypeId.HrefValue = ""
			Discountcodes.DiscountTypeId.TooltipValue = ""

		' ------------
		'  Search Row
		' ------------

		ElseIf Discountcodes.RowType = EW_ROWTYPE_SEARCH Then ' Search row

			' Discountid
			Discountcodes.Discountid.EditCustomAttributes = ""
			Discountcodes.Discountid.EditValue = ew_HtmlEncode(Discountcodes.Discountid.AdvancedSearch.SearchValue)

			' DiscountCode
			Discountcodes.DiscountCode.EditCustomAttributes = ""
			Discountcodes.DiscountCode.EditValue = ew_HtmlEncode(Discountcodes.DiscountCode.AdvancedSearch.SearchValue)

			' Active
			Discountcodes.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.Active.FldTagCaption(1) <> "", Discountcodes.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.Active.FldTagCaption(2) <> "", Discountcodes.Active.FldTagCaption(2), "No")
			Discountcodes.Active.EditValue = arwrk

			' used
			Discountcodes.used.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Discountcodes.used.FldTagCaption(1) <> "", Discountcodes.used.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Discountcodes.used.FldTagCaption(2) <> "", Discountcodes.used.FldTagCaption(2), "No")
			Discountcodes.used.EditValue = arwrk

			' OrderId
			Discountcodes.OrderId.EditCustomAttributes = ""
			Discountcodes.OrderId.EditValue = ew_HtmlEncode(Discountcodes.OrderId.AdvancedSearch.SearchValue)

			' Use_date
			Discountcodes.Use_date.EditCustomAttributes = ""
			Discountcodes.Use_date.EditValue = Discountcodes.Use_date.AdvancedSearch.SearchValue

			' DiscountTypeId
			Discountcodes.DiscountTypeId.EditCustomAttributes = ""
				sFilterWrk = ""
			sSqlWrk = "SELECT [DiscountTypeId], [DiscountType] AS [DispFld], '' AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [DiscountTypes]"
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
			Discountcodes.DiscountTypeId.EditValue = arwrk
		End If
		If Discountcodes.RowType = EW_ROWTYPE_ADD Or Discountcodes.RowType = EW_ROWTYPE_EDIT Or Discountcodes.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Discountcodes.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Discountcodes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Discountcodes.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate search
	'
	Function ValidateSearch()

		' Initialize
		gsSearchError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateSearch = True
			Exit Function
		End If
		If Not ew_CheckInteger(Discountcodes.Discountid.AdvancedSearch.SearchValue) Then
			Call ew_AddMessage(gsSearchError, Discountcodes.Discountid.FldErrMsg)
		End If
		If Not ew_CheckInteger(Discountcodes.OrderId.AdvancedSearch.SearchValue) Then
			Call ew_AddMessage(gsSearchError, Discountcodes.OrderId.FldErrMsg)
		End If
		If Not ew_CheckDate(Discountcodes.Use_date.AdvancedSearch.SearchValue) Then
			Call ew_AddMessage(gsSearchError, Discountcodes.Use_date.FldErrMsg)
		End If

		' Return validate result
		ValidateSearch = (gsSearchError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateSearch = ValidateSearch And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsSearchError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Load advanced search
	'
	Function LoadAdvancedSearch()
		Discountcodes.Discountid.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_Discountid")
		Discountcodes.DiscountCode.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_DiscountCode")
		Discountcodes.Active.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_Active")
		Discountcodes.used.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_used")
		Discountcodes.OrderId.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_OrderId")
		Discountcodes.Use_date.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_Use_date")
		Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_DiscountTypeId")
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
