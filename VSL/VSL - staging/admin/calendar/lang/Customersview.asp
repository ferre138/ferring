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
Dim Customers_view
Set Customers_view = New cCustomers_view
Set Page = Customers_view

' Page init processing
Call Customers_view.Page_Init()

' Page main processing
Call Customers_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Customers.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Customers_view = new ew_Page("Customers_view");
// page properties
Customers_view.PageID = "view"; // page ID
Customers_view.FormID = "fCustomersview"; // form ID
var EW_PAGE_ID = Customers_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Customers_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Customers_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Customers_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Customers_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<div id="ewDetailsDiv" style="visibility: hidden; z-index: 11000;" name="ewDetailsDivDiv"></div>
<script language="JavaScript" type="text/javascript">
<!--
// YUI container
var ewDetailsDiv;
var ew_AjaxDetailsTimer = null;
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Customers_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Customers.TableCaption %>
&nbsp;&nbsp;<% Customers_view.ExportOptions.Render "body", "" %>
</p>
<% If Customers.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Customers_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Customers_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Customers_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Customers_view.CopyUrl %>"><%= Language.Phrase("ViewPageCopyLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Customers_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Customers_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Customers.CustomerID.Visible Then ' CustomerID %>
	<tr id="r_CustomerID"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.CustomerID.FldCaption %></td>
		<td<%= Customers.CustomerID.CellAttributes %>>
<div<%= Customers.CustomerID.ViewAttributes %>><%= Customers.CustomerID.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.Inv_FirstName.Visible Then ' Inv_FirstName %>
	<tr id="r_Inv_FirstName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_FirstName.FldCaption %></td>
		<td<%= Customers.Inv_FirstName.CellAttributes %>>
<div<%= Customers.Inv_FirstName.ViewAttributes %>><%= Customers.Inv_FirstName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.Inv_LastName.Visible Then ' Inv_LastName %>
	<tr id="r_Inv_LastName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_LastName.FldCaption %></td>
		<td<%= Customers.Inv_LastName.CellAttributes %>>
<div<%= Customers.Inv_LastName.ViewAttributes %>><%= Customers.Inv_LastName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.Inv_Address.Visible Then ' Inv_Address %>
	<tr id="r_Inv_Address"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_Address.FldCaption %></td>
		<td<%= Customers.Inv_Address.CellAttributes %>>
<div<%= Customers.Inv_Address.ViewAttributes %>><%= Customers.Inv_Address.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_City.Visible Then ' inv_City %>
	<tr id="r_inv_City"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_City.FldCaption %></td>
		<td<%= Customers.inv_City.CellAttributes %>>
<div<%= Customers.inv_City.ViewAttributes %>><%= Customers.inv_City.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_Province.Visible Then ' inv_Province %>
	<tr id="r_inv_Province"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Province.FldCaption %></td>
		<td<%= Customers.inv_Province.CellAttributes %>>
<div<%= Customers.inv_Province.ViewAttributes %>><%= Customers.inv_Province.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_PostalCode.Visible Then ' inv_PostalCode %>
	<tr id="r_inv_PostalCode"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_PostalCode.FldCaption %></td>
		<td<%= Customers.inv_PostalCode.CellAttributes %>>
<div<%= Customers.inv_PostalCode.ViewAttributes %>><%= Customers.inv_PostalCode.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_Country.Visible Then ' inv_Country %>
	<tr id="r_inv_Country"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Country.FldCaption %></td>
		<td<%= Customers.inv_Country.CellAttributes %>>
<div<%= Customers.inv_Country.ViewAttributes %>><%= Customers.inv_Country.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_PhoneNumber.Visible Then ' inv_PhoneNumber %>
	<tr id="r_inv_PhoneNumber"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_PhoneNumber.FldCaption %></td>
		<td<%= Customers.inv_PhoneNumber.CellAttributes %>>
<div<%= Customers.inv_PhoneNumber.ViewAttributes %>><%= Customers.inv_PhoneNumber.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
	<tr id="r_inv_EmailAddress"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_EmailAddress.FldCaption %></td>
		<td<%= Customers.inv_EmailAddress.CellAttributes %>>
<div<%= Customers.inv_EmailAddress.ViewAttributes %>><%= Customers.inv_EmailAddress.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.Notes.Visible Then ' Notes %>
	<tr id="r_Notes"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Notes.FldCaption %></td>
		<td<%= Customers.Notes.CellAttributes %>>
<div<%= Customers.Notes.ViewAttributes %>><%= Customers.Notes.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.inv_Fax.Visible Then ' inv_Fax %>
	<tr id="r_inv_Fax"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.inv_Fax.FldCaption %></td>
		<td<%= Customers.inv_Fax.CellAttributes %>>
<div<%= Customers.inv_Fax.ViewAttributes %>><%= Customers.inv_Fax.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.Inv_Address2.Visible Then ' Inv_Address2 %>
	<tr id="r_Inv_Address2"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.Inv_Address2.FldCaption %></td>
		<td<%= Customers.Inv_Address2.CellAttributes %>>
<div<%= Customers.Inv_Address2.ViewAttributes %>><%= Customers.Inv_Address2.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.UserName.Visible Then ' UserName %>
	<tr id="r_UserName"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.UserName.FldCaption %></td>
		<td<%= Customers.UserName.CellAttributes %>>
<div<%= Customers.UserName.ViewAttributes %>><%= Customers.UserName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.passwrd.Visible Then ' passwrd %>
	<tr id="r_passwrd"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.passwrd.FldCaption %></td>
		<td<%= Customers.passwrd.CellAttributes %>>
<div<%= Customers.passwrd.ViewAttributes %>><%= Customers.passwrd.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Customers.NewCustomer.Visible Then ' NewCustomer %>
	<tr id="r_NewCustomer"<%= Customers.RowAttributes %>>
		<td class="ewTableHeader"><%= Customers.NewCustomer.FldCaption %></td>
		<td<%= Customers.NewCustomer.CellAttributes %>>
<% If ew_ConvertToBool(Customers.NewCustomer.CurrentValue) Then %>
<input type="checkbox" value="<%= Customers.NewCustomer.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Customers.NewCustomer.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Customers_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Customers.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Customers_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cCustomers_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Customers"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Customers_view"
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

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("CustomerID").Count > 0 Then
			ew_AddKey RecKey, "CustomerID", Request.QueryString("CustomerID")
			KeyUrl = KeyUrl & "&CustomerID=" & Server.URLEncode(Request.QueryString("CustomerID"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Customers"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.Tag = "span"
		ExportOptions.Separator = "&nbsp;&nbsp;"
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

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False
		If IsPageRequest Then ' Validate request
			If Request.QueryString("CustomerID").Count > 0 Then
				Customers.CustomerID.QueryStringValue = Request.QueryString("CustomerID")
			Else
				sReturnUrl = "Customerslist.asp" ' Return to list
			End If

			' Get action
			Customers.CurrentAction = "I" ' Display form
			Select Case Customers.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Customerslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Customerslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Customers.RowType = EW_ROWTYPE_VIEW
		Call Customers.ResetAttrs()
		Call RenderRow()
	End Sub
	Dim Pager

	' -----------------------------------------------------------------
	' Set up Starting Record parameters based on Pager Navigation
	'
	Sub SetUpStartRec()
		Dim PageNo

		' Exit if DisplayRecs = 0
		If DisplayRecs = 0 Then Exit Sub
		If IsPageRequest Then ' Validate request

			' Check for a START parameter
			If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
				StartRec = Request.QueryString(EW_TABLE_START_REC)
				Customers.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Customers.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Customers.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Customers.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Customers.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Customers.StartRecordNumber = StartRec
		End If
	End Sub

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
		AddUrl = Customers.AddUrl
		EditUrl = Customers.EditUrl("")
		CopyUrl = Customers.CopyUrl("")
		DeleteUrl = Customers.DeleteUrl
		ListUrl = Customers.ListUrl

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
			' CustomerID

			Customers.CustomerID.LinkCustomAttributes = ""
			Customers.CustomerID.HrefValue = ""
			Customers.CustomerID.TooltipValue = ""

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
		End If

		' Call Row Rendered event
		If Customers.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Customers.Row_Rendered()
		End If
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
End Class
%>
