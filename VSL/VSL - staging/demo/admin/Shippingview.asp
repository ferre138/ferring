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
Dim Shipping_view
Set Shipping_view = New cShipping_view
Set Page = Shipping_view

' Page init processing
Call Shipping_view.Page_Init()

' Page main processing
Call Shipping_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Shipping.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Shipping_view = new ew_Page("Shipping_view");
// page properties
Shipping_view.PageID = "view"; // page ID
Shipping_view.FormID = "fShippingview"; // form ID
var EW_PAGE_ID = Shipping_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Shipping_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Shipping_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Shipping_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Shipping_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Shipping.TableCaption %>
&nbsp;&nbsp;<% Shipping_view.ExportOptions.Render "body", "" %>
</p>
<% If Shipping.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Shipping_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Shipping_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Shipping_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Shipping_view.CopyUrl %>"><%= Language.Phrase("ViewPageCopyLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Shipping_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% End If %>
</p>
<% Shipping_view.ShowMessage %>
<p>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Shipping.AddressID.Visible Then ' AddressID %>
	<tr id="r_AddressID"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.AddressID.FldCaption %></td>
		<td<%= Shipping.AddressID.CellAttributes %>>
<div<%= Shipping.AddressID.ViewAttributes %>><%= Shipping.AddressID.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.CustomerId.Visible Then ' CustomerId %>
	<tr id="r_CustomerId"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.CustomerId.FldCaption %></td>
		<td<%= Shipping.CustomerId.CellAttributes %>>
<div<%= Shipping.CustomerId.ViewAttributes %>><%= Shipping.CustomerId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_FirstName.Visible Then ' ship_FirstName %>
	<tr id="r_ship_FirstName"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_FirstName.FldCaption %></td>
		<td<%= Shipping.ship_FirstName.CellAttributes %>>
<div<%= Shipping.ship_FirstName.ViewAttributes %>><%= Shipping.ship_FirstName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_LastName.Visible Then ' ship_LastName %>
	<tr id="r_ship_LastName"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_LastName.FldCaption %></td>
		<td<%= Shipping.ship_LastName.CellAttributes %>>
<div<%= Shipping.ship_LastName.ViewAttributes %>><%= Shipping.ship_LastName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_Address.Visible Then ' ship_Address %>
	<tr id="r_ship_Address"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Address.FldCaption %></td>
		<td<%= Shipping.ship_Address.CellAttributes %>>
<div<%= Shipping.ship_Address.ViewAttributes %>><%= Shipping.ship_Address.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_City.Visible Then ' ship_City %>
	<tr id="r_ship_City"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_City.FldCaption %></td>
		<td<%= Shipping.ship_City.CellAttributes %>>
<div<%= Shipping.ship_City.ViewAttributes %>><%= Shipping.ship_City.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_Province.Visible Then ' ship_Province %>
	<tr id="r_ship_Province"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Province.FldCaption %></td>
		<td<%= Shipping.ship_Province.CellAttributes %>>
<div<%= Shipping.ship_Province.ViewAttributes %>><%= Shipping.ship_Province.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_PostalCode.Visible Then ' ship_PostalCode %>
	<tr id="r_ship_PostalCode"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_PostalCode.FldCaption %></td>
		<td<%= Shipping.ship_PostalCode.CellAttributes %>>
<div<%= Shipping.ship_PostalCode.ViewAttributes %>><%= Shipping.ship_PostalCode.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_Country.Visible Then ' ship_Country %>
	<tr id="r_ship_Country"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Country.FldCaption %></td>
		<td<%= Shipping.ship_Country.CellAttributes %>>
<div<%= Shipping.ship_Country.ViewAttributes %>><%= Shipping.ship_Country.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
	<tr id="r_ship_EmailAddress"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_EmailAddress.FldCaption %></td>
		<td<%= Shipping.ship_EmailAddress.CellAttributes %>>
<div<%= Shipping.ship_EmailAddress.ViewAttributes %>><%= Shipping.ship_EmailAddress.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.HomePhone.Visible Then ' HomePhone %>
	<tr id="r_HomePhone"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.HomePhone.FldCaption %></td>
		<td<%= Shipping.HomePhone.CellAttributes %>>
<div<%= Shipping.HomePhone.ViewAttributes %>><%= Shipping.HomePhone.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.WorkPhone.Visible Then ' WorkPhone %>
	<tr id="r_WorkPhone"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.WorkPhone.FldCaption %></td>
		<td<%= Shipping.WorkPhone.CellAttributes %>>
<div<%= Shipping.WorkPhone.ViewAttributes %>><%= Shipping.WorkPhone.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Shipping.ship_Address2.Visible Then ' ship_Address2 %>
	<tr id="r_ship_Address2"<%= Shipping.RowAttributes %>>
		<td class="ewTableHeader"><%= Shipping.ship_Address2.FldCaption %></td>
		<td<%= Shipping.ship_Address2.CellAttributes %>>
<div<%= Shipping.ship_Address2.ViewAttributes %>><%= Shipping.ship_Address2.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
<p>
<%
Shipping_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Shipping.Export = "" Then %>
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
Set Shipping_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cShipping_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Shipping"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Shipping_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("AddressID").Count > 0 Then
			ew_AddKey RecKey, "AddressID", Request.QueryString("AddressID")
			KeyUrl = KeyUrl & "&AddressID=" & Server.URLEncode(Request.QueryString("AddressID"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Shipping"

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
			If Request.QueryString("AddressID").Count > 0 Then
				Shipping.AddressID.QueryStringValue = Request.QueryString("AddressID")
			Else
				sReturnUrl = "Shippinglist.asp" ' Return to list
			End If

			' Get action
			Shipping.CurrentAction = "I" ' Display form
			Select Case Shipping.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Shippinglist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Shippinglist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Shipping.RowType = EW_ROWTYPE_VIEW
		Call Shipping.ResetAttrs()
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
				Shipping.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Shipping.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Shipping.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Shipping.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Shipping.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Shipping.StartRecordNumber = StartRec
		End If
	End Sub

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
		AddUrl = Shipping.AddUrl
		EditUrl = Shipping.EditUrl("")
		CopyUrl = Shipping.CopyUrl("")
		DeleteUrl = Shipping.DeleteUrl
		ListUrl = Shipping.ListUrl

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
		End If

		' Call Row Rendered event
		If Shipping.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Shipping.Row_Rendered()
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
