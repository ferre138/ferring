<%@ CodePage="65001" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Ordersinfo.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, "utf-8") %>
<%

' Define page object
Dim Orders_preview
Set Orders_preview = New cOrders_preview
Set Page = Orders_preview

' Page init processing
Call Orders_preview.Page_Init()

' Page main processing
Call Orders_preview.Page_Main()
%>
<link href="css/vslpaypal.css" rel="stylesheet" type="text/css">
<% Orders_preview.ShowPageHeader() %>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Orders.TableCaption %>
<% If Orders_preview.TotalRecs > 0 Then %>
(<%= Orders_preview.TotalRecs %>&nbsp;<%= Language.Phrase("Record") %>)
<% Else %>
(<%= Language.Phrase("NoRecord") %>)
<% End If %>
</p>
<% If Orders_preview.TotalRecs > 0 Then %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="ewDetailsPreviewTable" name="ewDetailsPreviewTable" cellspacing="0" class="ewTable ewTableSeparate">
	<thead><!-- Table header -->
		<tr class="ewTableHeader">
<% If Orders.PromoCodeUsed.Visible Then ' OrderId %>
			<td valign="top"><%= Orders.OrderId.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' CustomerId %>
			<td valign="top"><%= Orders.CustomerId.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Amount %>
			<td valign="top"><%= Orders.Amount.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Ship_FirstName %>
			<td valign="top"><%= Orders.Ship_FirstName.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Ship_LastName %>
			<td valign="top"><%= Orders.Ship_LastName.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' payment_status %>
			<td valign="top"><%= Orders.payment_status.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Ordered_Date %>
			<td valign="top"><%= Orders.Ordered_Date.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' payer_email %>
			<td valign="top"><%= Orders.payer_email.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' payment_gross %>
			<td valign="top"><%= Orders.payment_gross.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' payment_fee %>
			<td valign="top"><%= Orders.payment_fee.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Tax %>
			<td valign="top"><%= Orders.Tax.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' Shipping %>
			<td valign="top"><%= Orders.Shipping.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' EmailSent %>
			<td valign="top"><%= Orders.EmailSent.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' EmailDate %>
			<td valign="top"><%= Orders.EmailDate.FldCaption %></td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' PromoCodeUsed %>
			<td valign="top"><%= Orders.PromoCodeUsed.FldCaption %></td>
<% End If %>
		</tr>
	</thead>
	<tbody><!-- Table body -->
<%
Orders_preview.RecCount = 0
Orders_preview.RowCnt = 0

'Orders_preview.Recordset.MoveFirst
Do While (Not Orders_preview.Recordset.Eof)

	' Init row class and style
	Orders_preview.RecCount = Orders_preview.RecCount + 1
	Orders_preview.RowCnt = Orders_preview.RowCnt + 1
	Orders.CssClass = ""
	Orders.CssStyle = ""
	Call Orders.LoadListRowValues(Orders_preview.Recordset)

	' Render row
	Orders.RowType = EW_ROWTYPE_PREVIEW ' Preview record
	Call Orders.RenderListRow()
%>
	<tr<%= Orders.RowAttributes %>>
<% If Orders.OrderId.Visible Then ' OrderId %>
		<!-- OrderId -->
		<td<%= Orders.OrderId.CellAttributes %>>
<div<%= Orders.OrderId.ViewAttributes %>><%= Orders.OrderId.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.CustomerId.Visible Then ' CustomerId %>
		<!-- CustomerId -->
		<td<%= Orders.CustomerId.CellAttributes %>>
<div<%= Orders.CustomerId.ViewAttributes %>>
<% If Orders.CustomerId.LinkAttributes <> "" Then %>
<a<%= Orders.CustomerId.LinkAttributes %>><%= Orders.CustomerId.ListViewValue %></a>
<% Else %>
<%= Orders.CustomerId.ListViewValue %>
<% End If %>
</div>
</td>
<% End If %>
<% If Orders.Amount.Visible Then ' Amount %>
		<!-- Amount -->
		<td<%= Orders.Amount.CellAttributes %>>
<div<%= Orders.Amount.ViewAttributes %>><%= Orders.Amount.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.Ship_FirstName.Visible Then ' Ship_FirstName %>
		<!-- Ship_FirstName -->
		<td<%= Orders.Ship_FirstName.CellAttributes %>>
<div<%= Orders.Ship_FirstName.ViewAttributes %>><%= Orders.Ship_FirstName.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.Ship_LastName.Visible Then ' Ship_LastName %>
		<!-- Ship_LastName -->
		<td<%= Orders.Ship_LastName.CellAttributes %>>
<div<%= Orders.Ship_LastName.ViewAttributes %>><%= Orders.Ship_LastName.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.payment_status.Visible Then ' payment_status %>
		<!-- payment_status -->
		<td<%= Orders.payment_status.CellAttributes %>>
<div<%= Orders.payment_status.ViewAttributes %>><%= Orders.payment_status.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.Ordered_Date.Visible Then ' Ordered_Date %>
		<!-- Ordered_Date -->
		<td<%= Orders.Ordered_Date.CellAttributes %>>
<div<%= Orders.Ordered_Date.ViewAttributes %>><%= Orders.Ordered_Date.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.payer_email.Visible Then ' payer_email %>
		<!-- payer_email -->
		<td<%= Orders.payer_email.CellAttributes %>>
<div<%= Orders.payer_email.ViewAttributes %>><%= Orders.payer_email.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.payment_gross.Visible Then ' payment_gross %>
		<!-- payment_gross -->
		<td<%= Orders.payment_gross.CellAttributes %>>
<div<%= Orders.payment_gross.ViewAttributes %>><%= Orders.payment_gross.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.payment_fee.Visible Then ' payment_fee %>
		<!-- payment_fee -->
		<td<%= Orders.payment_fee.CellAttributes %>>
<div<%= Orders.payment_fee.ViewAttributes %>><%= Orders.payment_fee.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.Tax.Visible Then ' Tax %>
		<!-- Tax -->
		<td<%= Orders.Tax.CellAttributes %>>
<div<%= Orders.Tax.ViewAttributes %>><%= Orders.Tax.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.Shipping.Visible Then ' Shipping %>
		<!-- Shipping -->
		<td<%= Orders.Shipping.CellAttributes %>>
<div<%= Orders.Shipping.ViewAttributes %>><%= Orders.Shipping.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.EmailSent.Visible Then ' EmailSent %>
		<!-- EmailSent -->
		<td<%= Orders.EmailSent.CellAttributes %>>
<div<%= Orders.EmailSent.ViewAttributes %>><%= Orders.EmailSent.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.EmailDate.Visible Then ' EmailDate %>
		<!-- EmailDate -->
		<td<%= Orders.EmailDate.CellAttributes %>>
<div<%= Orders.EmailDate.ViewAttributes %>><%= Orders.EmailDate.ListViewValue %></div>
</td>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' PromoCodeUsed %>
		<!-- PromoCodeUsed -->
		<td<%= Orders.PromoCodeUsed.CellAttributes %>>
<div<%= Orders.PromoCodeUsed.ViewAttributes %>><%= Orders.PromoCodeUsed.ListViewValue %></div>
</td>
<% End If %>
	</tr>
<%
	Orders_preview.Recordset.MoveNext
Loop
%>
	</tbody>
</table>
</div>
</td></tr></table>
<%
Orders_preview.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Close recordset and connection
Orders_preview.Recordset.Close
Set Orders_preview.Recordset = Nothing
%>
<% End If %>
<%

' Drop page object
Set Orders_preview = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrders_preview

	' Page ID
	Public Property Get PageID()
		PageID = "preview"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Orders"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Orders_preview"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Orders.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Orders.TableVar & "&" ' add page token
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
		If Orders.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Orders.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Orders.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Orders) Then Set Orders = New cOrders
		Set Table = Orders

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Customers) Then Set Customers = New cCustomers

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "preview"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Orders"

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
		If IsEmpty(Security) Then Set Security = New cAdvancedSecurity
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then
			Response.Write Language.Phrase("NoPermission")
			Set Security = Nothing
			Response.End
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
		Set Orders = Nothing
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

	Dim Recordset
	Dim TotalRecs
	Dim RecCount
	Dim RowCnt

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load filter
		Dim QS, filter
		QS = Split(Request.Querystring, "&")
		filter = GetValue(QS, "f")
		filter = TEAdecrypt(filter, EW_RANDOM_KEY)
		If filter = "" Then filter = "0=1"

		' Load recordset
		' Call Recordset Selecting event

		Call Orders.Recordset_Selecting(filter)
		Set Recordset = Orders.LoadRs(filter)
		If Not (Recordset Is Nothing) Then
			TotalRecs = Recordset.RecordCount
		Else
			TotalRecs = 0
		End If

		' Call Recordset Selected event
		Call Orders.Recordset_Selected(Recordset)
	End Sub

Function GetValue(QS, Key)
	Dim kv, i
	For i = 0 To UBound(QS)
		kv = Split(QS(i), "=")
		If (kv(0) = Key) Then
			GetValue = ew_Decode(kv(1))
			Exit Function
		End If
	Next
	GetValue = ""
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
