<%@ CodePage="65001" %>
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
<% Call ew_Header(False, "utf-8") %>
<%

' Define page object
Dim OrderDetails_preview
Set OrderDetails_preview = New cOrderDetails_preview
Set Page = OrderDetails_preview

' Page init processing
Call OrderDetails_preview.Page_Init()

' Page main processing
Call OrderDetails_preview.Page_Main()
%>
<link href="css/vslpaypal.css" rel="stylesheet" type="text/css">
<% OrderDetails_preview.ShowPageHeader() %>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= OrderDetails.TableCaption %>
<% If OrderDetails_preview.TotalRecs > 0 Then %>
(<%= OrderDetails_preview.TotalRecs %>&nbsp;<%= Language.Phrase("Record") %>)
<% Else %>
(<%= Language.Phrase("NoRecord") %>)
<% End If %>
</p>
<% If OrderDetails_preview.TotalRecs > 0 Then %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="ewDetailsPreviewTable" name="ewDetailsPreviewTable" cellspacing="0" class="ewTable ewTableSeparate">
	<thead><!-- Table header -->
		<tr class="ewTableHeader">
<% If OrderDetails.Price.Visible Then ' ProductId %>
			<td valign="top"><%= OrderDetails.ProductId.FldCaption %></td>
<% End If %>
<% If OrderDetails.Price.Visible Then ' Quantity %>
			<td valign="top"><%= OrderDetails.Quantity.FldCaption %></td>
<% End If %>
<% If OrderDetails.Price.Visible Then ' Price %>
			<td valign="top"><%= OrderDetails.Price.FldCaption %></td>
<% End If %>
		</tr>
	</thead>
	<tbody><!-- Table body -->
<%
OrderDetails_preview.RecCount = 0
OrderDetails_preview.RowCnt = 0

'OrderDetails_preview.Recordset.MoveFirst
Do While (Not OrderDetails_preview.Recordset.Eof)

	' Init row class and style
	OrderDetails_preview.RecCount = OrderDetails_preview.RecCount + 1
	OrderDetails_preview.RowCnt = OrderDetails_preview.RowCnt + 1
	OrderDetails.CssClass = ""
	OrderDetails.CssStyle = ""
	Call OrderDetails.LoadListRowValues(OrderDetails_preview.Recordset)

	' Render row
	OrderDetails.RowType = EW_ROWTYPE_PREVIEW ' Preview record
	Call OrderDetails.RenderListRow()
%>
	<tr<%= OrderDetails.RowAttributes %>>
<% If OrderDetails.ProductId.Visible Then ' ProductId %>
		<!-- ProductId -->
		<td<%= OrderDetails.ProductId.CellAttributes %>>
<div<%= OrderDetails.ProductId.ViewAttributes %>><%= OrderDetails.ProductId.ListViewValue %></div>
</td>
<% End If %>
<% If OrderDetails.Quantity.Visible Then ' Quantity %>
		<!-- Quantity -->
		<td<%= OrderDetails.Quantity.CellAttributes %>>
<div<%= OrderDetails.Quantity.ViewAttributes %>><%= OrderDetails.Quantity.ListViewValue %></div>
</td>
<% End If %>
<% If OrderDetails.Price.Visible Then ' Price %>
		<!-- Price -->
		<td<%= OrderDetails.Price.CellAttributes %>>
<div<%= OrderDetails.Price.ViewAttributes %>><%= OrderDetails.Price.ListViewValue %></div>
</td>
<% End If %>
	</tr>
<%
	OrderDetails_preview.Recordset.MoveNext
Loop
%>
	</tbody>
</table>
</div>
</td></tr></table>
<%
OrderDetails_preview.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Close recordset and connection
OrderDetails_preview.Recordset.Close
Set OrderDetails_preview.Recordset = Nothing
%>
<% End If %>
<%

' Drop page object
Set OrderDetails_preview = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrderDetails_preview

	' Page ID
	Public Property Get PageID()
		PageID = "preview"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "OrderDetails"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "OrderDetails_preview"
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
		EW_PAGE_ID = "preview"

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

		Call OrderDetails.Recordset_Selecting(filter)
		Set Recordset = OrderDetails.LoadRs(filter)
		If Not (Recordset Is Nothing) Then
			TotalRecs = Recordset.RecordCount
		Else
			TotalRecs = 0
		End If

		' Call Recordset Selected event
		Call OrderDetails.Recordset_Selected(Recordset)
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
