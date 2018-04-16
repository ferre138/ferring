<%@ CodePage="65001" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="DiscountTypesinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="Discountcodesinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, "utf-8") %>
<%

' Define page object
Dim DiscountTypes_preview
Set DiscountTypes_preview = New cDiscountTypes_preview
Set Page = DiscountTypes_preview

' Page init processing
Call DiscountTypes_preview.Page_Init()

' Page main processing
Call DiscountTypes_preview.Page_Main()
%>
<link href="css/vslpaypal.css" rel="stylesheet" type="text/css">
<% DiscountTypes_preview.ShowPageHeader() %>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= DiscountTypes.TableCaption %>
<% If DiscountTypes_preview.TotalRecs > 0 Then %>
(<%= DiscountTypes_preview.TotalRecs %>&nbsp;<%= Language.Phrase("Record") %>)
<% Else %>
(<%= Language.Phrase("NoRecord") %>)
<% End If %>
</p>
<% If DiscountTypes_preview.TotalRecs > 0 Then %>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="ewDetailsPreviewTable" name="ewDetailsPreviewTable" cellspacing="0" class="ewTable ewTableSeparate">
	<thead><!-- Table header -->
		<tr class="ewTableHeader">
<% If DiscountTypes.SpecialPrice.Visible Then ' DiscountType %>
			<td valign="top"><%= DiscountTypes.DiscountType.FldCaption %></td>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' DiscountTitle %>
			<td valign="top"><%= DiscountTypes.DiscountTitle.FldCaption %></td>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' freeShipping %>
			<td valign="top"><%= DiscountTypes.freeShipping.FldCaption %></td>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' FreePerQty %>
			<td valign="top"><%= DiscountTypes.FreePerQty.FldCaption %></td>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
			<td valign="top"><%= DiscountTypes.SpecialPrice.FldCaption %></td>
<% End If %>
		</tr>
	</thead>
	<tbody><!-- Table body -->
<%
DiscountTypes_preview.RecCount = 0
DiscountTypes_preview.RowCnt = 0

'DiscountTypes_preview.Recordset.MoveFirst
Do While (Not DiscountTypes_preview.Recordset.Eof)

	' Init row class and style
	DiscountTypes_preview.RecCount = DiscountTypes_preview.RecCount + 1
	DiscountTypes_preview.RowCnt = DiscountTypes_preview.RowCnt + 1
	DiscountTypes.CssClass = ""
	DiscountTypes.CssStyle = ""
	Call DiscountTypes.LoadListRowValues(DiscountTypes_preview.Recordset)

	' Render row
	DiscountTypes.RowType = EW_ROWTYPE_PREVIEW ' Preview record
	Call DiscountTypes.RenderListRow()
%>
	<tr<%= DiscountTypes.RowAttributes %>>
<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
		<!-- DiscountType -->
		<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ListViewValue %></div>
</td>
<% End If %>
<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
		<!-- DiscountTitle -->
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ListViewValue %></div>
</td>
<% End If %>
<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
		<!-- freeShipping -->
		<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
<% End If %>
<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
		<!-- FreePerQty -->
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ListViewValue %></div>
</td>
<% End If %>
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
		<!-- SpecialPrice -->
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ListViewValue %></div>
</td>
<% End If %>
	</tr>
<%
	DiscountTypes_preview.Recordset.MoveNext
Loop
%>
	</tbody>
</table>
</div>
</td></tr></table>
<%
DiscountTypes_preview.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Close recordset and connection
DiscountTypes_preview.Recordset.Close
Set DiscountTypes_preview.Recordset = Nothing
%>
<% End If %>
<%

' Drop page object
Set DiscountTypes_preview = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountTypes_preview

	' Page ID
	Public Property Get PageID()
		PageID = "preview"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "DiscountTypes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "DiscountTypes_preview"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If DiscountTypes.UseTokenInUrl Then PageUrl = PageUrl & "t=" & DiscountTypes.TableVar & "&" ' add page token
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
		If DiscountTypes.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (DiscountTypes.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (DiscountTypes.TableVar = Request.QueryString("t"))
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
		If IsEmpty(DiscountTypes) Then Set DiscountTypes = New cDiscountTypes
		Set Table = DiscountTypes

		' Initialize urls
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "preview"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "DiscountTypes"

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
		Set DiscountTypes = Nothing
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

		Call DiscountTypes.Recordset_Selecting(filter)
		Set Recordset = DiscountTypes.LoadRs(filter)
		If Not (Recordset Is Nothing) Then
			TotalRecs = Recordset.RecordCount
		Else
			TotalRecs = 0
		End If

		' Call Recordset Selected event
		Call DiscountTypes.Recordset_Selected(Recordset)
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
