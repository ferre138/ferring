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
Dim OrderDetails_list
Set OrderDetails_list = New cOrderDetails_list
Set Page = OrderDetails_list

' Page init processing
Call OrderDetails_list.Page_Init()

' Page main processing
Call OrderDetails_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If OrderDetails.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var OrderDetails_list = new ew_Page("OrderDetails_list");
// page properties
OrderDetails_list.PageID = "list"; // page ID
OrderDetails_list.FormID = "fOrderDetailslist"; // form ID
var EW_PAGE_ID = OrderDetails_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
OrderDetails_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
OrderDetails_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
OrderDetails_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
OrderDetails_list.ValidateRequired = false; // no JavaScript validation
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
<% End If %>
<% If (OrderDetails.Export = "") Or (EW_EXPORT_MASTER_RECORD And OrderDetails.Export = "print") Then %>
<%
gsMasterReturnUrl = "Orderslist.asp"
If OrderDetails_list.DbMasterFilter <> "" And OrderDetails.CurrentMasterTable = "Orders" Then
	If OrderDetails_list.MasterRecordExists Then
		If OrderDetails.CurrentMasterTable = OrderDetails.TableVar Then gsMasterReturnUrl = gsMasterReturnUrl & "?" & EW_TABLE_SHOW_MASTER & "="
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("MasterRecord") %><%= Orders.TableCaption %>
&nbsp;&nbsp;<% OrderDetails_list.ExportOptions.Render "body", "" %>
</p>
<% If OrderDetails.Export = "" Then %>
<p class="aspmaker"><a href="<%= gsMasterReturnUrl %>"><%= Language.Phrase("BackToMasterPage") %></a></p>
<% End If %>
<!--#include file="Ordersmaster.asp"-->
<%
	End If
End If
%>
<% End If %>
<% OrderDetails_list.ShowPageHeader() %>
<%

' Load recordset
Set OrderDetails_list.Recordset = OrderDetails_list.LoadRecordset()
	OrderDetails_list.TotalRecs = OrderDetails_list.Recordset.RecordCount
	OrderDetails_list.StartRec = 1
	If OrderDetails_list.DisplayRecs <= 0 Then ' Display all records
		OrderDetails_list.DisplayRecs = OrderDetails_list.TotalRecs
	End If
	If Not (OrderDetails.ExportAll And OrderDetails.Export <> "") Then
		OrderDetails_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= OrderDetails.TableCaption %>
<% If OrderDetails.CurrentMasterTable = "" Then %>
&nbsp;&nbsp;<% OrderDetails_list.ExportOptions.Render "body", "" %>
<% End If %>
</p>
<% OrderDetails_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If OrderDetails.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If OrderDetails.CurrentAction <> "gridadd" And OrderDetails.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(OrderDetails_list.Pager) Then Set OrderDetails_list.Pager = ew_NewNumericPager(OrderDetails_list.StartRec, OrderDetails_list.DisplayRecs, OrderDetails_list.TotalRecs, OrderDetails_list.RecRange) %>
<% If OrderDetails_list.Pager.RecordCount > 0 Then %>
	<% If OrderDetails_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In OrderDetails_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= OrderDetails_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If OrderDetails_list.Pager.NextButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.LastButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= OrderDetails_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= OrderDetails_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= OrderDetails_list.Pager.RecordCount %>
<% Else %>
	<% If OrderDetails_list.SearchWhere = "0=101" Then %>
	<%= Language.Phrase("EnterSearchCriteria") %>
	<% Else %>
	<%= Language.Phrase("NoRecord") %>
	<% End If %>
<% End If %>
</span>
		</td>
	</tr>
</table>
</form>
<% End If %>
<span class="aspmaker">
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= OrderDetails_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If OrderDetails_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fOrderDetailslist, '<%= OrderDetails_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fOrderDetailslist" id="fOrderDetailslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="OrderDetails">
<div id="gmp_OrderDetails" class="ewGridMiddlePanel">
<% If OrderDetails_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= OrderDetails.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call OrderDetails_list.RenderListOptions()

' Render list options (header, left)
OrderDetails_list.ListOptions.Render "header", "left"
%>
<% If OrderDetails.ProductId.Visible Then ' ProductId %>
	<% If OrderDetails.SortUrl(OrderDetails.ProductId) = "" Then %>
		<td><%= OrderDetails.ProductId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= OrderDetails.SortUrl(OrderDetails.ProductId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.ProductId.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.ProductId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.ProductId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If OrderDetails.Quantity.Visible Then ' Quantity %>
	<% If OrderDetails.SortUrl(OrderDetails.Quantity) = "" Then %>
		<td><%= OrderDetails.Quantity.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= OrderDetails.SortUrl(OrderDetails.Quantity) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.Quantity.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.Quantity.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.Quantity.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If OrderDetails.Price.Visible Then ' Price %>
	<% If OrderDetails.SortUrl(OrderDetails.Price) = "" Then %>
		<td><%= OrderDetails.Price.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= OrderDetails.SortUrl(OrderDetails.Price) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.Price.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.Price.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.Price.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
OrderDetails_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (OrderDetails.ExportAll And OrderDetails.Export <> "") Then
	OrderDetails_list.StopRec = OrderDetails_list.TotalRecs
Else

	' Set the last record to display
	If OrderDetails_list.TotalRecs > OrderDetails_list.StartRec + OrderDetails_list.DisplayRecs - 1 Then
		OrderDetails_list.StopRec = OrderDetails_list.StartRec + OrderDetails_list.DisplayRecs - 1
	Else
		OrderDetails_list.StopRec = OrderDetails_list.TotalRecs
	End If
End If

' Move to first record
OrderDetails_list.RecCnt = OrderDetails_list.StartRec - 1
If Not OrderDetails_list.Recordset.Eof Then
	OrderDetails_list.Recordset.MoveFirst
	If OrderDetails_list.StartRec > 1 Then OrderDetails_list.Recordset.Move OrderDetails_list.StartRec - 1
ElseIf Not OrderDetails.AllowAddDeleteRow And OrderDetails_list.StopRec = 0 Then
	OrderDetails_list.StopRec = OrderDetails.GridAddRowCount
End If

' Initialize Aggregate
OrderDetails.RowType = EW_ROWTYPE_AGGREGATEINIT
Call OrderDetails.ResetAttrs()
Call OrderDetails_list.RenderRow()
OrderDetails_list.RowCnt = 0

' Output date rows
Do While CLng(OrderDetails_list.RecCnt) < CLng(OrderDetails_list.StopRec)
	OrderDetails_list.RecCnt = OrderDetails_list.RecCnt + 1
	If CLng(OrderDetails_list.RecCnt) >= CLng(OrderDetails_list.StartRec) Then
		OrderDetails_list.RowCnt = OrderDetails_list.RowCnt + 1

	' Set up key count
	OrderDetails_list.KeyCount = OrderDetails_list.RowIndex
	Call OrderDetails.ResetAttrs()
	OrderDetails.CssClass = ""
	If OrderDetails.CurrentAction = "gridadd" Then
	Else
		Call OrderDetails_list.LoadRowValues(OrderDetails_list.Recordset) ' Load row values
	End If
	OrderDetails.RowType = EW_ROWTYPE_VIEW ' Render view
	OrderDetails.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call OrderDetails_list.RenderRow()

	' Render list options
	Call OrderDetails_list.RenderListOptions()
%>
	<tr<%= OrderDetails.RowAttributes %>>
<%

' Render list options (body, left)
OrderDetails_list.ListOptions.Render "body", "left"
%>
	<% If OrderDetails.ProductId.Visible Then ' ProductId %>
		<td<%= OrderDetails.ProductId.CellAttributes %>>
<div<%= OrderDetails.ProductId.ViewAttributes %>><%= OrderDetails.ProductId.ListViewValue %></div>
<a name="<%= OrderDetails_list.PageObjName & "_row_" & OrderDetails_list.RowCnt %>" id="<%= OrderDetails_list.PageObjName & "_row_" & OrderDetails_list.RowCnt %>"></a></td>
	<% End If %>
	<% If OrderDetails.Quantity.Visible Then ' Quantity %>
		<td<%= OrderDetails.Quantity.CellAttributes %>>
<div<%= OrderDetails.Quantity.ViewAttributes %>><%= OrderDetails.Quantity.ListViewValue %></div>
</td>
	<% End If %>
	<% If OrderDetails.Price.Visible Then ' Price %>
		<td<%= OrderDetails.Price.CellAttributes %>>
<div<%= OrderDetails.Price.ViewAttributes %>><%= OrderDetails.Price.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
OrderDetails_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If OrderDetails.CurrentAction <> "gridadd" Then
		OrderDetails_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
</div>
</form>
<%

' Close recordset and connection
OrderDetails_list.Recordset.Close
Set OrderDetails_list.Recordset = Nothing
%>
<% If OrderDetails_list.TotalRecs > 0 Then %>
<% If OrderDetails.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If OrderDetails.CurrentAction <> "gridadd" And OrderDetails.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(OrderDetails_list.Pager) Then Set OrderDetails_list.Pager = ew_NewNumericPager(OrderDetails_list.StartRec, OrderDetails_list.DisplayRecs, OrderDetails_list.TotalRecs, OrderDetails_list.RecRange) %>
<% If OrderDetails_list.Pager.RecordCount > 0 Then %>
	<% If OrderDetails_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In OrderDetails_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= OrderDetails_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If OrderDetails_list.Pager.NextButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.LastButton.Enabled Then %>
	<a href="<%= OrderDetails_list.PageUrl %>start=<%= OrderDetails_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If OrderDetails_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= OrderDetails_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= OrderDetails_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= OrderDetails_list.Pager.RecordCount %>
<% Else %>
	<% If OrderDetails_list.SearchWhere = "0=101" Then %>
	<%= Language.Phrase("EnterSearchCriteria") %>
	<% Else %>
	<%= Language.Phrase("NoRecord") %>
	<% End If %>
<% End If %>
</span>
		</td>
	</tr>
</table>
</form>
<% End If %>
<span class="aspmaker">
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= OrderDetails_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If OrderDetails_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fOrderDetailslist, '<%= OrderDetails_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If OrderDetails.Export = "" And OrderDetails.CurrentAction = "" Then %>
<% End If %>
<%
OrderDetails_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If OrderDetails.Export = "" Then %>
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
Set OrderDetails_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrderDetails_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "OrderDetails"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "OrderDetails_list"
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
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "OrderDetailsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "OrderDetailsdelete.asp"
		MultiUpdateUrl = "OrderDetailsupdate.asp"

		' Initialize other table object
		If IsEmpty(Orders) Then Set Orders = New cOrders

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "OrderDetails"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Initialize list options
		Set ListOptions = New cListOptions

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

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				OrderDetails.GridAddRowCount = gridaddcnt
			End If
		End If

		' Set up list options
		SetupListOptions()

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
		Set ListOptions = Nothing
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
	Dim SearchWhere
	Dim RecCnt
	Dim EditRowCnt
	Dim RowCnt, RowIndex
	Dim RecPerRow, ColCnt
	Dim KeyCount
	Dim RowAction
	Dim RowOldKey ' Row old key (for copy)
	Dim DbMasterFilter, DbDetailFilter
	Dim MasterRecordExists
	Dim ListOptions
	Dim ExportOptions
	Dim MultiSelectKey
	Dim RestoreSearch
	Dim Recordset, OldRecordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		DisplayRecs = 20
		RecRange = 10
		RecCnt = 0 ' Record count
		KeyCount = 0 ' Key count

		' Search filters
		Dim sSrchAdvanced, sSrchBasic, sFilter
		sSrchAdvanced = "" ' Advanced search filter
		sSrchBasic = "" ' Basic search filter
		SearchWhere = "" ' Search where clause
		sFilter = ""

		' Master/Detail
		DbMasterFilter = "" ' Master filter
		DbDetailFilter = "" ' Detail filter
		If IsPageRequest Then ' Validate request

			' Handle reset command
			ResetCmd()

			' Set up master detail parameters
			SetUpMasterParms()

			' Hide all options
			If OrderDetails.Export <> "" Or OrderDetails.CurrentAction = "gridadd" Or OrderDetails.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Set Up Sorting Order
			SetUpSortOrder()
		End If ' End Validate Request

		' Restore display records
		If OrderDetails.RecordsPerPage <> "" Then
			DisplayRecs = OrderDetails.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()
		sFilter = ""

		' Restore master/detail filter
		DbMasterFilter = OrderDetails.MasterFilter ' Restore master filter
		DbDetailFilter = OrderDetails.DetailFilter ' Restore detail filter
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)
		Dim RsMaster

		' Load master record
		If OrderDetails.MasterFilter <> "" And OrderDetails.CurrentMasterTable = "Orders" Then
			Set RsMaster = Orders.LoadRs(DbMasterFilter)
			MasterRecordExists = Not (RsMaster Is Nothing)
			If Not MasterRecordExists Then
				FailureMessage = Language.Phrase("NoRecord") ' Set no record found
				Call Page_Terminate(OrderDetails.ReturnUrl) ' Return to caller
			Else
				Call Orders.LoadListRowValues(RsMaster)
				Orders.RowType = EW_ROWTYPE_MASTER ' Master row
				Call Orders.RenderListRow()
				RsMaster.Close
				Set RsMaster = Nothing
			End If
		End If

		' Set up filter in Session
		OrderDetails.SessionWhere = sFilter
		OrderDetails.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Set up Sort parameters based on Sort Links clicked
	'
	Sub SetUpSortOrder()
		Dim sOrderBy
		Dim sSortField, sLastSort, sThisSort
		Dim bCtrl

		' Check for an Order parameter
		If Request.QueryString("order").Count > 0 Then
			OrderDetails.CurrentOrder = Request.QueryString("order")
			OrderDetails.CurrentOrderType = Request.QueryString("ordertype")

			' Field ProductId
			Call OrderDetails.UpdateSort(OrderDetails.ProductId)

			' Field Quantity
			Call OrderDetails.UpdateSort(OrderDetails.Quantity)

			' Field Price
			Call OrderDetails.UpdateSort(OrderDetails.Price)
			OrderDetails.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = OrderDetails.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If OrderDetails.SqlOrderBy <> "" Then
				sOrderBy = OrderDetails.SqlOrderBy
				OrderDetails.SessionOrderBy = sOrderBy
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Reset command based on querystring parameter cmd=
	' - RESET: reset search parameters
	' - RESETALL: reset search & master/detail parameters
	' - RESETSORT: reset sort parameters
	'
	Sub ResetCmd()
		Dim sCmd

		' Get reset cmd
		If Request.QueryString("cmd").Count > 0 Then
			sCmd = Request.QueryString("cmd")

			' Reset master/detail keys
			If LCase(sCmd) = "resetall" Then
				OrderDetails.CurrentMasterTable = "" ' Clear master table
				DbMasterFilter = ""
				DbDetailFilter = ""
				OrderDetails.OrderId.SessionValue = ""
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				OrderDetails.SessionOrderBy = sOrderBy
				OrderDetails.ProductId.Sort = ""
				OrderDetails.Quantity.Sort = ""
				OrderDetails.Price.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			OrderDetails.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		ListOptions.Add("view")
		ListOptions.GetItem("view").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("view").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("view").OnLeft = True
		ListOptions.Add("edit")
		ListOptions.GetItem("edit").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("edit").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("edit").OnLeft = True
		ListOptions.Add("copy")
		ListOptions.GetItem("copy").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("copy").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("copy").OnLeft = True
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("checkbox").OnLeft = True
		ListOptions.MoveItem "checkbox", 0 ' Move to first column
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""OrderDetails_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.IsLoggedIn() And ListOptions.GetItem("view").Visible Then
			ListOptions.GetItem("view").Body = "<a class=""ewRowLink"" href=""" & ViewUrl & """>" & "<img src=""images/view.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("ViewLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("ViewLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
		End If
		If Security.IsLoggedIn() And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
		End If
		If Security.IsLoggedIn() And ListOptions.GetItem("copy").Visible Then
			Set item = ListOptions.GetItem("copy")
			item.Body = "<a class=""ewRowLink"" href=""" & CopyUrl & """>" & "<img src=""images/copy.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("CopyLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("CopyLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
		End If
		If Security.IsLoggedIn() And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(OrderDetails.OrderDetailsId.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
		End If
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	Function RenderListOptionsExt()
	End Function
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
				OrderDetails.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					OrderDetails.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = OrderDetails.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			OrderDetails.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			OrderDetails.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			OrderDetails.StartRecordNumber = StartRec
		End If
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If OrderDetails.GetKey("OrderDetailsId")&"" <> "" Then
			OrderDetails.OrderDetailsId.CurrentValue = OrderDetails.GetKey("OrderDetailsId") ' OrderDetailsId
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			OrderDetails.CurrentFilter = OrderDetails.KeyFilter
			Dim sSql
			sSql = OrderDetails.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		ViewUrl = OrderDetails.ViewUrl
		EditUrl = OrderDetails.EditUrl("")
		InlineEditUrl = OrderDetails.InlineEditUrl
		CopyUrl = OrderDetails.CopyUrl("")
		InlineCopyUrl = OrderDetails.InlineCopyUrl
		DeleteUrl = OrderDetails.DeleteUrl

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

	' ListOptions Load event
	Sub ListOptions_Load()

		'Example: 
		' Dim opt
		' Set opt = ListOptions.Add("new")
		' opt.OnLeft = True ' Link on left
		' opt.MoveTo 0 ' Move to first column

	End Sub

	' ListOptions Rendered event
	Sub ListOptions_Rendered()

		'Example: 
		'ListOptions.GetItem("new").Body = "xxx"

	End Sub
End Class
%>
