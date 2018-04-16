<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Logins_list
Set Logins_list = New cLogins_list
Set Page = Logins_list

' Page init processing
Call Logins_list.Page_Init()

' Page main processing
Call Logins_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Logins.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Logins_list = new ew_Page("Logins_list");
// page properties
Logins_list.PageID = "list"; // page ID
Logins_list.FormID = "fLoginslist"; // form ID
var EW_PAGE_ID = Logins_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Logins_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Logins_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Logins_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Logins_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Logins.Export = "") Or (EW_EXPORT_MASTER_RECORD And Logins.Export = "print") Then %>
<% End If %>
<% Logins_list.ShowPageHeader() %>
<%

' Load recordset
Set Logins_list.Recordset = Logins_list.LoadRecordset()
	Logins_list.TotalRecs = Logins_list.Recordset.RecordCount
	Logins_list.StartRec = 1
	If Logins_list.DisplayRecs <= 0 Then ' Display all records
		Logins_list.DisplayRecs = Logins_list.TotalRecs
	End If
	If Not (Logins.ExportAll And Logins.Export <> "") Then
		Logins_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Logins.TableCaption %>
&nbsp;&nbsp;<% Logins_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Logins.Export = "" And Logins.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Logins_list);" style="text-decoration: none;"><img id="Logins_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Logins_list_SearchPanel">
<form name="fLoginslistsrch" id="fLoginslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Logins">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Logins.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Logins_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Logins.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Logins.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Logins.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Logins_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Logins.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Logins.CurrentAction <> "gridadd" And Logins.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Logins_list.Pager) Then Set Logins_list.Pager = ew_NewNumericPager(Logins_list.StartRec, Logins_list.DisplayRecs, Logins_list.TotalRecs, Logins_list.RecRange) %>
<% If Logins_list.Pager.RecordCount > 0 Then %>
	<% If Logins_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Logins_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Logins_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Logins_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Logins_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Logins_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Logins_list.Pager.RecordCount %>
<% Else %>
	<% If Logins_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Logins_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Logins_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fLoginslist, '<%= Logins_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fLoginslist" id="fLoginslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Logins">
<div id="gmp_Logins" class="ewGridMiddlePanel">
<% If Logins_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Logins.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Logins_list.RenderListOptions()

' Render list options (header, left)
Logins_list.ListOptions.Render "header", "left"
%>
<% If Logins.Loginname.Visible Then ' Loginname %>
	<% If Logins.SortUrl(Logins.Loginname) = "" Then %>
		<td><%= Logins.Loginname.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Logins.SortUrl(Logins.Loginname) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Logins.Loginname.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Logins.Loginname.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Logins.Loginname.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Logins.Loginpass.Visible Then ' Loginpass %>
	<% If Logins.SortUrl(Logins.Loginpass) = "" Then %>
		<td><%= Logins.Loginpass.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Logins.SortUrl(Logins.Loginpass) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Logins.Loginpass.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Logins.Loginpass.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Logins.Loginpass.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Logins_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Logins.ExportAll And Logins.Export <> "") Then
	Logins_list.StopRec = Logins_list.TotalRecs
Else

	' Set the last record to display
	If Logins_list.TotalRecs > Logins_list.StartRec + Logins_list.DisplayRecs - 1 Then
		Logins_list.StopRec = Logins_list.StartRec + Logins_list.DisplayRecs - 1
	Else
		Logins_list.StopRec = Logins_list.TotalRecs
	End If
End If

' Move to first record
Logins_list.RecCnt = Logins_list.StartRec - 1
If Not Logins_list.Recordset.Eof Then
	Logins_list.Recordset.MoveFirst
	If Logins_list.StartRec > 1 Then Logins_list.Recordset.Move Logins_list.StartRec - 1
ElseIf Not Logins.AllowAddDeleteRow And Logins_list.StopRec = 0 Then
	Logins_list.StopRec = Logins.GridAddRowCount
End If

' Initialize Aggregate
Logins.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Logins.ResetAttrs()
Call Logins_list.RenderRow()
Logins_list.RowCnt = 0

' Output date rows
Do While CLng(Logins_list.RecCnt) < CLng(Logins_list.StopRec)
	Logins_list.RecCnt = Logins_list.RecCnt + 1
	If CLng(Logins_list.RecCnt) >= CLng(Logins_list.StartRec) Then
		Logins_list.RowCnt = Logins_list.RowCnt + 1

	' Set up key count
	Logins_list.KeyCount = Logins_list.RowIndex
	Call Logins.ResetAttrs()
	Logins.CssClass = ""
	If Logins.CurrentAction = "gridadd" Then
	Else
		Call Logins_list.LoadRowValues(Logins_list.Recordset) ' Load row values
	End If
	Logins.RowType = EW_ROWTYPE_VIEW ' Render view
	Logins.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Logins_list.RenderRow()

	' Render list options
	Call Logins_list.RenderListOptions()
%>
	<tr<%= Logins.RowAttributes %>>
<%

' Render list options (body, left)
Logins_list.ListOptions.Render "body", "left"
%>
	<% If Logins.Loginname.Visible Then ' Loginname %>
		<td<%= Logins.Loginname.CellAttributes %>>
<div<%= Logins.Loginname.ViewAttributes %>><%= Logins.Loginname.ListViewValue %></div>
<a name="<%= Logins_list.PageObjName & "_row_" & Logins_list.RowCnt %>" id="<%= Logins_list.PageObjName & "_row_" & Logins_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Logins.Loginpass.Visible Then ' Loginpass %>
		<td<%= Logins.Loginpass.CellAttributes %>>
<div<%= Logins.Loginpass.ViewAttributes %>><%= Logins.Loginpass.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Logins_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Logins.CurrentAction <> "gridadd" Then
		Logins_list.Recordset.MoveNext()
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
Logins_list.Recordset.Close
Set Logins_list.Recordset = Nothing
%>
<% If Logins_list.TotalRecs > 0 Then %>
<% If Logins.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Logins.CurrentAction <> "gridadd" And Logins.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Logins_list.Pager) Then Set Logins_list.Pager = ew_NewNumericPager(Logins_list.StartRec, Logins_list.DisplayRecs, Logins_list.TotalRecs, Logins_list.RecRange) %>
<% If Logins_list.Pager.RecordCount > 0 Then %>
	<% If Logins_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Logins_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Logins_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Logins_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Logins_list.PageUrl %>start=<%= Logins_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Logins_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Logins_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Logins_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Logins_list.Pager.RecordCount %>
<% Else %>
	<% If Logins_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Logins_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Logins_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fLoginslist, '<%= Logins_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Logins.Export = "" And Logins.CurrentAction = "" Then %>
<% End If %>
<%
Logins_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Logins.Export = "" Then %>
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
Set Logins_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cLogins_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Logins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Logins_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Logins.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Logins.TableVar & "&" ' add page token
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
		If Logins.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Logins.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Logins.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Logins) Then Set Logins = New cLogins
		Set Table = Logins

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Loginsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Loginsdelete.asp"
		MultiUpdateUrl = "Loginsupdate.asp"

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Logins"

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
				Logins.GridAddRowCount = gridaddcnt
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
		Set Logins = Nothing
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

			' Hide all options
			If Logins.Export <> "" Or Logins.CurrentAction = "gridadd" Or Logins.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Logins.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Logins.RecordsPerPage <> "" Then
			DisplayRecs = Logins.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Logins.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Logins.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Logins.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Logins.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Logins.SessionWhere = sFilter
		Logins.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Logins.Loginname, Keyword)
			Call BuildBasicSearchSQL(sWhere, Logins.Loginpass, Keyword)
		BasicSearchSQL = sWhere
	End Function

	' -----------------------------------------------------------------
	' Build basic search sql
	'
	Sub BuildBasicSearchSql(Where, Fld, Keyword)
		Dim sFldExpression, lFldDataType
		Dim sWrk
		If Fld.FldVirtualExpression <> "" Then
			sFldExpression = Fld.FldVirtualExpression
		Else
			sFldExpression = Fld.FldExpression
		End If
		lFldDataType = Fld.FldDataType
		If Fld.FldIsVirtual Then lFldDataType = EW_DATATYPE_STRING
		If lFldDataType = EW_DATATYPE_NUMBER Then
			sWrk = sFldExpression & " = " & ew_QuotedValue(Keyword, lFldDataType)
		Else
			sWrk = sFldExpression & ew_Like(ew_QuotedValue("%" & Keyword & "%", lFldDataType))
		End If
		If Where <> "" Then Where = Where & " OR "
		Where = Where & sWrk
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search Where based on search keyword and type
	'
	Function BasicSearchWhere()
		Dim sSearchStr, sSearchKeyword, sSearchType
		Dim sSearch, arKeyword, sKeyword
		sSearchStr = ""
		sSearchKeyword = Logins.BasicSearchKeyword
		sSearchType = Logins.BasicSearchType
		If sSearchKeyword <> "" Then
			sSearch = Trim(sSearchKeyword)
			If sSearchType <> "" Then
				While InStr(sSearch, "  ") > 0
					sSearch = Replace(sSearch, "  ", " ")
				Wend
				arKeyword = Split(Trim(sSearch), " ")
				For Each sKeyword In arKeyword
					If sSearchStr <> "" Then sSearchStr = sSearchStr & " " & sSearchType & " "
					sSearchStr = sSearchStr & "(" & BasicSearchSQL(sKeyword) & ")"
				Next
			Else
				sSearchStr = BasicSearchSQL(sSearch)
			End If
		End If
		If sSearchKeyword <> "" then
			Logins.SessionBasicSearchKeyword = sSearchKeyword
			Logins.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Logins.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Logins.SessionBasicSearchKeyword = ""
		Logins.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Logins.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Logins.BasicSearchKeyword = Logins.SessionBasicSearchKeyword
			Logins.BasicSearchType = Logins.SessionBasicSearchType
		End If
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
			Logins.CurrentOrder = Request.QueryString("order")
			Logins.CurrentOrderType = Request.QueryString("ordertype")

			' Field Loginname
			Call Logins.UpdateSort(Logins.Loginname)

			' Field Loginpass
			Call Logins.UpdateSort(Logins.Loginpass)
			Logins.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Logins.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Logins.SqlOrderBy <> "" Then
				sOrderBy = Logins.SqlOrderBy
				Logins.SessionOrderBy = sOrderBy
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

			' Reset search criteria
			If LCase(sCmd) = "reset" Or LCase(sCmd) = "resetall" Then
				Call ResetSearchParms()
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				Logins.SessionOrderBy = sOrderBy
				Logins.Loginname.Sort = ""
				Logins.Loginpass.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Logins.StartRecordNumber = StartRec
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
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("checkbox").OnLeft = True
		ListOptions.MoveItem "checkbox", 0 ' Move to first column
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Logins_list.SelectAllKey(this);"">"
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
		If Security.IsLoggedIn() And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Logins.zUserId.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Logins.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Logins.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Logins.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Logins.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Logins.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Logins.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Logins.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Logins.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Logins.CurrentFilter
		Call Logins.Recordset_Selecting(sFilter)
		Logins.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Logins.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Logins.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Logins.KeyFilter

		' Call Row Selecting event
		Call Logins.Row_Selecting(sFilter)

		' Load sql based on filter
		Logins.CurrentFilter = sFilter
		sSql = Logins.SQL
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
		Call Logins.Row_Selected(RsRow)
		Logins.zUserId.DbValue = RsRow("UserId")
		Logins.Loginname.DbValue = RsRow("Loginname")
		Logins.Loginpass.DbValue = RsRow("Loginpass")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Logins.GetKey("zUserId")&"" <> "" Then
			Logins.zUserId.CurrentValue = Logins.GetKey("zUserId") ' UserId
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Logins.CurrentFilter = Logins.KeyFilter
			Dim sSql
			sSql = Logins.SQL
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
		ViewUrl = Logins.ViewUrl
		EditUrl = Logins.EditUrl("")
		InlineEditUrl = Logins.InlineEditUrl
		CopyUrl = Logins.CopyUrl("")
		InlineCopyUrl = Logins.InlineCopyUrl
		DeleteUrl = Logins.DeleteUrl

		' Call Row Rendering event
		Call Logins.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' UserId
		' Loginname
		' Loginpass
		' -----------
		'  View  Row
		' -----------

		If Logins.RowType = EW_ROWTYPE_VIEW Then ' View row

			' UserId
			Logins.zUserId.ViewValue = Logins.zUserId.CurrentValue
			Logins.zUserId.ViewCustomAttributes = ""

			' Loginname
			Logins.Loginname.ViewValue = Logins.Loginname.CurrentValue
			Logins.Loginname.ViewCustomAttributes = ""

			' Loginpass
			Logins.Loginpass.ViewValue = Logins.Loginpass.CurrentValue
			Logins.Loginpass.ViewCustomAttributes = ""

			' View refer script
			' Loginname

			Logins.Loginname.LinkCustomAttributes = ""
			Logins.Loginname.HrefValue = ""
			Logins.Loginname.TooltipValue = ""

			' Loginpass
			Logins.Loginpass.LinkCustomAttributes = ""
			Logins.Loginpass.HrefValue = ""
			Logins.Loginpass.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Logins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Logins.Row_Rendered()
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
