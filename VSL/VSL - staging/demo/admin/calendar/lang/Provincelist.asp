<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Provinceinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Province_list
Set Province_list = New cProvince_list
Set Page = Province_list

' Page init processing
Call Province_list.Page_Init()

' Page main processing
Call Province_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Province.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Province_list = new ew_Page("Province_list");
// page properties
Province_list.PageID = "list"; // page ID
Province_list.FormID = "fProvincelist"; // form ID
var EW_PAGE_ID = Province_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Province_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Province_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Province_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Province_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Province.Export = "") Or (EW_EXPORT_MASTER_RECORD And Province.Export = "print") Then %>
<% End If %>
<% Province_list.ShowPageHeader() %>
<%

' Load recordset
Set Province_list.Recordset = Province_list.LoadRecordset()
	Province_list.TotalRecs = Province_list.Recordset.RecordCount
	Province_list.StartRec = 1
	If Province_list.DisplayRecs <= 0 Then ' Display all records
		Province_list.DisplayRecs = Province_list.TotalRecs
	End If
	If Not (Province.ExportAll And Province.Export <> "") Then
		Province_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Province.TableCaption %>
&nbsp;&nbsp;<% Province_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Province.Export = "" And Province.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Province_list);" style="text-decoration: none;"><img id="Province_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Province_list_SearchPanel">
<form name="fProvincelistsrch" id="fProvincelistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Province">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Province.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Province_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Province.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Province.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Province.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Province_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Province.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Province.CurrentAction <> "gridadd" And Province.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Province_list.Pager) Then Set Province_list.Pager = ew_NewNumericPager(Province_list.StartRec, Province_list.DisplayRecs, Province_list.TotalRecs, Province_list.RecRange) %>
<% If Province_list.Pager.RecordCount > 0 Then %>
	<% If Province_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Province_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Province_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Province_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Province_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Province_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Province_list.Pager.RecordCount %>
<% Else %>
	<% If Province_list.SearchWhere = "0=101" Then %>
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
</span>
</div>
<% End If %>
<form name="fProvincelist" id="fProvincelist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Province">
<div id="gmp_Province" class="ewGridMiddlePanel">
<% If Province_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Province.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Province_list.RenderListOptions()

' Render list options (header, left)
Province_list.ListOptions.Render "header", "left"
%>
<% If Province.Prov.Visible Then ' Prov %>
	<% If Province.SortUrl(Province.Prov) = "" Then %>
		<td><%= Province.Prov.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.Prov) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.Prov.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Province.Prov.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.Prov.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Province.Province_1.Visible Then ' Province %>
	<% If Province.SortUrl(Province.Province_1) = "" Then %>
		<td><%= Province.Province_1.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.Province_1) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.Province_1.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Province.Province_1.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.Province_1.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Province.fProvince.Visible Then ' fProvince %>
	<% If Province.SortUrl(Province.fProvince) = "" Then %>
		<td><%= Province.fProvince.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.fProvince) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.fProvince.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Province.fProvince.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.fProvince.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Province.TaxRate.Visible Then ' TaxRate %>
	<% If Province.SortUrl(Province.TaxRate) = "" Then %>
		<td><%= Province.TaxRate.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.TaxRate) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.TaxRate.FldCaption %></td><td style="width: 10px;"><% If Province.TaxRate.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.TaxRate.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Province.ShipRate_first.Visible Then ' ShipRate_first %>
	<% If Province.SortUrl(Province.ShipRate_first) = "" Then %>
		<td><%= Province.ShipRate_first.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.ShipRate_first) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.ShipRate_first.FldCaption %></td><td style="width: 10px;"><% If Province.ShipRate_first.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.ShipRate_first.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Province.ShipRate_Rest.Visible Then ' ShipRate_Rest %>
	<% If Province.SortUrl(Province.ShipRate_Rest) = "" Then %>
		<td><%= Province.ShipRate_Rest.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Province.SortUrl(Province.ShipRate_Rest) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Province.ShipRate_Rest.FldCaption %></td><td style="width: 10px;"><% If Province.ShipRate_Rest.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Province.ShipRate_Rest.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Province_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Province.ExportAll And Province.Export <> "") Then
	Province_list.StopRec = Province_list.TotalRecs
Else

	' Set the last record to display
	If Province_list.TotalRecs > Province_list.StartRec + Province_list.DisplayRecs - 1 Then
		Province_list.StopRec = Province_list.StartRec + Province_list.DisplayRecs - 1
	Else
		Province_list.StopRec = Province_list.TotalRecs
	End If
End If

' Move to first record
Province_list.RecCnt = Province_list.StartRec - 1
If Not Province_list.Recordset.Eof Then
	Province_list.Recordset.MoveFirst
	If Province_list.StartRec > 1 Then Province_list.Recordset.Move Province_list.StartRec - 1
ElseIf Not Province.AllowAddDeleteRow And Province_list.StopRec = 0 Then
	Province_list.StopRec = Province.GridAddRowCount
End If

' Initialize Aggregate
Province.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Province.ResetAttrs()
Call Province_list.RenderRow()
Province_list.RowCnt = 0

' Output date rows
Do While CLng(Province_list.RecCnt) < CLng(Province_list.StopRec)
	Province_list.RecCnt = Province_list.RecCnt + 1
	If CLng(Province_list.RecCnt) >= CLng(Province_list.StartRec) Then
		Province_list.RowCnt = Province_list.RowCnt + 1

	' Set up key count
	Province_list.KeyCount = Province_list.RowIndex
	Call Province.ResetAttrs()
	Province.CssClass = ""
	If Province.CurrentAction = "gridadd" Then
	Else
		Call Province_list.LoadRowValues(Province_list.Recordset) ' Load row values
	End If
	Province.RowType = EW_ROWTYPE_VIEW ' Render view
	Province.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Province_list.RenderRow()

	' Render list options
	Call Province_list.RenderListOptions()
%>
	<tr<%= Province.RowAttributes %>>
<%

' Render list options (body, left)
Province_list.ListOptions.Render "body", "left"
%>
	<% If Province.Prov.Visible Then ' Prov %>
		<td<%= Province.Prov.CellAttributes %>>
<div<%= Province.Prov.ViewAttributes %>><%= Province.Prov.ListViewValue %></div>
<a name="<%= Province_list.PageObjName & "_row_" & Province_list.RowCnt %>" id="<%= Province_list.PageObjName & "_row_" & Province_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Province.Province_1.Visible Then ' Province %>
		<td<%= Province.Province_1.CellAttributes %>>
<div<%= Province.Province_1.ViewAttributes %>><%= Province.Province_1.ListViewValue %></div>
</td>
	<% End If %>
	<% If Province.fProvince.Visible Then ' fProvince %>
		<td<%= Province.fProvince.CellAttributes %>>
<div<%= Province.fProvince.ViewAttributes %>><%= Province.fProvince.ListViewValue %></div>
</td>
	<% End If %>
	<% If Province.TaxRate.Visible Then ' TaxRate %>
		<td<%= Province.TaxRate.CellAttributes %>>
<div<%= Province.TaxRate.ViewAttributes %>><%= Province.TaxRate.ListViewValue %></div>
</td>
	<% End If %>
	<% If Province.ShipRate_first.Visible Then ' ShipRate_first %>
		<td<%= Province.ShipRate_first.CellAttributes %>>
<div<%= Province.ShipRate_first.ViewAttributes %>><%= Province.ShipRate_first.ListViewValue %></div>
</td>
	<% End If %>
	<% If Province.ShipRate_Rest.Visible Then ' ShipRate_Rest %>
		<td<%= Province.ShipRate_Rest.CellAttributes %>>
<div<%= Province.ShipRate_Rest.ViewAttributes %>><%= Province.ShipRate_Rest.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Province_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Province.CurrentAction <> "gridadd" Then
		Province_list.Recordset.MoveNext()
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
Province_list.Recordset.Close
Set Province_list.Recordset = Nothing
%>
<% If Province_list.TotalRecs > 0 Then %>
<% If Province.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Province.CurrentAction <> "gridadd" And Province.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Province_list.Pager) Then Set Province_list.Pager = ew_NewNumericPager(Province_list.StartRec, Province_list.DisplayRecs, Province_list.TotalRecs, Province_list.RecRange) %>
<% If Province_list.Pager.RecordCount > 0 Then %>
	<% If Province_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Province_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Province_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Province_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Province_list.PageUrl %>start=<%= Province_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Province_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Province_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Province_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Province_list.Pager.RecordCount %>
<% Else %>
	<% If Province_list.SearchWhere = "0=101" Then %>
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
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Province.Export = "" And Province.CurrentAction = "" Then %>
<% End If %>
<%
Province_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Province.Export = "" Then %>
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
Set Province_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProvince_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Province"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Province_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Province.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Province.TableVar & "&" ' add page token
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
		If Province.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Province.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Province.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Province) Then Set Province = New cProvince
		Set Table = Province

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Provinceadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Provincedelete.asp"
		MultiUpdateUrl = "Provinceupdate.asp"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Province"

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
				Province.GridAddRowCount = gridaddcnt
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
		Set Province = Nothing
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
			If Province.Export <> "" Or Province.CurrentAction = "gridadd" Or Province.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Province.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Province.RecordsPerPage <> "" Then
			DisplayRecs = Province.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Province.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Province.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Province.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Province.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Province.SessionWhere = sFilter
		Province.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Province.Prov, Keyword)
			Call BuildBasicSearchSQL(sWhere, Province.Province_1, Keyword)
			Call BuildBasicSearchSQL(sWhere, Province.fProvince, Keyword)
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
		sSearchKeyword = Province.BasicSearchKeyword
		sSearchType = Province.BasicSearchType
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
			Province.SessionBasicSearchKeyword = sSearchKeyword
			Province.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Province.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Province.SessionBasicSearchKeyword = ""
		Province.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Province.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Province.BasicSearchKeyword = Province.SessionBasicSearchKeyword
			Province.BasicSearchType = Province.SessionBasicSearchType
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
			Province.CurrentOrder = Request.QueryString("order")
			Province.CurrentOrderType = Request.QueryString("ordertype")

			' Field Prov
			Call Province.UpdateSort(Province.Prov)

			' Field Province
			Call Province.UpdateSort(Province.Province_1)

			' Field fProvince
			Call Province.UpdateSort(Province.fProvince)

			' Field TaxRate
			Call Province.UpdateSort(Province.TaxRate)

			' Field ShipRate_first
			Call Province.UpdateSort(Province.ShipRate_first)

			' Field ShipRate_Rest
			Call Province.UpdateSort(Province.ShipRate_Rest)
			Province.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Province.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Province.SqlOrderBy <> "" Then
				sOrderBy = Province.SqlOrderBy
				Province.SessionOrderBy = sOrderBy
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
				Province.SessionOrderBy = sOrderBy
				Province.Prov.Sort = ""
				Province.Province_1.Sort = ""
				Province.fProvince.Sort = ""
				Province.TaxRate.Sort = ""
				Province.ShipRate_first.Sort = ""
				Province.ShipRate_Rest.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Province.StartRecordNumber = StartRec
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
				Province.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Province.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Province.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Province.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Province.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Province.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Province.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Province.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Province.CurrentFilter
		Call Province.Recordset_Selecting(sFilter)
		Province.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Province.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Province.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Province.KeyFilter

		' Call Row Selecting event
		Call Province.Row_Selecting(sFilter)

		' Load sql based on filter
		Province.CurrentFilter = sFilter
		sSql = Province.SQL
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
		Call Province.Row_Selected(RsRow)
		Province.Prov.DbValue = RsRow("Prov")
		Province.Province_1.DbValue = RsRow("Province")
		Province.fProvince.DbValue = RsRow("fProvince")
		Province.TaxRate.DbValue = RsRow("TaxRate")
		Province.ShipRate_first.DbValue = RsRow("ShipRate_first")
		Province.ShipRate_Rest.DbValue = RsRow("ShipRate_Rest")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Province.GetKey("Prov")&"" <> "" Then
			Province.Prov.CurrentValue = Province.GetKey("Prov") ' Prov
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Province.CurrentFilter = Province.KeyFilter
			Dim sSql
			sSql = Province.SQL
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
		ViewUrl = Province.ViewUrl
		EditUrl = Province.EditUrl("")
		InlineEditUrl = Province.InlineEditUrl
		CopyUrl = Province.CopyUrl("")
		InlineCopyUrl = Province.InlineCopyUrl
		DeleteUrl = Province.DeleteUrl

		' Call Row Rendering event
		Call Province.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Prov
		' Province
		' fProvince
		' TaxRate
		' ShipRate_first
		' ShipRate_Rest
		' -----------
		'  View  Row
		' -----------

		If Province.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Prov
			Province.Prov.ViewValue = Province.Prov.CurrentValue
			Province.Prov.ViewCustomAttributes = ""

			' Province
			Province.Province_1.ViewValue = Province.Province_1.CurrentValue
			Province.Province_1.ViewCustomAttributes = ""

			' fProvince
			Province.fProvince.ViewValue = Province.fProvince.CurrentValue
			Province.fProvince.ViewCustomAttributes = ""

			' TaxRate
			Province.TaxRate.ViewValue = Province.TaxRate.CurrentValue
			Province.TaxRate.ViewCustomAttributes = ""

			' ShipRate_first
			Province.ShipRate_first.ViewValue = Province.ShipRate_first.CurrentValue
			Province.ShipRate_first.ViewCustomAttributes = ""

			' ShipRate_Rest
			Province.ShipRate_Rest.ViewValue = Province.ShipRate_Rest.CurrentValue
			Province.ShipRate_Rest.ViewCustomAttributes = ""

			' View refer script
			' Prov

			Province.Prov.LinkCustomAttributes = ""
			Province.Prov.HrefValue = ""
			Province.Prov.TooltipValue = ""

			' Province
			Province.Province_1.LinkCustomAttributes = ""
			Province.Province_1.HrefValue = ""
			Province.Province_1.TooltipValue = ""

			' fProvince
			Province.fProvince.LinkCustomAttributes = ""
			Province.fProvince.HrefValue = ""
			Province.fProvince.TooltipValue = ""

			' TaxRate
			Province.TaxRate.LinkCustomAttributes = ""
			Province.TaxRate.HrefValue = ""
			Province.TaxRate.TooltipValue = ""

			' ShipRate_first
			Province.ShipRate_first.LinkCustomAttributes = ""
			Province.ShipRate_first.HrefValue = ""
			Province.ShipRate_first.TooltipValue = ""

			' ShipRate_Rest
			Province.ShipRate_Rest.LinkCustomAttributes = ""
			Province.ShipRate_Rest.HrefValue = ""
			Province.ShipRate_Rest.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Province.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Province.Row_Rendered()
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
