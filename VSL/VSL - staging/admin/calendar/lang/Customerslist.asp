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
Dim Customers_list
Set Customers_list = New cCustomers_list
Set Page = Customers_list

' Page init processing
Call Customers_list.Page_Init()

' Page main processing
Call Customers_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Customers.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Customers_list = new ew_Page("Customers_list");
// page properties
Customers_list.PageID = "list"; // page ID
Customers_list.FormID = "fCustomerslist"; // form ID
var EW_PAGE_ID = Customers_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Customers_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Customers_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Customers_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Customers_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Customers.Export = "") Or (EW_EXPORT_MASTER_RECORD And Customers.Export = "print") Then %>
<% End If %>
<% Customers_list.ShowPageHeader() %>
<%

' Load recordset
Set Customers_list.Recordset = Customers_list.LoadRecordset()
	Customers_list.TotalRecs = Customers_list.Recordset.RecordCount
	Customers_list.StartRec = 1
	If Customers_list.DisplayRecs <= 0 Then ' Display all records
		Customers_list.DisplayRecs = Customers_list.TotalRecs
	End If
	If Not (Customers.ExportAll And Customers.Export <> "") Then
		Customers_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Customers.TableCaption %>
&nbsp;&nbsp;<% Customers_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Customers.Export = "" And Customers.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Customers_list);" style="text-decoration: none;"><img id="Customers_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Customers_list_SearchPanel">
<form name="fCustomerslistsrch" id="fCustomerslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Customers">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Customers.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Customers_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Customers.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Customers.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Customers.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Customers_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Customers.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Customers.CurrentAction <> "gridadd" And Customers.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Customers_list.Pager) Then Set Customers_list.Pager = ew_NewNumericPager(Customers_list.StartRec, Customers_list.DisplayRecs, Customers_list.TotalRecs, Customers_list.RecRange) %>
<% If Customers_list.Pager.RecordCount > 0 Then %>
	<% If Customers_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Customers_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Customers_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Customers_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Customers_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Customers_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Customers_list.Pager.RecordCount %>
<% Else %>
	<% If Customers_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Customers_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Customers_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fCustomerslist, '<%= Customers_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fCustomerslist" id="fCustomerslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Customers">
<div id="gmp_Customers" class="ewGridMiddlePanel">
<% If Customers_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Customers.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Customers_list.RenderListOptions()

' Render list options (header, left)
Customers_list.ListOptions.Render "header", "left"
%>
<% If Customers.Inv_FirstName.Visible Then ' Inv_FirstName %>
	<% If Customers.SortUrl(Customers.Inv_FirstName) = "" Then %>
		<td><%= Customers.Inv_FirstName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.Inv_FirstName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.Inv_FirstName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.Inv_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.Inv_LastName.Visible Then ' Inv_LastName %>
	<% If Customers.SortUrl(Customers.Inv_LastName) = "" Then %>
		<td><%= Customers.Inv_LastName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.Inv_LastName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.Inv_LastName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.Inv_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.Inv_Address.Visible Then ' Inv_Address %>
	<% If Customers.SortUrl(Customers.Inv_Address) = "" Then %>
		<td><%= Customers.Inv_Address.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.Inv_Address) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.Inv_Address.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.Inv_Address.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_Address.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.inv_City.Visible Then ' inv_City %>
	<% If Customers.SortUrl(Customers.inv_City) = "" Then %>
		<td><%= Customers.inv_City.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.inv_City) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.inv_City.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.inv_City.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_City.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
	<% If Customers.SortUrl(Customers.inv_EmailAddress) = "" Then %>
		<td><%= Customers.inv_EmailAddress.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.inv_EmailAddress) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.inv_EmailAddress.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.inv_EmailAddress.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_EmailAddress.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.UserName.Visible Then ' UserName %>
	<% If Customers.SortUrl(Customers.UserName) = "" Then %>
		<td><%= Customers.UserName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.UserName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.UserName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Customers.UserName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.UserName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Customers.NewCustomer.Visible Then ' NewCustomer %>
	<% If Customers.SortUrl(Customers.NewCustomer) = "" Then %>
		<td><%= Customers.NewCustomer.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Customers.SortUrl(Customers.NewCustomer) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Customers.NewCustomer.FldCaption %></td><td style="width: 10px;"><% If Customers.NewCustomer.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.NewCustomer.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Customers_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Customers.ExportAll And Customers.Export <> "") Then
	Customers_list.StopRec = Customers_list.TotalRecs
Else

	' Set the last record to display
	If Customers_list.TotalRecs > Customers_list.StartRec + Customers_list.DisplayRecs - 1 Then
		Customers_list.StopRec = Customers_list.StartRec + Customers_list.DisplayRecs - 1
	Else
		Customers_list.StopRec = Customers_list.TotalRecs
	End If
End If

' Move to first record
Customers_list.RecCnt = Customers_list.StartRec - 1
If Not Customers_list.Recordset.Eof Then
	Customers_list.Recordset.MoveFirst
	If Customers_list.StartRec > 1 Then Customers_list.Recordset.Move Customers_list.StartRec - 1
ElseIf Not Customers.AllowAddDeleteRow And Customers_list.StopRec = 0 Then
	Customers_list.StopRec = Customers.GridAddRowCount
End If

' Initialize Aggregate
Customers.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Customers.ResetAttrs()
Call Customers_list.RenderRow()
Customers_list.RowCnt = 0

' Output date rows
Do While CLng(Customers_list.RecCnt) < CLng(Customers_list.StopRec)
	Customers_list.RecCnt = Customers_list.RecCnt + 1
	If CLng(Customers_list.RecCnt) >= CLng(Customers_list.StartRec) Then
		Customers_list.RowCnt = Customers_list.RowCnt + 1

	' Set up key count
	Customers_list.KeyCount = Customers_list.RowIndex
	Call Customers.ResetAttrs()
	Customers.CssClass = ""
	If Customers.CurrentAction = "gridadd" Then
	Else
		Call Customers_list.LoadRowValues(Customers_list.Recordset) ' Load row values
	End If
	Customers.RowType = EW_ROWTYPE_VIEW ' Render view
	Customers.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Customers_list.RenderRow()

	' Render list options
	Call Customers_list.RenderListOptions()
%>
	<tr<%= Customers.RowAttributes %>>
<%

' Render list options (body, left)
Customers_list.ListOptions.Render "body", "left"
%>
	<% If Customers.Inv_FirstName.Visible Then ' Inv_FirstName %>
		<td<%= Customers.Inv_FirstName.CellAttributes %>>
<div<%= Customers.Inv_FirstName.ViewAttributes %>><%= Customers.Inv_FirstName.ListViewValue %></div>
<a name="<%= Customers_list.PageObjName & "_row_" & Customers_list.RowCnt %>" id="<%= Customers_list.PageObjName & "_row_" & Customers_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Customers.Inv_LastName.Visible Then ' Inv_LastName %>
		<td<%= Customers.Inv_LastName.CellAttributes %>>
<div<%= Customers.Inv_LastName.ViewAttributes %>><%= Customers.Inv_LastName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Customers.Inv_Address.Visible Then ' Inv_Address %>
		<td<%= Customers.Inv_Address.CellAttributes %>>
<div<%= Customers.Inv_Address.ViewAttributes %>><%= Customers.Inv_Address.ListViewValue %></div>
</td>
	<% End If %>
	<% If Customers.inv_City.Visible Then ' inv_City %>
		<td<%= Customers.inv_City.CellAttributes %>>
<div<%= Customers.inv_City.ViewAttributes %>><%= Customers.inv_City.ListViewValue %></div>
</td>
	<% End If %>
	<% If Customers.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
		<td<%= Customers.inv_EmailAddress.CellAttributes %>>
<div<%= Customers.inv_EmailAddress.ViewAttributes %>><%= Customers.inv_EmailAddress.ListViewValue %></div>
</td>
	<% End If %>
	<% If Customers.UserName.Visible Then ' UserName %>
		<td<%= Customers.UserName.CellAttributes %>>
<div<%= Customers.UserName.ViewAttributes %>><%= Customers.UserName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Customers.NewCustomer.Visible Then ' NewCustomer %>
		<td<%= Customers.NewCustomer.CellAttributes %>>
<% If ew_ConvertToBool(Customers.NewCustomer.CurrentValue) Then %>
<input type="checkbox" value="<%= Customers.NewCustomer.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Customers.NewCustomer.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
<%

' Render list options (body, right)
Customers_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Customers.CurrentAction <> "gridadd" Then
		Customers_list.Recordset.MoveNext()
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
Customers_list.Recordset.Close
Set Customers_list.Recordset = Nothing
%>
<% If Customers_list.TotalRecs > 0 Then %>
<% If Customers.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Customers.CurrentAction <> "gridadd" And Customers.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Customers_list.Pager) Then Set Customers_list.Pager = ew_NewNumericPager(Customers_list.StartRec, Customers_list.DisplayRecs, Customers_list.TotalRecs, Customers_list.RecRange) %>
<% If Customers_list.Pager.RecordCount > 0 Then %>
	<% If Customers_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Customers_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Customers_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Customers_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Customers_list.PageUrl %>start=<%= Customers_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Customers_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Customers_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Customers_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Customers_list.Pager.RecordCount %>
<% Else %>
	<% If Customers_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Customers_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Customers_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fCustomerslist, '<%= Customers_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Customers.Export = "" And Customers.CurrentAction = "" Then %>
<% End If %>
<%
Customers_list.ShowPageFooter()
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
Set Customers_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cCustomers_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Customers"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Customers_list"
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
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Customersadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Customersdelete.asp"
		MultiUpdateUrl = "Customersupdate.asp"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Customers"

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
				Customers.GridAddRowCount = gridaddcnt
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
		Set Customers = Nothing
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
			If Customers.Export <> "" Or Customers.CurrentAction = "gridadd" Or Customers.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Customers.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Customers.RecordsPerPage <> "" Then
			DisplayRecs = Customers.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Customers.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Customers.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Customers.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Customers.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Customers.SessionWhere = sFilter
		Customers.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Customers.Inv_FirstName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.Inv_LastName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.Inv_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_PostalCode, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_PhoneNumber, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_EmailAddress, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.Notes, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.inv_Fax, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.Inv_Address2, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.UserName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Customers.passwrd, Keyword)
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
		sSearchKeyword = Customers.BasicSearchKeyword
		sSearchType = Customers.BasicSearchType
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
			Customers.SessionBasicSearchKeyword = sSearchKeyword
			Customers.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Customers.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Customers.SessionBasicSearchKeyword = ""
		Customers.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Customers.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Customers.BasicSearchKeyword = Customers.SessionBasicSearchKeyword
			Customers.BasicSearchType = Customers.SessionBasicSearchType
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
			Customers.CurrentOrder = Request.QueryString("order")
			Customers.CurrentOrderType = Request.QueryString("ordertype")

			' Field Inv_FirstName
			Call Customers.UpdateSort(Customers.Inv_FirstName)

			' Field Inv_LastName
			Call Customers.UpdateSort(Customers.Inv_LastName)

			' Field Inv_Address
			Call Customers.UpdateSort(Customers.Inv_Address)

			' Field inv_City
			Call Customers.UpdateSort(Customers.inv_City)

			' Field inv_EmailAddress
			Call Customers.UpdateSort(Customers.inv_EmailAddress)

			' Field UserName
			Call Customers.UpdateSort(Customers.UserName)

			' Field NewCustomer
			Call Customers.UpdateSort(Customers.NewCustomer)
			Customers.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Customers.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Customers.SqlOrderBy <> "" Then
				sOrderBy = Customers.SqlOrderBy
				Customers.SessionOrderBy = sOrderBy
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
				Customers.SessionOrderBy = sOrderBy
				Customers.Inv_FirstName.Sort = ""
				Customers.Inv_LastName.Sort = ""
				Customers.Inv_Address.Sort = ""
				Customers.inv_City.Sort = ""
				Customers.inv_EmailAddress.Sort = ""
				Customers.UserName.Sort = ""
				Customers.NewCustomer.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Customers.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Customers_list.SelectAllKey(this);"">"
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
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Customers.CustomerID.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Customers.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Customers.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Customers.CurrentFilter
		Call Customers.Recordset_Selecting(sFilter)
		Customers.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Customers.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Customers.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Customers.GetKey("CustomerID")&"" <> "" Then
			Customers.CustomerID.CurrentValue = Customers.GetKey("CustomerID") ' CustomerID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Customers.CurrentFilter = Customers.KeyFilter
			Dim sSql
			sSql = Customers.SQL
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
		ViewUrl = Customers.ViewUrl
		EditUrl = Customers.EditUrl("")
		InlineEditUrl = Customers.InlineEditUrl
		CopyUrl = Customers.CopyUrl("")
		InlineCopyUrl = Customers.InlineCopyUrl
		DeleteUrl = Customers.DeleteUrl

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

			' inv_EmailAddress
			Customers.inv_EmailAddress.LinkCustomAttributes = ""
			Customers.inv_EmailAddress.HrefValue = ""
			Customers.inv_EmailAddress.TooltipValue = ""

			' UserName
			Customers.UserName.LinkCustomAttributes = ""
			Customers.UserName.HrefValue = ""
			Customers.UserName.TooltipValue = ""

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
