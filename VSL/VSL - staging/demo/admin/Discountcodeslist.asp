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
Dim Discountcodes_list
Set Discountcodes_list = New cDiscountcodes_list
Set Page = Discountcodes_list

' Page init processing
Call Discountcodes_list.Page_Init()

' Page main processing
Call Discountcodes_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Discountcodes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Discountcodes_list = new ew_Page("Discountcodes_list");
// page properties
Discountcodes_list.PageID = "list"; // page ID
Discountcodes_list.FormID = "fDiscountcodeslist"; // form ID
var EW_PAGE_ID = Discountcodes_list.PageID; // for backward compatibility
// extend page with validate function for search
Discountcodes_list.ValidateSearch = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (this.ValidateRequired) {
		var infix = "";
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
Discountcodes_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Discountcodes_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Discountcodes_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Discountcodes_list.ValidateRequired = false; // no JavaScript validation
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
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-1">
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% If (Discountcodes.Export = "") Or (EW_EXPORT_MASTER_RECORD And Discountcodes.Export = "print") Then %>
<%
gsMasterReturnUrl = "DiscountTypeslist.asp"
If Discountcodes_list.DbMasterFilter <> "" And Discountcodes.CurrentMasterTable = "DiscountTypes" Then
	If Discountcodes_list.MasterRecordExists Then
		If Discountcodes.CurrentMasterTable = Discountcodes.TableVar Then gsMasterReturnUrl = gsMasterReturnUrl & "?" & EW_TABLE_SHOW_MASTER & "="
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("MasterRecord") %><%= DiscountTypes.TableCaption %>
&nbsp;&nbsp;<% Discountcodes_list.ExportOptions.Render "body", "" %>
</p>
<% If Discountcodes.Export = "" Then %>
<p class="aspmaker"><a href="<%= gsMasterReturnUrl %>"><%= Language.Phrase("BackToMasterPage") %></a></p>
<% End If %>
<!--#include file="DiscountTypesmaster.asp"-->
<%
	End If
End If
%>
<% End If %>
<% Discountcodes_list.ShowPageHeader() %>
<%

' Load recordset
Set Discountcodes_list.Recordset = Discountcodes_list.LoadRecordset()
	Discountcodes_list.TotalRecs = Discountcodes_list.Recordset.RecordCount
	Discountcodes_list.StartRec = 1
	If Discountcodes_list.DisplayRecs <= 0 Then ' Display all records
		Discountcodes_list.DisplayRecs = Discountcodes_list.TotalRecs
	End If
	If Not (Discountcodes.ExportAll And Discountcodes.Export <> "") Then
		Discountcodes_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Discountcodes.TableCaption %>
<% If Discountcodes.CurrentMasterTable = "" Then %>
&nbsp;&nbsp;<% Discountcodes_list.ExportOptions.Render "body", "" %>
<% End If %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Discountcodes.Export = "" And Discountcodes.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Discountcodes_list);" style="text-decoration: none;"><img id="Discountcodes_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Discountcodes_list_SearchPanel">
<form name="fDiscountcodeslistsrch" id="fDiscountcodeslistsrch" class="ewForm" action="<%= ew_CurrentPage %>" onsubmit="return Discountcodes_list.ValidateSearch(this);">
<input type="hidden" id="t" name="t" value="Discountcodes">
<div class="ewBasicSearch">
<%
If gsSearchError = "" Then
	Call Discountcodes_list.LoadAdvancedSearch() ' Load advanced search
End If

' Render for search
Discountcodes.RowType = EW_ROWTYPE_SEARCH

' Render row
Call Discountcodes.ResetAttrs()
Call Discountcodes_list.RenderRow()
%>
<div id="xsr_1" class="ewCssTableRow">
	<span id="xsc_DiscountTypeId" class="ewCssTableCell">
		<span class="ewSearchCaption"><%= Discountcodes.DiscountTypeId.FldCaption %></span>
		<span class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_DiscountTypeId" id="z_DiscountTypeId" value="="></span>
		<span class="ewSearchField">
<% If Discountcodes.DiscountTypeId.SessionValue <> "" Then %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
<input type="hidden" id="x_DiscountTypeId" name="x_DiscountTypeId" value="<%= ew_HtmlEncode(Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue) %>">
<% Else %>
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
<% End If %>
</span>
	</span>
</div>
<div id="xsr_2" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Discountcodes.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Discountcodes_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
	<a href="Discountcodessrch.asp"><%= Language.Phrase("AdvancedSearch") %></a>&nbsp;
</div>
<div id="xsr_3" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Discountcodes.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Discountcodes.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Discountcodes.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Discountcodes_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Discountcodes.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Discountcodes.CurrentAction <> "gridadd" And Discountcodes.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Discountcodes_list.Pager) Then Set Discountcodes_list.Pager = ew_NewNumericPager(Discountcodes_list.StartRec, Discountcodes_list.DisplayRecs, Discountcodes_list.TotalRecs, Discountcodes_list.RecRange) %>
<% If Discountcodes_list.Pager.RecordCount > 0 Then %>
	<% If Discountcodes_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Discountcodes_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Discountcodes_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Discountcodes_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Discountcodes_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Discountcodes_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Discountcodes_list.Pager.RecordCount %>
<% Else %>
	<% If Discountcodes_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Discountcodes_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Discountcodes_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fDiscountcodeslist, '<%= Discountcodes_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fDiscountcodeslist" id="fDiscountcodeslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Discountcodes">
<div id="gmp_Discountcodes" class="ewGridMiddlePanel">
<% If Discountcodes_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Discountcodes.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Discountcodes_list.RenderListOptions()

' Render list options (header, left)
Discountcodes_list.ListOptions.Render "header", "left"
%>
<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
	<% If Discountcodes.SortUrl(Discountcodes.DiscountCode) = "" Then %>
		<td><%= Discountcodes.DiscountCode.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.DiscountCode) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.DiscountCode.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Discountcodes.DiscountCode.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.DiscountCode.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.Active.Visible Then ' Active %>
	<% If Discountcodes.SortUrl(Discountcodes.Active) = "" Then %>
		<td><%= Discountcodes.Active.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.Active) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.Active.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.Active.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.Active.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.used.Visible Then ' used %>
	<% If Discountcodes.SortUrl(Discountcodes.used) = "" Then %>
		<td><%= Discountcodes.used.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.used) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.used.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.used.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.used.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.OrderId.Visible Then ' OrderId %>
	<% If Discountcodes.SortUrl(Discountcodes.OrderId) = "" Then %>
		<td><%= Discountcodes.OrderId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.OrderId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.OrderId.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.OrderId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.OrderId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.Use_date.Visible Then ' Use_date %>
	<% If Discountcodes.SortUrl(Discountcodes.Use_date) = "" Then %>
		<td><%= Discountcodes.Use_date.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.Use_date) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.Use_date.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.Use_date.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.Use_date.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<% If Discountcodes.SortUrl(Discountcodes.DiscountTypeId) = "" Then %>
		<td><%= Discountcodes.DiscountTypeId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Discountcodes.SortUrl(Discountcodes.DiscountTypeId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.DiscountTypeId.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.DiscountTypeId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.DiscountTypeId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Discountcodes_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Discountcodes.ExportAll And Discountcodes.Export <> "") Then
	Discountcodes_list.StopRec = Discountcodes_list.TotalRecs
Else

	' Set the last record to display
	If Discountcodes_list.TotalRecs > Discountcodes_list.StartRec + Discountcodes_list.DisplayRecs - 1 Then
		Discountcodes_list.StopRec = Discountcodes_list.StartRec + Discountcodes_list.DisplayRecs - 1
	Else
		Discountcodes_list.StopRec = Discountcodes_list.TotalRecs
	End If
End If

' Move to first record
Discountcodes_list.RecCnt = Discountcodes_list.StartRec - 1
If Not Discountcodes_list.Recordset.Eof Then
	Discountcodes_list.Recordset.MoveFirst
	If Discountcodes_list.StartRec > 1 Then Discountcodes_list.Recordset.Move Discountcodes_list.StartRec - 1
ElseIf Not Discountcodes.AllowAddDeleteRow And Discountcodes_list.StopRec = 0 Then
	Discountcodes_list.StopRec = Discountcodes.GridAddRowCount
End If

' Initialize Aggregate
Discountcodes.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Discountcodes.ResetAttrs()
Call Discountcodes_list.RenderRow()
Discountcodes_list.RowCnt = 0

' Output date rows
Do While CLng(Discountcodes_list.RecCnt) < CLng(Discountcodes_list.StopRec)
	Discountcodes_list.RecCnt = Discountcodes_list.RecCnt + 1
	If CLng(Discountcodes_list.RecCnt) >= CLng(Discountcodes_list.StartRec) Then
		Discountcodes_list.RowCnt = Discountcodes_list.RowCnt + 1

	' Set up key count
	Discountcodes_list.KeyCount = Discountcodes_list.RowIndex
	Call Discountcodes.ResetAttrs()
	Discountcodes.CssClass = ""
	If Discountcodes.CurrentAction = "gridadd" Then
	Else
		Call Discountcodes_list.LoadRowValues(Discountcodes_list.Recordset) ' Load row values
	End If
	Discountcodes.RowType = EW_ROWTYPE_VIEW ' Render view
	Discountcodes.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Discountcodes_list.RenderRow()

	' Render list options
	Call Discountcodes_list.RenderListOptions()
%>
	<tr<%= Discountcodes.RowAttributes %>>
<%

' Render list options (body, left)
Discountcodes_list.ListOptions.Render "body", "left"
%>
	<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>>
<div<%= Discountcodes.DiscountCode.ViewAttributes %>><%= Discountcodes.DiscountCode.ListViewValue %></div>
<a name="<%= Discountcodes_list.PageObjName & "_row_" & Discountcodes_list.RowCnt %>" id="<%= Discountcodes_list.PageObjName & "_row_" & Discountcodes_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Discountcodes.Active.Visible Then ' Active %>
		<td<%= Discountcodes.Active.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.used.Visible Then ' used %>
		<td<%= Discountcodes.used.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.OrderId.Visible Then ' OrderId %>
		<td<%= Discountcodes.OrderId.CellAttributes %>>
<div<%= Discountcodes.OrderId.ViewAttributes %>>
<% If Discountcodes.OrderId.LinkAttributes <> "" Then %>
<a<%= Discountcodes.OrderId.LinkAttributes %>><%= Discountcodes.OrderId.ListViewValue %></a>
<% Else %>
<%= Discountcodes.OrderId.ListViewValue %>
<% End If %>
</div>
</td>
	<% End If %>
	<% If Discountcodes.Use_date.Visible Then ' Use_date %>
		<td<%= Discountcodes.Use_date.CellAttributes %>>
<div<%= Discountcodes.Use_date.ViewAttributes %>><%= Discountcodes.Use_date.ListViewValue %></div>
</td>
	<% End If %>
	<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Discountcodes_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Discountcodes.CurrentAction <> "gridadd" Then
		Discountcodes_list.Recordset.MoveNext()
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
Discountcodes_list.Recordset.Close
Set Discountcodes_list.Recordset = Nothing
%>
<% If Discountcodes_list.TotalRecs > 0 Then %>
<% If Discountcodes.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Discountcodes.CurrentAction <> "gridadd" And Discountcodes.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Discountcodes_list.Pager) Then Set Discountcodes_list.Pager = ew_NewNumericPager(Discountcodes_list.StartRec, Discountcodes_list.DisplayRecs, Discountcodes_list.TotalRecs, Discountcodes_list.RecRange) %>
<% If Discountcodes_list.Pager.RecordCount > 0 Then %>
	<% If Discountcodes_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Discountcodes_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Discountcodes_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Discountcodes_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Discountcodes_list.PageUrl %>start=<%= Discountcodes_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Discountcodes_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Discountcodes_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Discountcodes_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Discountcodes_list.Pager.RecordCount %>
<% Else %>
	<% If Discountcodes_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Discountcodes_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Discountcodes_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fDiscountcodeslist, '<%= Discountcodes_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Discountcodes.Export = "" And Discountcodes.CurrentAction = "" Then %>
<% End If %>
<%
Discountcodes_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Discountcodes.Export = "" Then %>
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
Set Discountcodes_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountcodes_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Discountcodes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Discountcodes_list"
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
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Discountcodesadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Discountcodesdelete.asp"
		MultiUpdateUrl = "Discountcodesupdate.asp"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(DiscountTypes) Then Set DiscountTypes = New cDiscountTypes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Discountcodes"

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
				Discountcodes.GridAddRowCount = gridaddcnt
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
		Set Discountcodes = Nothing
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
			If Discountcodes.Export <> "" Or Discountcodes.CurrentAction = "gridadd" Or Discountcodes.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Get and validate search values for advanced search
			Call LoadSearchValues() ' Get search values
			If ValidateSearch() Then

				' Nothing to do
			Else
				FailureMessage = gsSearchError
			End If

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Discountcodes.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If

			' Get search criteria for advanced search
			If gsSearchError = "" Then
				sSrchAdvanced = AdvancedSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Discountcodes.RecordsPerPage <> "" Then
			DisplayRecs = Discountcodes.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Discountcodes.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			If sSrchAdvanced = "" Then Call ResetAdvancedSearchParms()
			Discountcodes.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Discountcodes.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Discountcodes.SearchWhere
		End If
		sFilter = ""

		' Restore master/detail filter
		DbMasterFilter = Discountcodes.MasterFilter ' Restore master filter
		DbDetailFilter = Discountcodes.DetailFilter ' Restore detail filter
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)
		Dim RsMaster

		' Load master record
		If Discountcodes.MasterFilter <> "" And Discountcodes.CurrentMasterTable = "DiscountTypes" Then
			Set RsMaster = DiscountTypes.LoadRs(DbMasterFilter)
			MasterRecordExists = Not (RsMaster Is Nothing)
			If Not MasterRecordExists Then
				FailureMessage = Language.Phrase("NoRecord") ' Set no record found
				Call Page_Terminate(Discountcodes.ReturnUrl) ' Return to caller
			Else
				Call DiscountTypes.LoadListRowValues(RsMaster)
				DiscountTypes.RowType = EW_ROWTYPE_MASTER ' Master row
				Call DiscountTypes.RenderListRow()
				RsMaster.Close
				Set RsMaster = Nothing
			End If
		End If

		' Set up filter in Session
		Discountcodes.SessionWhere = sFilter
		Discountcodes.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Advanced Search Where based on QueryString parameters
	'
	Function AdvancedSearchWhere()
		Dim sWhere
		sWhere = ""

		' Field Discountid
		Call BuildSearchSql(sWhere, Discountcodes.Discountid, False)

		' Field DiscountCode
		Call BuildSearchSql(sWhere, Discountcodes.DiscountCode, False)

		' Field Active
		Call BuildSearchSql(sWhere, Discountcodes.Active, False)

		' Field used
		Call BuildSearchSql(sWhere, Discountcodes.used, False)

		' Field OrderId
		Call BuildSearchSql(sWhere, Discountcodes.OrderId, False)

		' Field Use_date
		Call BuildSearchSql(sWhere, Discountcodes.Use_date, False)

		' Field DiscountTypeId
		Call BuildSearchSql(sWhere, Discountcodes.DiscountTypeId, False)
		AdvancedSearchWhere = sWhere

		' Set up search parm
		If AdvancedSearchWhere <> "" Then

			' Field Discountid
			Call SetSearchParm(Discountcodes.Discountid)

			' Field DiscountCode
			Call SetSearchParm(Discountcodes.DiscountCode)

			' Field Active
			Call SetSearchParm(Discountcodes.Active)

			' Field used
			Call SetSearchParm(Discountcodes.used)

			' Field OrderId
			Call SetSearchParm(Discountcodes.OrderId)

			' Field Use_date
			Call SetSearchParm(Discountcodes.Use_date)

			' Field DiscountTypeId
			Call SetSearchParm(Discountcodes.DiscountTypeId)
		End If
	End Function

	' -----------------------------------------------------------------
	' Build search sql
	'
	Sub BuildSearchSql(Where, Fld, MultiValue)
		Dim FldParm, FldVal, FldOpr, FldCond, FldVal2, FldOpr2
		FldParm = Mid(Fld.FldVar, 3)
		FldVal = Fld.AdvancedSearch.SearchValue
		FldOpr = Fld.AdvancedSearch.SearchOperator
		FldCond = Fld.AdvancedSearch.SearchCondition
		FldVal2 = Fld.AdvancedSearch.SearchValue2
		FldOpr2 = Fld.AdvancedSearch.SearchOperator2
		Dim sWrk
		sWrk = ""
		FldOpr = UCase(Trim(FldOpr))
		If (FldOpr = "") Then FldOpr = "="
		FldOpr2 = UCase(Trim(FldOpr2))
		If FldOpr2 = "" Then FldOpr2 = "="
		If EW_SEARCH_MULTI_VALUE_OPTION = 1 Then MultiValue = False
		If FldOpr <> "LIKE" Then MultiValue = False
		If FldOpr2 <> "LIKE" And FldVal2 <> "" Then MultiValue = False
		If MultiValue Then
			Dim sWrk1, sWrk2

			' Field value 1
			If FldVal <> "" Then
				sWrk1 = ew_GetMultiSearchSql(Fld, FldVal)
			Else
				sWrk1 = ""
			End If

			' Field value 2
			If FldVal2 <> "" And FldCond <> "" Then
				sWrk2 = ew_GetMultiSearchSql(Fld, FldVal2)
			Else
				sWrk2 = ""
			End If

			' Build final SQL
			sWrk = sWrk1
			If sWrk2 <> "" Then
				If sWrk <> "" Then
					sWrk = "(" & sWrk & ") " & FldCond & " (" & sWrk2 & ")"
				Else
					sWrk = sWrk2
				End If
			End If
		Else
			FldVal = ConvertSearchValue(Fld, FldVal)
			FldVal2 = ConvertSearchValue(Fld, FldVal2)
			sWrk = ew_GetSearchSql(Fld, FldVal, FldOpr, FldCond, FldVal2, FldOpr2)
		End If
		Call ew_AddFilter(Where, sWrk)
	End Sub

	' -----------------------------------------------------------------
	' Set search parm
	'
	Sub SetSearchParm(Fld)
		Dim FldParm
		FldParm = Mid(Fld.FldVar, 3)
		Call Discountcodes.SetAdvancedSearch("x_" & FldParm, Fld.AdvancedSearch.SearchValue)
		Call Discountcodes.SetAdvancedSearch("z_" & FldParm, Fld.AdvancedSearch.SearchOperator)
		Call Discountcodes.SetAdvancedSearch("v_" & FldParm, Fld.AdvancedSearch.SearchCondition)
		Call Discountcodes.SetAdvancedSearch("y_" & FldParm, Fld.AdvancedSearch.SearchValue2)
		Call Discountcodes.SetAdvancedSearch("w_" & FldParm, Fld.AdvancedSearch.SearchOperator2)
	End Sub

	' -----------------------------------------------------------------
	' Get search parm
	'
	Sub GetSearchParm(Fld)
		Dim FldParm
		FldParm = Mid(Fld.FldVar, 3)
		Fld.AdvancedSearch.SearchValue = Discountcodes.GetAdvancedSearch("x_" & FldParm)
		Fld.AdvancedSearch.SearchOperator = Discountcodes.GetAdvancedSearch("z_" & FldParm)
		Fld.AdvancedSearch.SearchCondition = Discountcodes.GetAdvancedSearch("v_" & FldParm)
		Fld.AdvancedSearch.SearchValue2 = Discountcodes.GetAdvancedSearch("y_" & FldParm)
		Fld.AdvancedSearch.SearchOperator2 = Discountcodes.GetAdvancedSearch("w_" & FldParm)
	End Sub

	' -----------------------------------------------------------------
	' Convert search value
	'
	Function ConvertSearchValue(Fld, FldVal)
		ConvertSearchValue = FldVal
		If Fld.FldDataType = EW_DATATYPE_BOOLEAN Then
			If FldVal <> "" Then ConvertSearchValue = ew_IIf(FldVal&"" = "1", "True", "False")
		ElseIf Fld.FldDataType = EW_DATATYPE_DATE Then
			If FldVal <> "" Then ConvertSearchValue = ew_UnFormatDateTime(FldVal, Fld.FldDateTimeFormat)
		End If
	End Function

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Discountcodes.DiscountCode, Keyword)
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
		sSearchKeyword = Discountcodes.BasicSearchKeyword
		sSearchType = Discountcodes.BasicSearchType
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
			Discountcodes.SessionBasicSearchKeyword = sSearchKeyword
			Discountcodes.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Discountcodes.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()

		' Clear advanced search parameters
		Call ResetAdvancedSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Discountcodes.SessionBasicSearchKeyword = ""
		Discountcodes.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Clear all advanced search parameters
	'
	Sub ResetAdvancedSearchParms()

		' Clear advanced search parameters
		Call Discountcodes.SetAdvancedSearch("x_Discountid", "")
		Call Discountcodes.SetAdvancedSearch("x_DiscountCode", "")
		Call Discountcodes.SetAdvancedSearch("x_Active", "")
		Call Discountcodes.SetAdvancedSearch("x_used", "")
		Call Discountcodes.SetAdvancedSearch("x_OrderId", "")
		Call Discountcodes.SetAdvancedSearch("x_Use_date", "")
		Call Discountcodes.SetAdvancedSearch("x_DiscountTypeId", "")
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Discountcodes.BasicSearchKeyword & "" <> "" Then bRestore = False
		If Discountcodes.Discountid.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Discountcodes.DiscountCode.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Discountcodes.Active.AdvancedSearch.SearchValue & "" <> Discountcodes.GetAdvancedSearch("x_Active") & "" Then bRestore = False
		If Discountcodes.used.AdvancedSearch.SearchValue & "" <> Discountcodes.GetAdvancedSearch("x_used") & "" Then bRestore = False
		If Discountcodes.OrderId.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Discountcodes.Use_date.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Discountcodes.BasicSearchKeyword = Discountcodes.SessionBasicSearchKeyword
			Discountcodes.BasicSearchType = Discountcodes.SessionBasicSearchType

			' Restore advanced search values
			Call GetSearchParm(Discountcodes.Discountid)
			Call GetSearchParm(Discountcodes.DiscountCode)
			Call GetSearchParm(Discountcodes.Active)
			Call GetSearchParm(Discountcodes.used)
			Call GetSearchParm(Discountcodes.OrderId)
			Call GetSearchParm(Discountcodes.Use_date)
			Call GetSearchParm(Discountcodes.DiscountTypeId)
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
			Discountcodes.CurrentOrder = Request.QueryString("order")
			Discountcodes.CurrentOrderType = Request.QueryString("ordertype")

			' Field DiscountCode
			Call Discountcodes.UpdateSort(Discountcodes.DiscountCode)

			' Field Active
			Call Discountcodes.UpdateSort(Discountcodes.Active)

			' Field used
			Call Discountcodes.UpdateSort(Discountcodes.used)

			' Field OrderId
			Call Discountcodes.UpdateSort(Discountcodes.OrderId)

			' Field Use_date
			Call Discountcodes.UpdateSort(Discountcodes.Use_date)

			' Field DiscountTypeId
			Call Discountcodes.UpdateSort(Discountcodes.DiscountTypeId)
			Discountcodes.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Discountcodes.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Discountcodes.SqlOrderBy <> "" Then
				sOrderBy = Discountcodes.SqlOrderBy
				Discountcodes.SessionOrderBy = sOrderBy
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

			' Reset master/detail keys
			If LCase(sCmd) = "resetall" Then
				Discountcodes.CurrentMasterTable = "" ' Clear master table
				DbMasterFilter = ""
				DbDetailFilter = ""
				Discountcodes.DiscountTypeId.SessionValue = ""
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				Discountcodes.SessionOrderBy = sOrderBy
				Discountcodes.DiscountCode.Sort = ""
				Discountcodes.Active.Sort = ""
				Discountcodes.used.Sort = ""
				Discountcodes.OrderId.Sort = ""
				Discountcodes.Use_date.Sort = ""
				Discountcodes.DiscountTypeId.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Discountcodes.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		ListOptions.Add("edit")
		ListOptions.GetItem("edit").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("edit").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("edit").OnLeft = True
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("checkbox").OnLeft = True
		ListOptions.MoveItem "checkbox", 0 ' Move to first column
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Discountcodes_list.SelectAllKey(this);"">"
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.IsLoggedIn() And ListOptions.GetItem("edit").Visible Then
			Set item = ListOptions.GetItem("edit")
			item.Body = "<a class=""ewRowLink"" href=""" & EditUrl & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
		End If
		If Security.IsLoggedIn() And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Discountcodes.Discountid.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Discountcodes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Discountcodes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Discountcodes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Discountcodes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Discountcodes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Discountcodes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Discountcodes.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Discountcodes.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	'  Load search values for validation
	'
	Function LoadSearchValues()

		' Load search values
		Discountcodes.Discountid.AdvancedSearch.SearchValue = Request.QueryString("x_Discountid")
		Discountcodes.Discountid.AdvancedSearch.SearchOperator = Request.QueryString("z_Discountid")
		Discountcodes.DiscountCode.AdvancedSearch.SearchValue = Request.QueryString("x_DiscountCode")
		Discountcodes.DiscountCode.AdvancedSearch.SearchOperator = Request.QueryString("z_DiscountCode")
		Discountcodes.Active.AdvancedSearch.SearchValue = Request.QueryString("x_Active")
		Discountcodes.Active.AdvancedSearch.SearchOperator = Request.QueryString("z_Active")
		Discountcodes.used.AdvancedSearch.SearchValue = Request.QueryString("x_used")
		Discountcodes.used.AdvancedSearch.SearchOperator = Request.QueryString("z_used")
		Discountcodes.OrderId.AdvancedSearch.SearchValue = Request.QueryString("x_OrderId")
		Discountcodes.OrderId.AdvancedSearch.SearchOperator = Request.QueryString("z_OrderId")
		Discountcodes.Use_date.AdvancedSearch.SearchValue = Request.QueryString("x_Use_date")
		Discountcodes.Use_date.AdvancedSearch.SearchOperator = Request.QueryString("z_Use_date")
		Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue = Request.QueryString("x_DiscountTypeId")
		Discountcodes.DiscountTypeId.AdvancedSearch.SearchOperator = Request.QueryString("z_DiscountTypeId")
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Discountcodes.CurrentFilter
		Call Discountcodes.Recordset_Selecting(sFilter)
		Discountcodes.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Discountcodes.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Discountcodes.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Discountcodes.KeyFilter

		' Call Row Selecting event
		Call Discountcodes.Row_Selecting(sFilter)

		' Load sql based on filter
		Discountcodes.CurrentFilter = sFilter
		sSql = Discountcodes.SQL
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
		Call Discountcodes.Row_Selected(RsRow)
		Discountcodes.Discountid.DbValue = RsRow("Discountid")
		Discountcodes.DiscountCode.DbValue = RsRow("DiscountCode")
		Discountcodes.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Discountcodes.used.DbValue = ew_IIf(RsRow("used"), "1", "0")
		Discountcodes.OrderId.DbValue = RsRow("OrderId")
		Discountcodes.Use_date.DbValue = RsRow("Use_date")
		Discountcodes.DiscountTypeId.DbValue = RsRow("DiscountTypeId")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Discountcodes.GetKey("Discountid")&"" <> "" Then
			Discountcodes.Discountid.CurrentValue = Discountcodes.GetKey("Discountid") ' Discountid
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Discountcodes.CurrentFilter = Discountcodes.KeyFilter
			Dim sSql
			sSql = Discountcodes.SQL
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
		ViewUrl = Discountcodes.ViewUrl
		EditUrl = Discountcodes.EditUrl("")
		InlineEditUrl = Discountcodes.InlineEditUrl
		CopyUrl = Discountcodes.CopyUrl("")
		InlineCopyUrl = Discountcodes.InlineCopyUrl
		DeleteUrl = Discountcodes.DeleteUrl

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
			If Discountcodes.DiscountTypeId.SessionValue <> "" Then
				Discountcodes.DiscountTypeId.AdvancedSearch.SearchValue = Discountcodes.DiscountTypeId.SessionValue
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
			Else
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
			If sMasterTblVar = "DiscountTypes" Then
				bValidMaster = True
				If Request.QueryString("DiscountTypeId").Count > 0 Then
					DiscountTypes.DiscountTypeId.QueryStringValue = Request.QueryString("DiscountTypeId")
					Discountcodes.DiscountTypeId.QueryStringValue = DiscountTypes.DiscountTypeId.QueryStringValue
					Discountcodes.DiscountTypeId.SessionValue = Discountcodes.DiscountTypeId.QueryStringValue
					If Not IsNumeric(DiscountTypes.DiscountTypeId.QueryStringValue) Then bValidMaster = False
				Else
					bValidMaster = False
				End If
			End If
		End If
		If bValidMaster Then

			' Save current master table
			Discountcodes.CurrentMasterTable = sMasterTblVar

			' Reset start record counter (new master key)
			StartRec = 1
			Discountcodes.StartRecordNumber = StartRec

			' Clear previous master session values
			If sMasterTblVar <> "DiscountTypes" Then
				If Discountcodes.DiscountTypeId.QueryStringValue = "" Then Discountcodes.DiscountTypeId.SessionValue = ""
			End If
		End If
		DbMasterFilter = Discountcodes.MasterFilter '  Get master filter
		DbDetailFilter = Discountcodes.DetailFilter ' Get detail filter
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
