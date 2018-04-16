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
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim DiscountTypes_list
Set DiscountTypes_list = New cDiscountTypes_list
Set Page = DiscountTypes_list

' Page init processing
Call DiscountTypes_list.Page_Init()

' Page main processing
Call DiscountTypes_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If DiscountTypes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var DiscountTypes_list = new ew_Page("DiscountTypes_list");
// page properties
DiscountTypes_list.PageID = "list"; // page ID
DiscountTypes_list.FormID = "fDiscountTypeslist"; // form ID
var EW_PAGE_ID = DiscountTypes_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
DiscountTypes_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
DiscountTypes_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
DiscountTypes_list.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<style>
/* styles for detail preview panel */
#ewDetailsDiv.yui-overlay { position:absolute;background:#fff;border:2px solid orange;padding:4px;margin:10px; }
#ewDetailsDiv.yui-overlay .hd { border:1px solid red;padding:5px; }
#ewDetailsDiv.yui-overlay .bd { border:0px solid green;padding:5px; }
#ewDetailsDiv.yui-overlay .ft { border:1px solid blue;padding:5px; }
</style>
<div id="ewDetailsDiv" style="visibility: hidden; z-index: 11000;" name="ewDetailsDivDiv"></div>
<script language="JavaScript" type="text/javascript">
<!--
// YUI container
var ewDetailsDiv;
var ew_AjaxDetailsTimer = null;
// init details div
function ew_InitDetailsDiv() {
	ewDetailsDiv = new ewWidget.Overlay("ewDetailsDiv", { context:null, visible:false} );
	//ewDetailsDiv.beforeMoveEvent.subscribe(ew_EnforceConstraints, ewDetailsDiv, true); //'***A8*** to be removed
	ewDetailsDiv.render();
}
// init details div on window.load
ewEvent.addListener(window, "load", ew_InitDetailsDiv);
// show results in details div
var ew_AjaxHandleSuccess = function(o) {
	if (ewDetailsDiv && o.responseText !== undefined) {
		ewDetailsDiv.cfg.applyConfig({context:[o.argument.id,o.argument.elcorner,o.argument.ctxcorner], visible:false, constraintoviewport:true, preventcontextoverlap:true}, true);
		ewDetailsDiv.setBody(o.responseText);
		ewDetailsDiv.render();
		ew_SetupTable(ewDom.get("ewDetailsPreviewTable"));
		ewDetailsDiv.show();
	}
}
// show error in details div
var ew_AjaxHandleFailure = function(o) {
	if (ewDetailsDiv && o.responseText != "") {
		ewDetailsDiv.cfg.applyConfig({context:[o.argument.id,o.argument.elcorner,o.argument.ctxcorner], visible:false, constraintoviewport:true, preventcontextoverlap:true}, true);
		ewDetailsDiv.setBody(o.responseText);
		ewDetailsDiv.render();
		ewDetailsDiv.show();
	}
}
// show details div
function ew_AjaxShowDetails(obj, url) {
	if (ew_AjaxDetailsTimer)
		clearTimeout(ew_AjaxDetailsTimer);
	ew_AjaxDetailsTimer = setTimeout(function() { ewConnect.asyncRequest('GET', url, {success: ew_AjaxHandleSuccess , failure: ew_AjaxHandleFailure, argument:{id: obj.id, elcorner: "tl", ctxcorner: "tr"}}) }, 200);
}
// hide details div
function ew_AjaxHideDetails(obj) {
	if (ew_AjaxDetailsTimer)
		clearTimeout(ew_AjaxDetailsTimer);
	if (ewDetailsDiv)
		ewDetailsDiv.hide();
}
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
<% If (DiscountTypes.Export = "") Or (EW_EXPORT_MASTER_RECORD And DiscountTypes.Export = "print") Then %>
<% End If %>
<% DiscountTypes_list.ShowPageHeader() %>
<%

' Load recordset
Set DiscountTypes_list.Recordset = DiscountTypes_list.LoadRecordset()
	DiscountTypes_list.TotalRecs = DiscountTypes_list.Recordset.RecordCount
	DiscountTypes_list.StartRec = 1
	If DiscountTypes_list.DisplayRecs <= 0 Then ' Display all records
		DiscountTypes_list.DisplayRecs = DiscountTypes_list.TotalRecs
	End If
	If Not (DiscountTypes.ExportAll And DiscountTypes.Export <> "") Then
		DiscountTypes_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= DiscountTypes.TableCaption %>
&nbsp;&nbsp;<% DiscountTypes_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If DiscountTypes.Export = "" And DiscountTypes.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(DiscountTypes_list);" style="text-decoration: none;"><img id="DiscountTypes_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="DiscountTypes_list_SearchPanel">
<form name="fDiscountTypeslistsrch" id="fDiscountTypeslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="DiscountTypes">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(DiscountTypes.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= DiscountTypes_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If DiscountTypes.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If DiscountTypes.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If DiscountTypes.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% DiscountTypes_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If DiscountTypes.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If DiscountTypes.CurrentAction <> "gridadd" And DiscountTypes.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(DiscountTypes_list.Pager) Then Set DiscountTypes_list.Pager = ew_NewNumericPager(DiscountTypes_list.StartRec, DiscountTypes_list.DisplayRecs, DiscountTypes_list.TotalRecs, DiscountTypes_list.RecRange) %>
<% If DiscountTypes_list.Pager.RecordCount > 0 Then %>
	<% If DiscountTypes_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In DiscountTypes_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= DiscountTypes_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If DiscountTypes_list.Pager.NextButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.LastButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= DiscountTypes_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= DiscountTypes_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= DiscountTypes_list.Pager.RecordCount %>
<% Else %>
	<% If DiscountTypes_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= DiscountTypes_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% If Discountcodes.DetailAdd And Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= DiscountTypes.AddUrl & "?" & EW_TABLE_SHOW_DETAIL & "=Discountcodes" %>"><%= Language.Phrase("AddLink") %>&nbsp;<%= DiscountTypes.TableCaption %>/<%= Discountcodes.TableCaption %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fDiscountTypeslist" id="fDiscountTypeslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="DiscountTypes">
<div id="gmp_DiscountTypes" class="ewGridMiddlePanel">
<% If DiscountTypes_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= DiscountTypes.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call DiscountTypes_list.RenderListOptions()

' Render list options (header, left)
DiscountTypes_list.ListOptions.Render "header", "left"
%>
<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
	<% If DiscountTypes.SortUrl(DiscountTypes.DiscountType) = "" Then %>
		<td><%= DiscountTypes.DiscountType.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.DiscountType) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.DiscountType.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If DiscountTypes.DiscountType.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.DiscountType.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
	<% If DiscountTypes.SortUrl(DiscountTypes.DiscountTitle) = "" Then %>
		<td><%= DiscountTypes.DiscountTitle.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.DiscountTitle) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.DiscountTitle.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If DiscountTypes.DiscountTitle.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.DiscountTitle.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
	<% If DiscountTypes.SortUrl(DiscountTypes.freeShipping) = "" Then %>
		<td><%= DiscountTypes.freeShipping.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.freeShipping) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.freeShipping.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.freeShipping.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.freeShipping.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
	<% If DiscountTypes.SortUrl(DiscountTypes.FreePerQty) = "" Then %>
		<td><%= DiscountTypes.FreePerQty.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.FreePerQty) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.FreePerQty.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.FreePerQty.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.FreePerQty.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
	<% If DiscountTypes.SortUrl(DiscountTypes.SpecialPrice) = "" Then %>
		<td><%= DiscountTypes.SpecialPrice.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.SpecialPrice) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.SpecialPrice.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.SpecialPrice.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.SpecialPrice.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.fDiscountTitle.Visible Then ' fDiscountTitle %>
	<% If DiscountTypes.SortUrl(DiscountTypes.fDiscountTitle) = "" Then %>
		<td><%= DiscountTypes.fDiscountTitle.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.fDiscountTitle) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.fDiscountTitle.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If DiscountTypes.fDiscountTitle.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.fDiscountTitle.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.StartDate.Visible Then ' StartDate %>
	<% If DiscountTypes.SortUrl(DiscountTypes.StartDate) = "" Then %>
		<td><%= DiscountTypes.StartDate.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.StartDate) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.StartDate.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.StartDate.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.StartDate.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.EndDate.Visible Then ' EndDate %>
	<% If DiscountTypes.SortUrl(DiscountTypes.EndDate) = "" Then %>
		<td><%= DiscountTypes.EndDate.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.EndDate) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.EndDate.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.EndDate.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.EndDate.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.DiscountPerc.Visible Then ' DiscountPerc %>
	<% If DiscountTypes.SortUrl(DiscountTypes.DiscountPerc) = "" Then %>
		<td><%= DiscountTypes.DiscountPerc.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= DiscountTypes.SortUrl(DiscountTypes.DiscountPerc) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.DiscountPerc.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.DiscountPerc.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.DiscountPerc.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
DiscountTypes_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (DiscountTypes.ExportAll And DiscountTypes.Export <> "") Then
	DiscountTypes_list.StopRec = DiscountTypes_list.TotalRecs
Else

	' Set the last record to display
	If DiscountTypes_list.TotalRecs > DiscountTypes_list.StartRec + DiscountTypes_list.DisplayRecs - 1 Then
		DiscountTypes_list.StopRec = DiscountTypes_list.StartRec + DiscountTypes_list.DisplayRecs - 1
	Else
		DiscountTypes_list.StopRec = DiscountTypes_list.TotalRecs
	End If
End If

' Move to first record
DiscountTypes_list.RecCnt = DiscountTypes_list.StartRec - 1
If Not DiscountTypes_list.Recordset.Eof Then
	DiscountTypes_list.Recordset.MoveFirst
	If DiscountTypes_list.StartRec > 1 Then DiscountTypes_list.Recordset.Move DiscountTypes_list.StartRec - 1
ElseIf Not DiscountTypes.AllowAddDeleteRow And DiscountTypes_list.StopRec = 0 Then
	DiscountTypes_list.StopRec = DiscountTypes.GridAddRowCount
End If

' Initialize Aggregate
DiscountTypes.RowType = EW_ROWTYPE_AGGREGATEINIT
Call DiscountTypes.ResetAttrs()
Call DiscountTypes_list.RenderRow()
DiscountTypes_list.RowCnt = 0

' Output date rows
Do While CLng(DiscountTypes_list.RecCnt) < CLng(DiscountTypes_list.StopRec)
	DiscountTypes_list.RecCnt = DiscountTypes_list.RecCnt + 1
	If CLng(DiscountTypes_list.RecCnt) >= CLng(DiscountTypes_list.StartRec) Then
		DiscountTypes_list.RowCnt = DiscountTypes_list.RowCnt + 1

	' Set up key count
	DiscountTypes_list.KeyCount = DiscountTypes_list.RowIndex
	Call DiscountTypes.ResetAttrs()
	DiscountTypes.CssClass = ""
	If DiscountTypes.CurrentAction = "gridadd" Then
	Else
		Call DiscountTypes_list.LoadRowValues(DiscountTypes_list.Recordset) ' Load row values
	End If
	DiscountTypes.RowType = EW_ROWTYPE_VIEW ' Render view
	DiscountTypes.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call DiscountTypes_list.RenderRow()

	' Render list options
	Call DiscountTypes_list.RenderListOptions()
%>
	<tr<%= DiscountTypes.RowAttributes %>>
<%

' Render list options (body, left)
DiscountTypes_list.ListOptions.Render "body", "left"
%>
	<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
		<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ListViewValue %></div>
<a name="<%= DiscountTypes_list.PageObjName & "_row_" & DiscountTypes_list.RowCnt %>" id="<%= DiscountTypes_list.PageObjName & "_row_" & DiscountTypes_list.RowCnt %>"></a></td>
	<% End If %>
	<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
		<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
	<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.fDiscountTitle.Visible Then ' fDiscountTitle %>
		<td<%= DiscountTypes.fDiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.fDiscountTitle.ViewAttributes %>><%= DiscountTypes.fDiscountTitle.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.StartDate.Visible Then ' StartDate %>
		<td<%= DiscountTypes.StartDate.CellAttributes %>>
<div<%= DiscountTypes.StartDate.ViewAttributes %>><%= DiscountTypes.StartDate.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.EndDate.Visible Then ' EndDate %>
		<td<%= DiscountTypes.EndDate.CellAttributes %>>
<div<%= DiscountTypes.EndDate.ViewAttributes %>><%= DiscountTypes.EndDate.ListViewValue %></div>
</td>
	<% End If %>
	<% If DiscountTypes.DiscountPerc.Visible Then ' DiscountPerc %>
		<td<%= DiscountTypes.DiscountPerc.CellAttributes %>>
<div<%= DiscountTypes.DiscountPerc.ViewAttributes %>><%= DiscountTypes.DiscountPerc.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
DiscountTypes_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If DiscountTypes.CurrentAction <> "gridadd" Then
		DiscountTypes_list.Recordset.MoveNext()
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
DiscountTypes_list.Recordset.Close
Set DiscountTypes_list.Recordset = Nothing
%>
<% If DiscountTypes_list.TotalRecs > 0 Then %>
<% If DiscountTypes.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If DiscountTypes.CurrentAction <> "gridadd" And DiscountTypes.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(DiscountTypes_list.Pager) Then Set DiscountTypes_list.Pager = ew_NewNumericPager(DiscountTypes_list.StartRec, DiscountTypes_list.DisplayRecs, DiscountTypes_list.TotalRecs, DiscountTypes_list.RecRange) %>
<% If DiscountTypes_list.Pager.RecordCount > 0 Then %>
	<% If DiscountTypes_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In DiscountTypes_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= DiscountTypes_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If DiscountTypes_list.Pager.NextButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.LastButton.Enabled Then %>
	<a href="<%= DiscountTypes_list.PageUrl %>start=<%= DiscountTypes_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If DiscountTypes_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= DiscountTypes_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= DiscountTypes_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= DiscountTypes_list.Pager.RecordCount %>
<% Else %>
	<% If DiscountTypes_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= DiscountTypes_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% If Discountcodes.DetailAdd And Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= DiscountTypes.AddUrl & "?" & EW_TABLE_SHOW_DETAIL & "=Discountcodes" %>"><%= Language.Phrase("AddLink") %>&nbsp;<%= DiscountTypes.TableCaption %>/<%= Discountcodes.TableCaption %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If DiscountTypes.Export = "" And DiscountTypes.CurrentAction = "" Then %>
<% End If %>
<%
DiscountTypes_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If DiscountTypes.Export = "" Then %>
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
Set DiscountTypes_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDiscountTypes_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "DiscountTypes"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "DiscountTypes_list"
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
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "DiscountTypesadd.asp?" & EW_TABLE_SHOW_DETAIL & "="
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "DiscountTypesdelete.asp"
		MultiUpdateUrl = "DiscountTypesupdate.asp"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize other table object
		If IsEmpty(Discountcodes) Then Set Discountcodes = New cDiscountcodes

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "DiscountTypes"

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
				DiscountTypes.GridAddRowCount = gridaddcnt
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
		Set DiscountTypes = Nothing
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
			If DiscountTypes.Export <> "" Or DiscountTypes.CurrentAction = "gridadd" Or DiscountTypes.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call DiscountTypes.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If DiscountTypes.RecordsPerPage <> "" Then
			DisplayRecs = DiscountTypes.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call DiscountTypes.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			DiscountTypes.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				DiscountTypes.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = DiscountTypes.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		DiscountTypes.SessionWhere = sFilter
		DiscountTypes.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, DiscountTypes.DiscountType, Keyword)
			Call BuildBasicSearchSQL(sWhere, DiscountTypes.DiscountTitle, Keyword)
			Call BuildBasicSearchSQL(sWhere, DiscountTypes.fDiscountTitle, Keyword)
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
		sSearchKeyword = DiscountTypes.BasicSearchKeyword
		sSearchType = DiscountTypes.BasicSearchType
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
			DiscountTypes.SessionBasicSearchKeyword = sSearchKeyword
			DiscountTypes.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		DiscountTypes.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		DiscountTypes.SessionBasicSearchKeyword = ""
		DiscountTypes.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If DiscountTypes.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			DiscountTypes.BasicSearchKeyword = DiscountTypes.SessionBasicSearchKeyword
			DiscountTypes.BasicSearchType = DiscountTypes.SessionBasicSearchType
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
			DiscountTypes.CurrentOrder = Request.QueryString("order")
			DiscountTypes.CurrentOrderType = Request.QueryString("ordertype")

			' Field DiscountType
			Call DiscountTypes.UpdateSort(DiscountTypes.DiscountType)

			' Field DiscountTitle
			Call DiscountTypes.UpdateSort(DiscountTypes.DiscountTitle)

			' Field freeShipping
			Call DiscountTypes.UpdateSort(DiscountTypes.freeShipping)

			' Field FreePerQty
			Call DiscountTypes.UpdateSort(DiscountTypes.FreePerQty)

			' Field SpecialPrice
			Call DiscountTypes.UpdateSort(DiscountTypes.SpecialPrice)

			' Field fDiscountTitle
			Call DiscountTypes.UpdateSort(DiscountTypes.fDiscountTitle)

			' Field StartDate
			Call DiscountTypes.UpdateSort(DiscountTypes.StartDate)

			' Field EndDate
			Call DiscountTypes.UpdateSort(DiscountTypes.EndDate)

			' Field DiscountPerc
			Call DiscountTypes.UpdateSort(DiscountTypes.DiscountPerc)
			DiscountTypes.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = DiscountTypes.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If DiscountTypes.SqlOrderBy <> "" Then
				sOrderBy = DiscountTypes.SqlOrderBy
				DiscountTypes.SessionOrderBy = sOrderBy
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
				DiscountTypes.SessionOrderBy = sOrderBy
				DiscountTypes.DiscountType.Sort = ""
				DiscountTypes.DiscountTitle.Sort = ""
				DiscountTypes.freeShipping.Sort = ""
				DiscountTypes.FreePerQty.Sort = ""
				DiscountTypes.SpecialPrice.Sort = ""
				DiscountTypes.fDiscountTitle.Sort = ""
				DiscountTypes.StartDate.Sort = ""
				DiscountTypes.EndDate.Sort = ""
				DiscountTypes.DiscountPerc.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			DiscountTypes.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		ListOptions.Add("edit")
		ListOptions.GetItem("edit").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("edit").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("edit").OnLeft = True
		ListOptions.Add("delete")
		ListOptions.GetItem("delete").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("delete").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("delete").OnLeft = True
		ListOptions.Add("detail_Discountcodes")
		ListOptions.GetItem("detail_Discountcodes").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("detail_Discountcodes").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("detail_Discountcodes").OnLeft = True
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
		If Security.IsLoggedIn() And ListOptions.GetItem("delete").Visible Then
			ListOptions.GetItem("delete").Body = "<a class=""ewRowLink""" & "" & " href=""" & DeleteUrl & """>" & "<img src=""images/delete.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
		End If
		If Security.IsLoggedIn() Then
			Set item = ListOptions.GetItem("detail_Discountcodes")
			item.Body = "<img src=""images/detail.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ width=""16"" height=""16"" border=""0"">" & Language.TablePhrase("Discountcodes", "TblCaption")
			item.Body = "<a class=""ewRowLink"" href=""Discountcodeslist.asp?" & EW_TABLE_SHOW_MASTER & "=DiscountTypes&DiscountTypeId=" & Server.URLEncode(DiscountTypes.DiscountTypeId.CurrentValue&"") & """>" & item.Body & "</a>"
			links = ""
			If Discountcodes.DetailEdit And Security.IsLoggedIn() And Security.IsLoggedIn() Then
				links = links & "<a class=""ewRowLink"" href=""" & DiscountTypes.EditUrl(EW_TABLE_SHOW_DETAIL & "=Discountcodes") & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>&nbsp;"
			End If
			If links <> "" Then item.Body = item.Body & "<br>" & links
		End If
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	Function RenderListOptionsExt()
		Dim sHyperLinkParm, oListOpt, links
		sSqlWrk = "[DiscountTypeId]=" & ew_AdjustSql(DiscountTypes.DiscountTypeId.CurrentValue) & ""
		sSqlWrk = ew_Encode(TEAencrypt(sSqlWrk, EW_RANDOM_KEY))
		sSqlWrk = Replace(sSqlWrk, "'", "\'")
		Set oListOpt = ListOptions.GetItem("detail_Discountcodes")
		oListOpt.Body = "<img src=""images/detail.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ width=""16"" height=""16"" border=""0"">" & Language.TablePhrase("Discountcodes", "TblCaption")
		sHyperLinkParm = " href=""Discountcodeslist.asp?" & EW_TABLE_SHOW_MASTER & "=DiscountTypes&DiscountTypeId=" & Server.URLEncode(DiscountTypes.DiscountTypeId.CurrentValue&"") & """ name=""dl%i_DiscountTypes_Discountcodes"" id=""dl%i_DiscountTypes_Discountcodes"" onmouseover=""ew_AjaxShowDetails(this, 'Discountcodespreview.asp?f=%s')"" onmouseout=""ew_AjaxHideDetails(this);"""
		sHyperLinkParm = Replace(sHyperLinkParm,"%i",RowCnt)
		sHyperLinkParm = Replace(sHyperLinkParm,"%s",sSqlWrk)
		oListOpt.Body = "<a" & sHyperLinkParm & ">" & oListOpt.Body & "</a>"
		links = ""
		If Discountcodes.DetailEdit And Security.IsLoggedIn() And Security.IsLoggedIn() Then
			links = links & "<a href=""" & DiscountTypes.EditUrl(EW_TABLE_SHOW_DETAIL & "=Discountcodes") & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>&nbsp;"
		End If
		If links <> "" Then oListOpt.Body = oListOpt.Body & "<br>" & links
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
				DiscountTypes.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					DiscountTypes.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = DiscountTypes.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			DiscountTypes.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			DiscountTypes.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			DiscountTypes.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		DiscountTypes.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		DiscountTypes.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = DiscountTypes.CurrentFilter
		Call DiscountTypes.Recordset_Selecting(sFilter)
		DiscountTypes.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = DiscountTypes.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call DiscountTypes.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = DiscountTypes.KeyFilter

		' Call Row Selecting event
		Call DiscountTypes.Row_Selecting(sFilter)

		' Load sql based on filter
		DiscountTypes.CurrentFilter = sFilter
		sSql = DiscountTypes.SQL
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
		Call DiscountTypes.Row_Selected(RsRow)
		DiscountTypes.DiscountTypeId.DbValue = RsRow("DiscountTypeId")
		DiscountTypes.DiscountType.DbValue = RsRow("DiscountType")
		DiscountTypes.DiscountTitle.DbValue = RsRow("DiscountTitle")
		DiscountTypes.freeShipping.DbValue = ew_IIf(RsRow("freeShipping"), "1", "0")
		DiscountTypes.FreePerQty.DbValue = RsRow("FreePerQty")
		DiscountTypes.SpecialPrice.DbValue = RsRow("SpecialPrice")
		DiscountTypes.fDiscountTitle.DbValue = RsRow("fDiscountTitle")
		DiscountTypes.StartDate.DbValue = RsRow("StartDate")
		DiscountTypes.EndDate.DbValue = RsRow("EndDate")
		DiscountTypes.DiscountPerc.DbValue = RsRow("DiscountPerc")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If DiscountTypes.GetKey("DiscountTypeId")&"" <> "" Then
			DiscountTypes.DiscountTypeId.CurrentValue = DiscountTypes.GetKey("DiscountTypeId") ' DiscountTypeId
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			DiscountTypes.CurrentFilter = DiscountTypes.KeyFilter
			Dim sSql
			sSql = DiscountTypes.SQL
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
		ViewUrl = DiscountTypes.ViewUrl
		EditUrl = DiscountTypes.EditUrl("")
		InlineEditUrl = DiscountTypes.InlineEditUrl
		CopyUrl = DiscountTypes.CopyUrl("")
		InlineCopyUrl = DiscountTypes.InlineCopyUrl
		DeleteUrl = DiscountTypes.DeleteUrl

		' Call Row Rendering event
		Call DiscountTypes.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' DiscountTypeId
		' DiscountType
		' DiscountTitle
		' freeShipping
		' FreePerQty
		' SpecialPrice
		' fDiscountTitle
		' StartDate
		' EndDate
		' DiscountPerc
		' -----------
		'  View  Row
		' -----------

		If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View row

			' DiscountTypeId
			DiscountTypes.DiscountTypeId.ViewValue = DiscountTypes.DiscountTypeId.CurrentValue
			DiscountTypes.DiscountTypeId.ViewCustomAttributes = ""

			' DiscountType
			DiscountTypes.DiscountType.ViewValue = DiscountTypes.DiscountType.CurrentValue
			DiscountTypes.DiscountType.ViewCustomAttributes = ""

			' DiscountTitle
			DiscountTypes.DiscountTitle.ViewValue = DiscountTypes.DiscountTitle.CurrentValue
			DiscountTypes.DiscountTitle.ViewCustomAttributes = ""

			' freeShipping
			If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then
				DiscountTypes.freeShipping.ViewValue = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(1) <> "", DiscountTypes.freeShipping.FldTagCaption(1), "Yes")
			Else
				DiscountTypes.freeShipping.ViewValue = ew_IIf(DiscountTypes.freeShipping.FldTagCaption(2) <> "", DiscountTypes.freeShipping.FldTagCaption(2), "No")
			End If
			DiscountTypes.freeShipping.ViewCustomAttributes = ""

			' FreePerQty
			DiscountTypes.FreePerQty.ViewValue = DiscountTypes.FreePerQty.CurrentValue
			DiscountTypes.FreePerQty.ViewCustomAttributes = ""

			' SpecialPrice
			DiscountTypes.SpecialPrice.ViewValue = DiscountTypes.SpecialPrice.CurrentValue
			DiscountTypes.SpecialPrice.ViewCustomAttributes = ""

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.ViewValue = DiscountTypes.fDiscountTitle.CurrentValue
			DiscountTypes.fDiscountTitle.ViewCustomAttributes = ""

			' StartDate
			DiscountTypes.StartDate.ViewValue = DiscountTypes.StartDate.CurrentValue
			DiscountTypes.StartDate.ViewCustomAttributes = ""

			' EndDate
			DiscountTypes.EndDate.ViewValue = DiscountTypes.EndDate.CurrentValue
			DiscountTypes.EndDate.ViewCustomAttributes = ""

			' DiscountPerc
			DiscountTypes.DiscountPerc.ViewValue = DiscountTypes.DiscountPerc.CurrentValue
			DiscountTypes.DiscountPerc.ViewCustomAttributes = ""

			' View refer script
			' DiscountType

			DiscountTypes.DiscountType.LinkCustomAttributes = ""
			DiscountTypes.DiscountType.HrefValue = ""
			DiscountTypes.DiscountType.TooltipValue = ""

			' DiscountTitle
			DiscountTypes.DiscountTitle.LinkCustomAttributes = ""
			DiscountTypes.DiscountTitle.HrefValue = ""
			DiscountTypes.DiscountTitle.TooltipValue = ""

			' freeShipping
			DiscountTypes.freeShipping.LinkCustomAttributes = ""
			DiscountTypes.freeShipping.HrefValue = ""
			DiscountTypes.freeShipping.TooltipValue = ""

			' FreePerQty
			DiscountTypes.FreePerQty.LinkCustomAttributes = ""
			DiscountTypes.FreePerQty.HrefValue = ""
			DiscountTypes.FreePerQty.TooltipValue = ""

			' SpecialPrice
			DiscountTypes.SpecialPrice.LinkCustomAttributes = ""
			DiscountTypes.SpecialPrice.HrefValue = ""
			DiscountTypes.SpecialPrice.TooltipValue = ""

			' fDiscountTitle
			DiscountTypes.fDiscountTitle.LinkCustomAttributes = ""
			DiscountTypes.fDiscountTitle.HrefValue = ""
			DiscountTypes.fDiscountTitle.TooltipValue = ""

			' StartDate
			DiscountTypes.StartDate.LinkCustomAttributes = ""
			DiscountTypes.StartDate.HrefValue = ""
			DiscountTypes.StartDate.TooltipValue = ""

			' EndDate
			DiscountTypes.EndDate.LinkCustomAttributes = ""
			DiscountTypes.EndDate.HrefValue = ""
			DiscountTypes.EndDate.TooltipValue = ""

			' DiscountPerc
			DiscountTypes.DiscountPerc.LinkCustomAttributes = ""
			DiscountTypes.DiscountPerc.HrefValue = ""
			DiscountTypes.DiscountPerc.TooltipValue = ""
		End If

		' Call Row Rendered event
		If DiscountTypes.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call DiscountTypes.Row_Rendered()
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
