<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Ordersinfo.asp"-->
<!--#include file="OrderDetailsinfo.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Orders_list
Set Orders_list = New cOrders_list
Set Page = Orders_list

' Page init processing
Call Orders_list.Page_Init()

' Page main processing
Call Orders_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Orders.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Orders_list = new ew_Page("Orders_list");
// page properties
Orders_list.PageID = "list"; // page ID
Orders_list.FormID = "fOrderslist"; // form ID
var EW_PAGE_ID = Orders_list.PageID; // for backward compatibility
// extend page with validate function for search
Orders_list.ValidateSearch = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (this.ValidateRequired) {
		var infix = "";
		elm = fobj.elements["x" + infix + "_OrderId"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.OrderId.FldErrMsg) %>");
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
Orders_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Orders_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Orders_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Orders_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Orders.Export = "") Or (EW_EXPORT_MASTER_RECORD And Orders.Export = "print") Then %>
<%
gsMasterReturnUrl = "Customerslist.asp"
If Orders_list.DbMasterFilter <> "" And Orders.CurrentMasterTable = "Customers" Then
	If Orders_list.MasterRecordExists Then
		If Orders.CurrentMasterTable = Orders.TableVar Then gsMasterReturnUrl = gsMasterReturnUrl & "?" & EW_TABLE_SHOW_MASTER & "="
%>
<p class="aspmaker ewTitle"><%= Language.Phrase("MasterRecord") %><%= Customers.TableCaption %>
&nbsp;&nbsp;<% Orders_list.ExportOptions.Render "body", "" %>
</p>
<% If Orders.Export = "" Then %>
<p class="aspmaker"><a href="<%= gsMasterReturnUrl %>"><%= Language.Phrase("BackToMasterPage") %></a></p>
<% End If %>
<!--#include file="Customersmaster.asp"-->
<%
	End If
End If
%>
<% End If %>
<% Orders_list.ShowPageHeader() %>
<%

' Load recordset
Set Orders_list.Recordset = Orders_list.LoadRecordset()
	Orders_list.TotalRecs = Orders_list.Recordset.RecordCount
	Orders_list.StartRec = 1
	If Orders_list.DisplayRecs <= 0 Then ' Display all records
		Orders_list.DisplayRecs = Orders_list.TotalRecs
	End If
	If Not (Orders.ExportAll And Orders.Export <> "") Then
		Orders_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Orders.TableCaption %>
<% If Orders.CurrentMasterTable = "" Then %>
&nbsp;&nbsp;<% Orders_list.ExportOptions.Render "body", "" %>
<% End If %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Orders.Export = "" And Orders.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Orders_list);" style="text-decoration: none;"><img id="Orders_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Orders_list_SearchPanel">
<form name="fOrderslistsrch" id="fOrderslistsrch" class="ewForm" action="<%= ew_CurrentPage %>" onsubmit="return Orders_list.ValidateSearch(this);">
<input type="hidden" id="t" name="t" value="Orders">
<div class="ewBasicSearch">
<%
If gsSearchError = "" Then
	Call Orders_list.LoadAdvancedSearch() ' Load advanced search
End If

' Render for search
Orders.RowType = EW_ROWTYPE_SEARCH

' Render row
Call Orders.ResetAttrs()
Call Orders_list.RenderRow()
%>
<div id="xsr_1" class="ewCssTableRow">
	<span id="xsc_OrderId" class="ewCssTableCell">
		<span class="ewSearchCaption"><%= Orders.OrderId.FldCaption %></span>
		<span class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_OrderId" id="z_OrderId" value="="></span>
		<span class="ewSearchField">
<input type="text" name="x_OrderId" id="x_OrderId" value="<%= Orders.OrderId.EditValue %>"<%= Orders.OrderId.EditAttributes %>>
</span>
	</span>
</div>
<div id="xsr_2" class="ewCssTableRow">
	<span id="xsc_CustomerId" class="ewCssTableCell">
		<span class="ewSearchCaption"><%= Orders.CustomerId.FldCaption %></span>
		<span class="ewSearchOprCell"><%= Language.Phrase("=") %><input type="hidden" name="z_CustomerId" id="z_CustomerId" value="="></span>
		<span class="ewSearchField">
<% If Orders.CustomerId.SessionValue <> "" Then %>
<div<%= Orders.CustomerId.ViewAttributes %>>
<% If Orders.CustomerId.LinkAttributes <> "" Then %>
<a<%= Orders.CustomerId.LinkAttributes %>><%= Orders.CustomerId.ListViewValue %></a>
<% Else %>
<%= Orders.CustomerId.ListViewValue %>
<% End If %>
</div>
<input type="hidden" id="x_CustomerId" name="x_CustomerId" value="<%= ew_HtmlEncode(Orders.CustomerId.AdvancedSearch.SearchValue) %>">
<% Else %>
<select id="x_CustomerId" name="x_CustomerId"<%= Orders.CustomerId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Orders.CustomerId.EditValue) Then
	arwrk = Orders.CustomerId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Orders.CustomerId.AdvancedSearch.SearchValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
<% If arwrk(2, rowcntwrk) <> "" Then %>
<%= ew_ValueSeparator(rowcntwrk,1,Orders.CustomerId) %><%= arwrk(2, rowcntwrk) %>
<% End If %>
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
<div id="xsr_3" class="ewCssTableRow">
	<span id="xsc_payment_status" class="ewCssTableCell">
		<span class="ewSearchCaption"><%= Orders.payment_status.FldCaption %></span>
		<span class="ewSearchOprCell"><%= Language.Phrase("LIKE") %><input type="hidden" name="z_payment_status" id="z_payment_status" value="LIKE"></span>
		<span class="ewSearchField">
<select id="x_payment_status" name="x_payment_status"<%= Orders.payment_status.EditAttributes %>>
<%
emptywrk = True
If IsArray(Orders.payment_status.EditValue) Then
	arwrk = Orders.payment_status.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Orders.payment_status.AdvancedSearch.SearchValue&"" Then
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
</span>
	</span>
</div>
<div id="xsr_4" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Orders.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Orders_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_5" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Orders.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Orders.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Orders.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Orders_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Orders.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Orders.CurrentAction <> "gridadd" And Orders.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Orders_list.Pager) Then Set Orders_list.Pager = ew_NewNumericPager(Orders_list.StartRec, Orders_list.DisplayRecs, Orders_list.TotalRecs, Orders_list.RecRange) %>
<% If Orders_list.Pager.RecordCount > 0 Then %>
	<% If Orders_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Orders_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Orders_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Orders_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Orders_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Orders_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Orders_list.Pager.RecordCount %>
<% Else %>
	<% If Orders_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Orders_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% If OrderDetails.DetailAdd And Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= Orders.AddUrl & "?" & EW_TABLE_SHOW_DETAIL & "=OrderDetails" %>"><%= Language.Phrase("AddLink") %>&nbsp;<%= Orders.TableCaption %>/<%= OrderDetails.TableCaption %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
<% If Orders_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fOrderslist, '<%= Orders_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fOrderslist" id="fOrderslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Orders">
<div id="gmp_Orders" class="ewGridMiddlePanel">
<% If Orders_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Orders.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Orders_list.RenderListOptions()

' Render list options (header, left)
Orders_list.ListOptions.Render "header", "left"
%>
<% If Orders.OrderId.Visible Then ' OrderId %>
	<% If Orders.SortUrl(Orders.OrderId) = "" Then %>
		<td><%= Orders.OrderId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.OrderId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.OrderId.FldCaption %></td><td style="width: 10px;"><% If Orders.OrderId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.OrderId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.CustomerId.Visible Then ' CustomerId %>
	<% If Orders.SortUrl(Orders.CustomerId) = "" Then %>
		<td><%= Orders.CustomerId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.CustomerId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.CustomerId.FldCaption %></td><td style="width: 10px;"><% If Orders.CustomerId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.CustomerId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Amount.Visible Then ' Amount %>
	<% If Orders.SortUrl(Orders.Amount) = "" Then %>
		<td><%= Orders.Amount.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Amount) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Amount.FldCaption %></td><td style="width: 10px;"><% If Orders.Amount.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Amount.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Ship_FirstName.Visible Then ' Ship_FirstName %>
	<% If Orders.SortUrl(Orders.Ship_FirstName) = "" Then %>
		<td><%= Orders.Ship_FirstName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Ship_FirstName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Ship_FirstName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Orders.Ship_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Ship_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Ship_LastName.Visible Then ' Ship_LastName %>
	<% If Orders.SortUrl(Orders.Ship_LastName) = "" Then %>
		<td><%= Orders.Ship_LastName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Ship_LastName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Ship_LastName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Orders.Ship_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Ship_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.payment_status.Visible Then ' payment_status %>
	<% If Orders.SortUrl(Orders.payment_status) = "" Then %>
		<td><%= Orders.payment_status.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.payment_status) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.payment_status.FldCaption %></td><td style="width: 10px;"><% If Orders.payment_status.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.payment_status.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Ordered_Date.Visible Then ' Ordered_Date %>
	<% If Orders.SortUrl(Orders.Ordered_Date) = "" Then %>
		<td><%= Orders.Ordered_Date.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Ordered_Date) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Ordered_Date.FldCaption %></td><td style="width: 10px;"><% If Orders.Ordered_Date.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Ordered_Date.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.payer_email.Visible Then ' payer_email %>
	<% If Orders.SortUrl(Orders.payer_email) = "" Then %>
		<td><%= Orders.payer_email.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.payer_email) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.payer_email.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Orders.payer_email.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.payer_email.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.payment_gross.Visible Then ' payment_gross %>
	<% If Orders.SortUrl(Orders.payment_gross) = "" Then %>
		<td><%= Orders.payment_gross.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.payment_gross) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.payment_gross.FldCaption %></td><td style="width: 10px;"><% If Orders.payment_gross.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.payment_gross.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.payment_fee.Visible Then ' payment_fee %>
	<% If Orders.SortUrl(Orders.payment_fee) = "" Then %>
		<td><%= Orders.payment_fee.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.payment_fee) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.payment_fee.FldCaption %></td><td style="width: 10px;"><% If Orders.payment_fee.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.payment_fee.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Tax.Visible Then ' Tax %>
	<% If Orders.SortUrl(Orders.Tax) = "" Then %>
		<td><%= Orders.Tax.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Tax) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Tax.FldCaption %></td><td style="width: 10px;"><% If Orders.Tax.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Tax.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.Shipping.Visible Then ' Shipping %>
	<% If Orders.SortUrl(Orders.Shipping) = "" Then %>
		<td><%= Orders.Shipping.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.Shipping) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.Shipping.FldCaption %></td><td style="width: 10px;"><% If Orders.Shipping.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.Shipping.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.EmailSent.Visible Then ' EmailSent %>
	<% If Orders.SortUrl(Orders.EmailSent) = "" Then %>
		<td><%= Orders.EmailSent.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.EmailSent) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.EmailSent.FldCaption %></td><td style="width: 10px;"><% If Orders.EmailSent.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.EmailSent.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.EmailDate.Visible Then ' EmailDate %>
	<% If Orders.SortUrl(Orders.EmailDate) = "" Then %>
		<td><%= Orders.EmailDate.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.EmailDate) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.EmailDate.FldCaption %></td><td style="width: 10px;"><% If Orders.EmailDate.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.EmailDate.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Orders.PromoCodeUsed.Visible Then ' PromoCodeUsed %>
	<% If Orders.SortUrl(Orders.PromoCodeUsed) = "" Then %>
		<td><%= Orders.PromoCodeUsed.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Orders.SortUrl(Orders.PromoCodeUsed) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Orders.PromoCodeUsed.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Orders.PromoCodeUsed.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Orders.PromoCodeUsed.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Orders_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Orders.ExportAll And Orders.Export <> "") Then
	Orders_list.StopRec = Orders_list.TotalRecs
Else

	' Set the last record to display
	If Orders_list.TotalRecs > Orders_list.StartRec + Orders_list.DisplayRecs - 1 Then
		Orders_list.StopRec = Orders_list.StartRec + Orders_list.DisplayRecs - 1
	Else
		Orders_list.StopRec = Orders_list.TotalRecs
	End If
End If

' Move to first record
Orders_list.RecCnt = Orders_list.StartRec - 1
If Not Orders_list.Recordset.Eof Then
	Orders_list.Recordset.MoveFirst
	If Orders_list.StartRec > 1 Then Orders_list.Recordset.Move Orders_list.StartRec - 1
ElseIf Not Orders.AllowAddDeleteRow And Orders_list.StopRec = 0 Then
	Orders_list.StopRec = Orders.GridAddRowCount
End If

' Initialize Aggregate
Orders.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Orders.ResetAttrs()
Call Orders_list.RenderRow()
Orders_list.RowCnt = 0

' Output date rows
Do While CLng(Orders_list.RecCnt) < CLng(Orders_list.StopRec)
	Orders_list.RecCnt = Orders_list.RecCnt + 1
	If CLng(Orders_list.RecCnt) >= CLng(Orders_list.StartRec) Then
		Orders_list.RowCnt = Orders_list.RowCnt + 1

	' Set up key count
	Orders_list.KeyCount = Orders_list.RowIndex
	Call Orders.ResetAttrs()
	Orders.CssClass = ""
	If Orders.CurrentAction = "gridadd" Then
	Else
		Call Orders_list.LoadRowValues(Orders_list.Recordset) ' Load row values
	End If
	Orders.RowType = EW_ROWTYPE_VIEW ' Render view
	Orders.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Orders_list.RenderRow()

	' Render list options
	Call Orders_list.RenderListOptions()
%>
	<tr<%= Orders.RowAttributes %>>
<%

' Render list options (body, left)
Orders_list.ListOptions.Render "body", "left"
%>
	<% If Orders.OrderId.Visible Then ' OrderId %>
		<td<%= Orders.OrderId.CellAttributes %>>
<div<%= Orders.OrderId.ViewAttributes %>><%= Orders.OrderId.ListViewValue %></div>
<a name="<%= Orders_list.PageObjName & "_row_" & Orders_list.RowCnt %>" id="<%= Orders_list.PageObjName & "_row_" & Orders_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Orders.CustomerId.Visible Then ' CustomerId %>
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
		<td<%= Orders.Amount.CellAttributes %>>
<div<%= Orders.Amount.ViewAttributes %>><%= Orders.Amount.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.Ship_FirstName.Visible Then ' Ship_FirstName %>
		<td<%= Orders.Ship_FirstName.CellAttributes %>>
<div<%= Orders.Ship_FirstName.ViewAttributes %>><%= Orders.Ship_FirstName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.Ship_LastName.Visible Then ' Ship_LastName %>
		<td<%= Orders.Ship_LastName.CellAttributes %>>
<div<%= Orders.Ship_LastName.ViewAttributes %>><%= Orders.Ship_LastName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.payment_status.Visible Then ' payment_status %>
		<td<%= Orders.payment_status.CellAttributes %>>
<div<%= Orders.payment_status.ViewAttributes %>><%= Orders.payment_status.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.Ordered_Date.Visible Then ' Ordered_Date %>
		<td<%= Orders.Ordered_Date.CellAttributes %>>
<div<%= Orders.Ordered_Date.ViewAttributes %>><%= Orders.Ordered_Date.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.payer_email.Visible Then ' payer_email %>
		<td<%= Orders.payer_email.CellAttributes %>>
<div<%= Orders.payer_email.ViewAttributes %>><%= Orders.payer_email.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.payment_gross.Visible Then ' payment_gross %>
		<td<%= Orders.payment_gross.CellAttributes %>>
<div<%= Orders.payment_gross.ViewAttributes %>><%= Orders.payment_gross.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.payment_fee.Visible Then ' payment_fee %>
		<td<%= Orders.payment_fee.CellAttributes %>>
<div<%= Orders.payment_fee.ViewAttributes %>><%= Orders.payment_fee.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.Tax.Visible Then ' Tax %>
		<td<%= Orders.Tax.CellAttributes %>>
<div<%= Orders.Tax.ViewAttributes %>><%= Orders.Tax.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.Shipping.Visible Then ' Shipping %>
		<td<%= Orders.Shipping.CellAttributes %>>
<div<%= Orders.Shipping.ViewAttributes %>><%= Orders.Shipping.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.EmailSent.Visible Then ' EmailSent %>
		<td<%= Orders.EmailSent.CellAttributes %>>
<div<%= Orders.EmailSent.ViewAttributes %>><%= Orders.EmailSent.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.EmailDate.Visible Then ' EmailDate %>
		<td<%= Orders.EmailDate.CellAttributes %>>
<div<%= Orders.EmailDate.ViewAttributes %>><%= Orders.EmailDate.ListViewValue %></div>
</td>
	<% End If %>
	<% If Orders.PromoCodeUsed.Visible Then ' PromoCodeUsed %>
		<td<%= Orders.PromoCodeUsed.CellAttributes %>>
<div<%= Orders.PromoCodeUsed.ViewAttributes %>><%= Orders.PromoCodeUsed.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Orders_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Orders.CurrentAction <> "gridadd" Then
		Orders_list.Recordset.MoveNext()
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
Orders_list.Recordset.Close
Set Orders_list.Recordset = Nothing
%>
<% If Orders_list.TotalRecs > 0 Then %>
<% If Orders.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Orders.CurrentAction <> "gridadd" And Orders.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Orders_list.Pager) Then Set Orders_list.Pager = ew_NewNumericPager(Orders_list.StartRec, Orders_list.DisplayRecs, Orders_list.TotalRecs, Orders_list.RecRange) %>
<% If Orders_list.Pager.RecordCount > 0 Then %>
	<% If Orders_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Orders_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Orders_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Orders_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Orders_list.PageUrl %>start=<%= Orders_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Orders_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Orders_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Orders_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Orders_list.Pager.RecordCount %>
<% Else %>
	<% If Orders_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Orders_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% If OrderDetails.DetailAdd And Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="<%= Orders.AddUrl & "?" & EW_TABLE_SHOW_DETAIL & "=OrderDetails" %>"><%= Language.Phrase("AddLink") %>&nbsp;<%= Orders.TableCaption %>/<%= OrderDetails.TableCaption %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
<% If Orders_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fOrderslist, '<%= Orders_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Orders.Export = "" And Orders.CurrentAction = "" Then %>
<% End If %>
<%
Orders_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Orders.Export = "" Then %>
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
Set Orders_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrders_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Orders"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Orders_list"
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
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Ordersadd.asp?" & EW_TABLE_SHOW_DETAIL & "="
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Ordersdelete.asp"
		MultiUpdateUrl = "Ordersupdate.asp"

		' Initialize other table object
		If IsEmpty(OrderDetails) Then Set OrderDetails = New cOrderDetails

		' Initialize other table object
		If IsEmpty(Customers) Then Set Customers = New cCustomers

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Orders"

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
				Orders.GridAddRowCount = gridaddcnt
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
		Set Orders = Nothing
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
			If Orders.Export <> "" Or Orders.CurrentAction = "gridadd" Or Orders.CurrentAction = "gridedit" Then
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
			Call Orders.Recordset_SearchValidated()

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
		If Orders.RecordsPerPage <> "" Then
			DisplayRecs = Orders.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Orders.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			If sSrchAdvanced = "" Then Call ResetAdvancedSearchParms()
			Orders.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Orders.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Orders.SearchWhere
		End If
		sFilter = ""

		' Restore master/detail filter
		DbMasterFilter = Orders.MasterFilter ' Restore master filter
		DbDetailFilter = Orders.DetailFilter ' Restore detail filter
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)
		Dim RsMaster

		' Load master record
		If Orders.MasterFilter <> "" And Orders.CurrentMasterTable = "Customers" Then
			Set RsMaster = Customers.LoadRs(DbMasterFilter)
			MasterRecordExists = Not (RsMaster Is Nothing)
			If Not MasterRecordExists Then
				FailureMessage = Language.Phrase("NoRecord") ' Set no record found
				Call Page_Terminate(Orders.ReturnUrl) ' Return to caller
			Else
				Call Customers.LoadListRowValues(RsMaster)
				Customers.RowType = EW_ROWTYPE_MASTER ' Master row
				Call Customers.RenderListRow()
				RsMaster.Close
				Set RsMaster = Nothing
			End If
		End If

		' Set up filter in Session
		Orders.SessionWhere = sFilter
		Orders.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Advanced Search Where based on QueryString parameters
	'
	Function AdvancedSearchWhere()
		Dim sWhere
		sWhere = ""

		' Field OrderId
		Call BuildSearchSql(sWhere, Orders.OrderId, False)

		' Field CustomerId
		Call BuildSearchSql(sWhere, Orders.CustomerId, False)

		' Field InvoiceId
		Call BuildSearchSql(sWhere, Orders.InvoiceId, False)

		' Field Amount
		Call BuildSearchSql(sWhere, Orders.Amount, False)

		' Field Ship_FirstName
		Call BuildSearchSql(sWhere, Orders.Ship_FirstName, False)

		' Field Ship_LastName
		Call BuildSearchSql(sWhere, Orders.Ship_LastName, False)

		' Field Ship_Address
		Call BuildSearchSql(sWhere, Orders.Ship_Address, False)

		' Field Ship_Address2
		Call BuildSearchSql(sWhere, Orders.Ship_Address2, False)

		' Field Ship_City
		Call BuildSearchSql(sWhere, Orders.Ship_City, False)

		' Field Ship_Province
		Call BuildSearchSql(sWhere, Orders.Ship_Province, False)

		' Field Ship_Postal
		Call BuildSearchSql(sWhere, Orders.Ship_Postal, False)

		' Field Ship_Country
		Call BuildSearchSql(sWhere, Orders.Ship_Country, False)

		' Field Ship_Phone
		Call BuildSearchSql(sWhere, Orders.Ship_Phone, False)

		' Field Ship_Email
		Call BuildSearchSql(sWhere, Orders.Ship_Email, False)

		' Field payment_status
		Call BuildSearchSql(sWhere, Orders.payment_status, False)

		' Field Ordered_Date
		Call BuildSearchSql(sWhere, Orders.Ordered_Date, False)

		' Field payment_date
		Call BuildSearchSql(sWhere, Orders.payment_date, False)

		' Field pfirst_name
		Call BuildSearchSql(sWhere, Orders.pfirst_name, False)

		' Field plast_name
		Call BuildSearchSql(sWhere, Orders.plast_name, False)

		' Field payer_email
		Call BuildSearchSql(sWhere, Orders.payer_email, False)

		' Field txn_id
		Call BuildSearchSql(sWhere, Orders.txn_id, False)

		' Field payment_gross
		Call BuildSearchSql(sWhere, Orders.payment_gross, False)

		' Field payment_fee
		Call BuildSearchSql(sWhere, Orders.payment_fee, False)

		' Field payment_type
		Call BuildSearchSql(sWhere, Orders.payment_type, False)

		' Field txn_type
		Call BuildSearchSql(sWhere, Orders.txn_type, False)

		' Field receiver_email
		Call BuildSearchSql(sWhere, Orders.receiver_email, False)

		' Field pShip_Name
		Call BuildSearchSql(sWhere, Orders.pShip_Name, False)

		' Field pShip_Address
		Call BuildSearchSql(sWhere, Orders.pShip_Address, False)

		' Field pShip_City
		Call BuildSearchSql(sWhere, Orders.pShip_City, False)

		' Field pShip_Province
		Call BuildSearchSql(sWhere, Orders.pShip_Province, False)

		' Field pShip_Postal
		Call BuildSearchSql(sWhere, Orders.pShip_Postal, False)

		' Field pShip_Country
		Call BuildSearchSql(sWhere, Orders.pShip_Country, False)

		' Field Tax
		Call BuildSearchSql(sWhere, Orders.Tax, False)

		' Field Shipping
		Call BuildSearchSql(sWhere, Orders.Shipping, False)

		' Field EmailSent
		Call BuildSearchSql(sWhere, Orders.EmailSent, False)

		' Field EmailDate
		Call BuildSearchSql(sWhere, Orders.EmailDate, False)

		' Field PromoCodeUsed
		Call BuildSearchSql(sWhere, Orders.PromoCodeUsed, False)
		AdvancedSearchWhere = sWhere

		' Set up search parm
		If AdvancedSearchWhere <> "" Then

			' Field OrderId
			Call SetSearchParm(Orders.OrderId)

			' Field CustomerId
			Call SetSearchParm(Orders.CustomerId)

			' Field InvoiceId
			Call SetSearchParm(Orders.InvoiceId)

			' Field Amount
			Call SetSearchParm(Orders.Amount)

			' Field Ship_FirstName
			Call SetSearchParm(Orders.Ship_FirstName)

			' Field Ship_LastName
			Call SetSearchParm(Orders.Ship_LastName)

			' Field Ship_Address
			Call SetSearchParm(Orders.Ship_Address)

			' Field Ship_Address2
			Call SetSearchParm(Orders.Ship_Address2)

			' Field Ship_City
			Call SetSearchParm(Orders.Ship_City)

			' Field Ship_Province
			Call SetSearchParm(Orders.Ship_Province)

			' Field Ship_Postal
			Call SetSearchParm(Orders.Ship_Postal)

			' Field Ship_Country
			Call SetSearchParm(Orders.Ship_Country)

			' Field Ship_Phone
			Call SetSearchParm(Orders.Ship_Phone)

			' Field Ship_Email
			Call SetSearchParm(Orders.Ship_Email)

			' Field payment_status
			Call SetSearchParm(Orders.payment_status)

			' Field Ordered_Date
			Call SetSearchParm(Orders.Ordered_Date)

			' Field payment_date
			Call SetSearchParm(Orders.payment_date)

			' Field pfirst_name
			Call SetSearchParm(Orders.pfirst_name)

			' Field plast_name
			Call SetSearchParm(Orders.plast_name)

			' Field payer_email
			Call SetSearchParm(Orders.payer_email)

			' Field txn_id
			Call SetSearchParm(Orders.txn_id)

			' Field payment_gross
			Call SetSearchParm(Orders.payment_gross)

			' Field payment_fee
			Call SetSearchParm(Orders.payment_fee)

			' Field payment_type
			Call SetSearchParm(Orders.payment_type)

			' Field txn_type
			Call SetSearchParm(Orders.txn_type)

			' Field receiver_email
			Call SetSearchParm(Orders.receiver_email)

			' Field pShip_Name
			Call SetSearchParm(Orders.pShip_Name)

			' Field pShip_Address
			Call SetSearchParm(Orders.pShip_Address)

			' Field pShip_City
			Call SetSearchParm(Orders.pShip_City)

			' Field pShip_Province
			Call SetSearchParm(Orders.pShip_Province)

			' Field pShip_Postal
			Call SetSearchParm(Orders.pShip_Postal)

			' Field pShip_Country
			Call SetSearchParm(Orders.pShip_Country)

			' Field Tax
			Call SetSearchParm(Orders.Tax)

			' Field Shipping
			Call SetSearchParm(Orders.Shipping)

			' Field EmailSent
			Call SetSearchParm(Orders.EmailSent)

			' Field EmailDate
			Call SetSearchParm(Orders.EmailDate)

			' Field PromoCodeUsed
			Call SetSearchParm(Orders.PromoCodeUsed)
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
		Call Orders.SetAdvancedSearch("x_" & FldParm, Fld.AdvancedSearch.SearchValue)
		Call Orders.SetAdvancedSearch("z_" & FldParm, Fld.AdvancedSearch.SearchOperator)
		Call Orders.SetAdvancedSearch("v_" & FldParm, Fld.AdvancedSearch.SearchCondition)
		Call Orders.SetAdvancedSearch("y_" & FldParm, Fld.AdvancedSearch.SearchValue2)
		Call Orders.SetAdvancedSearch("w_" & FldParm, Fld.AdvancedSearch.SearchOperator2)
	End Sub

	' -----------------------------------------------------------------
	' Get search parm
	'
	Sub GetSearchParm(Fld)
		Dim FldParm
		FldParm = Mid(Fld.FldVar, 3)
		Fld.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_" & FldParm)
		Fld.AdvancedSearch.SearchOperator = Orders.GetAdvancedSearch("z_" & FldParm)
		Fld.AdvancedSearch.SearchCondition = Orders.GetAdvancedSearch("v_" & FldParm)
		Fld.AdvancedSearch.SearchValue2 = Orders.GetAdvancedSearch("y_" & FldParm)
		Fld.AdvancedSearch.SearchOperator2 = Orders.GetAdvancedSearch("w_" & FldParm)
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
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_FirstName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_LastName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Address2, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Postal, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Phone, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.Ship_Email, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.payment_status, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pfirst_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.plast_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.payer_email, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.txn_id, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.payment_type, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.txn_type, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.receiver_email, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_Name, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_Postal, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.pShip_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.EmailSent, Keyword)
			Call BuildBasicSearchSQL(sWhere, Orders.PromoCodeUsed, Keyword)
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
		sSearchKeyword = Orders.BasicSearchKeyword
		sSearchType = Orders.BasicSearchType
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
			Orders.SessionBasicSearchKeyword = sSearchKeyword
			Orders.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Orders.SearchWhere = SearchWhere

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
		Orders.SessionBasicSearchKeyword = ""
		Orders.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Clear all advanced search parameters
	'
	Sub ResetAdvancedSearchParms()

		' Clear advanced search parameters
		Call Orders.SetAdvancedSearch("x_OrderId", "")
		Call Orders.SetAdvancedSearch("x_CustomerId", "")
		Call Orders.SetAdvancedSearch("x_InvoiceId", "")
		Call Orders.SetAdvancedSearch("x_Amount", "")
		Call Orders.SetAdvancedSearch("x_Ship_FirstName", "")
		Call Orders.SetAdvancedSearch("x_Ship_LastName", "")
		Call Orders.SetAdvancedSearch("x_Ship_Address", "")
		Call Orders.SetAdvancedSearch("x_Ship_Address2", "")
		Call Orders.SetAdvancedSearch("x_Ship_City", "")
		Call Orders.SetAdvancedSearch("x_Ship_Province", "")
		Call Orders.SetAdvancedSearch("x_Ship_Postal", "")
		Call Orders.SetAdvancedSearch("x_Ship_Country", "")
		Call Orders.SetAdvancedSearch("x_Ship_Phone", "")
		Call Orders.SetAdvancedSearch("x_Ship_Email", "")
		Call Orders.SetAdvancedSearch("x_payment_status", "")
		Call Orders.SetAdvancedSearch("x_Ordered_Date", "")
		Call Orders.SetAdvancedSearch("x_payment_date", "")
		Call Orders.SetAdvancedSearch("x_pfirst_name", "")
		Call Orders.SetAdvancedSearch("x_plast_name", "")
		Call Orders.SetAdvancedSearch("x_payer_email", "")
		Call Orders.SetAdvancedSearch("x_txn_id", "")
		Call Orders.SetAdvancedSearch("x_payment_gross", "")
		Call Orders.SetAdvancedSearch("x_payment_fee", "")
		Call Orders.SetAdvancedSearch("x_payment_type", "")
		Call Orders.SetAdvancedSearch("x_txn_type", "")
		Call Orders.SetAdvancedSearch("x_receiver_email", "")
		Call Orders.SetAdvancedSearch("x_pShip_Name", "")
		Call Orders.SetAdvancedSearch("x_pShip_Address", "")
		Call Orders.SetAdvancedSearch("x_pShip_City", "")
		Call Orders.SetAdvancedSearch("x_pShip_Province", "")
		Call Orders.SetAdvancedSearch("x_pShip_Postal", "")
		Call Orders.SetAdvancedSearch("x_pShip_Country", "")
		Call Orders.SetAdvancedSearch("x_Tax", "")
		Call Orders.SetAdvancedSearch("x_Shipping", "")
		Call Orders.SetAdvancedSearch("x_EmailSent", "")
		Call Orders.SetAdvancedSearch("x_EmailDate", "")
		Call Orders.SetAdvancedSearch("x_PromoCodeUsed", "")
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Orders.BasicSearchKeyword & "" <> "" Then bRestore = False
		If Orders.OrderId.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.CustomerId.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.InvoiceId.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Amount.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_FirstName.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_LastName.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Address.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Address2.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_City.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Province.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Postal.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Country.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Phone.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ship_Email.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payment_status.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Ordered_Date.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payment_date.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pfirst_name.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.plast_name.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payer_email.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.txn_id.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payment_gross.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payment_fee.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.payment_type.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.txn_type.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.receiver_email.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_Name.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_Address.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_City.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_Province.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_Postal.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.pShip_Country.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Tax.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.Shipping.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.EmailSent.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.EmailDate.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		If Orders.PromoCodeUsed.AdvancedSearch.SearchValue & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Orders.BasicSearchKeyword = Orders.SessionBasicSearchKeyword
			Orders.BasicSearchType = Orders.SessionBasicSearchType

			' Restore advanced search values
			Call GetSearchParm(Orders.OrderId)
			Call GetSearchParm(Orders.CustomerId)
			Call GetSearchParm(Orders.InvoiceId)
			Call GetSearchParm(Orders.Amount)
			Call GetSearchParm(Orders.Ship_FirstName)
			Call GetSearchParm(Orders.Ship_LastName)
			Call GetSearchParm(Orders.Ship_Address)
			Call GetSearchParm(Orders.Ship_Address2)
			Call GetSearchParm(Orders.Ship_City)
			Call GetSearchParm(Orders.Ship_Province)
			Call GetSearchParm(Orders.Ship_Postal)
			Call GetSearchParm(Orders.Ship_Country)
			Call GetSearchParm(Orders.Ship_Phone)
			Call GetSearchParm(Orders.Ship_Email)
			Call GetSearchParm(Orders.payment_status)
			Call GetSearchParm(Orders.Ordered_Date)
			Call GetSearchParm(Orders.payment_date)
			Call GetSearchParm(Orders.pfirst_name)
			Call GetSearchParm(Orders.plast_name)
			Call GetSearchParm(Orders.payer_email)
			Call GetSearchParm(Orders.txn_id)
			Call GetSearchParm(Orders.payment_gross)
			Call GetSearchParm(Orders.payment_fee)
			Call GetSearchParm(Orders.payment_type)
			Call GetSearchParm(Orders.txn_type)
			Call GetSearchParm(Orders.receiver_email)
			Call GetSearchParm(Orders.pShip_Name)
			Call GetSearchParm(Orders.pShip_Address)
			Call GetSearchParm(Orders.pShip_City)
			Call GetSearchParm(Orders.pShip_Province)
			Call GetSearchParm(Orders.pShip_Postal)
			Call GetSearchParm(Orders.pShip_Country)
			Call GetSearchParm(Orders.Tax)
			Call GetSearchParm(Orders.Shipping)
			Call GetSearchParm(Orders.EmailSent)
			Call GetSearchParm(Orders.EmailDate)
			Call GetSearchParm(Orders.PromoCodeUsed)
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
			Orders.CurrentOrder = Request.QueryString("order")
			Orders.CurrentOrderType = Request.QueryString("ordertype")

			' Field OrderId
			Call Orders.UpdateSort(Orders.OrderId)

			' Field CustomerId
			Call Orders.UpdateSort(Orders.CustomerId)

			' Field Amount
			Call Orders.UpdateSort(Orders.Amount)

			' Field Ship_FirstName
			Call Orders.UpdateSort(Orders.Ship_FirstName)

			' Field Ship_LastName
			Call Orders.UpdateSort(Orders.Ship_LastName)

			' Field payment_status
			Call Orders.UpdateSort(Orders.payment_status)

			' Field Ordered_Date
			Call Orders.UpdateSort(Orders.Ordered_Date)

			' Field payer_email
			Call Orders.UpdateSort(Orders.payer_email)

			' Field payment_gross
			Call Orders.UpdateSort(Orders.payment_gross)

			' Field payment_fee
			Call Orders.UpdateSort(Orders.payment_fee)

			' Field Tax
			Call Orders.UpdateSort(Orders.Tax)

			' Field Shipping
			Call Orders.UpdateSort(Orders.Shipping)

			' Field EmailSent
			Call Orders.UpdateSort(Orders.EmailSent)

			' Field EmailDate
			Call Orders.UpdateSort(Orders.EmailDate)

			' Field PromoCodeUsed
			Call Orders.UpdateSort(Orders.PromoCodeUsed)
			Orders.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Orders.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Orders.SqlOrderBy <> "" Then
				sOrderBy = Orders.SqlOrderBy
				Orders.SessionOrderBy = sOrderBy
				Orders.OrderId.Sort = "DESC"
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
				Orders.CurrentMasterTable = "" ' Clear master table
				DbMasterFilter = ""
				DbDetailFilter = ""
				Orders.CustomerId.SessionValue = ""
			End If

			' Reset Sort Criteria
			If LCase(sCmd) = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				Orders.SessionOrderBy = sOrderBy
				Orders.OrderId.Sort = ""
				Orders.CustomerId.Sort = ""
				Orders.Amount.Sort = ""
				Orders.Ship_FirstName.Sort = ""
				Orders.Ship_LastName.Sort = ""
				Orders.payment_status.Sort = ""
				Orders.Ordered_Date.Sort = ""
				Orders.payer_email.Sort = ""
				Orders.payment_gross.Sort = ""
				Orders.payment_fee.Sort = ""
				Orders.Tax.Sort = ""
				Orders.Shipping.Sort = ""
				Orders.EmailSent.Sort = ""
				Orders.EmailDate.Sort = ""
				Orders.PromoCodeUsed.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Orders.StartRecordNumber = StartRec
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
		ListOptions.Add("detail_OrderDetails")
		ListOptions.GetItem("detail_OrderDetails").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("detail_OrderDetails").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("detail_OrderDetails").OnLeft = True
		ListOptions.Add("checkbox")
		ListOptions.GetItem("checkbox").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("checkbox").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("checkbox").OnLeft = True
		ListOptions.MoveItem "checkbox", 0 ' Move to first column
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Orders_list.SelectAllKey(this);"">"
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
		If Security.IsLoggedIn() Then
			Set item = ListOptions.GetItem("detail_OrderDetails")
			item.Body = "<img src=""images/detail.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ width=""16"" height=""16"" border=""0"">" & Language.TablePhrase("OrderDetails", "TblCaption")
			item.Body = "<a class=""ewRowLink"" href=""OrderDetailslist.asp?" & EW_TABLE_SHOW_MASTER & "=Orders&OrderId=" & Server.URLEncode(Orders.OrderId.CurrentValue&"") & """>" & item.Body & "</a>"
			links = ""
			If OrderDetails.DetailEdit And Security.IsLoggedIn() And Security.IsLoggedIn() Then
				links = links & "<a class=""ewRowLink"" href=""" & Orders.EditUrl(EW_TABLE_SHOW_DETAIL & "=OrderDetails") & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>&nbsp;"
			End If
			If links <> "" Then item.Body = item.Body & "<br>" & links
		End If
		If Security.IsLoggedIn() And ListOptions.GetItem("checkbox").Visible Then
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Orders.OrderId.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
		End If
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	Function RenderListOptionsExt()
		Dim sHyperLinkParm, oListOpt, links
		sSqlWrk = "[OrderId]=" & ew_AdjustSql(Orders.OrderId.CurrentValue) & ""
		sSqlWrk = ew_Encode(TEAencrypt(sSqlWrk, EW_RANDOM_KEY))
		sSqlWrk = Replace(sSqlWrk, "'", "\'")
		Set oListOpt = ListOptions.GetItem("detail_OrderDetails")
		oListOpt.Body = "<img src=""images/detail.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DetailLink")) & """ width=""16"" height=""16"" border=""0"">" & Language.TablePhrase("OrderDetails", "TblCaption")
		sHyperLinkParm = " href=""OrderDetailslist.asp?" & EW_TABLE_SHOW_MASTER & "=Orders&OrderId=" & Server.URLEncode(Orders.OrderId.CurrentValue&"") & """ name=""dl%i_Orders_OrderDetails"" id=""dl%i_Orders_OrderDetails"" onmouseover=""ew_AjaxShowDetails(this, 'OrderDetailspreview.asp?f=%s')"" onmouseout=""ew_AjaxHideDetails(this);"""
		sHyperLinkParm = Replace(sHyperLinkParm,"%i",RowCnt)
		sHyperLinkParm = Replace(sHyperLinkParm,"%s",sSqlWrk)
		oListOpt.Body = "<a" & sHyperLinkParm & ">" & oListOpt.Body & "</a>"
		links = ""
		If OrderDetails.DetailEdit And Security.IsLoggedIn() And Security.IsLoggedIn() Then
			links = links & "<a href=""" & Orders.EditUrl(EW_TABLE_SHOW_DETAIL & "=OrderDetails") & """>" & "<img src=""images/edit.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("EditLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>&nbsp;"
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
				Orders.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Orders.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Orders.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Orders.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Orders.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Orders.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Orders.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Orders.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	'  Load search values for validation
	'
	Function LoadSearchValues()

		' Load search values
		Orders.OrderId.AdvancedSearch.SearchValue = Request.QueryString("x_OrderId")
		Orders.OrderId.AdvancedSearch.SearchOperator = Request.QueryString("z_OrderId")
		Orders.CustomerId.AdvancedSearch.SearchValue = Request.QueryString("x_CustomerId")
		Orders.CustomerId.AdvancedSearch.SearchOperator = Request.QueryString("z_CustomerId")
		Orders.InvoiceId.AdvancedSearch.SearchValue = Request.QueryString("x_InvoiceId")
		Orders.InvoiceId.AdvancedSearch.SearchOperator = Request.QueryString("z_InvoiceId")
		Orders.Amount.AdvancedSearch.SearchValue = Request.QueryString("x_Amount")
		Orders.Amount.AdvancedSearch.SearchOperator = Request.QueryString("z_Amount")
		Orders.Ship_FirstName.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_FirstName")
		Orders.Ship_FirstName.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_FirstName")
		Orders.Ship_LastName.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_LastName")
		Orders.Ship_LastName.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_LastName")
		Orders.Ship_Address.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Address")
		Orders.Ship_Address.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Address")
		Orders.Ship_Address2.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Address2")
		Orders.Ship_Address2.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Address2")
		Orders.Ship_City.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_City")
		Orders.Ship_City.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_City")
		Orders.Ship_Province.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Province")
		Orders.Ship_Province.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Province")
		Orders.Ship_Postal.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Postal")
		Orders.Ship_Postal.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Postal")
		Orders.Ship_Country.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Country")
		Orders.Ship_Country.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Country")
		Orders.Ship_Phone.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Phone")
		Orders.Ship_Phone.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Phone")
		Orders.Ship_Email.AdvancedSearch.SearchValue = Request.QueryString("x_Ship_Email")
		Orders.Ship_Email.AdvancedSearch.SearchOperator = Request.QueryString("z_Ship_Email")
		Orders.payment_status.AdvancedSearch.SearchValue = Request.QueryString("x_payment_status")
		Orders.payment_status.AdvancedSearch.SearchOperator = Request.QueryString("z_payment_status")
		Orders.Ordered_Date.AdvancedSearch.SearchValue = Request.QueryString("x_Ordered_Date")
		Orders.Ordered_Date.AdvancedSearch.SearchOperator = Request.QueryString("z_Ordered_Date")
		Orders.payment_date.AdvancedSearch.SearchValue = Request.QueryString("x_payment_date")
		Orders.payment_date.AdvancedSearch.SearchOperator = Request.QueryString("z_payment_date")
		Orders.pfirst_name.AdvancedSearch.SearchValue = Request.QueryString("x_pfirst_name")
		Orders.pfirst_name.AdvancedSearch.SearchOperator = Request.QueryString("z_pfirst_name")
		Orders.plast_name.AdvancedSearch.SearchValue = Request.QueryString("x_plast_name")
		Orders.plast_name.AdvancedSearch.SearchOperator = Request.QueryString("z_plast_name")
		Orders.payer_email.AdvancedSearch.SearchValue = Request.QueryString("x_payer_email")
		Orders.payer_email.AdvancedSearch.SearchOperator = Request.QueryString("z_payer_email")
		Orders.txn_id.AdvancedSearch.SearchValue = Request.QueryString("x_txn_id")
		Orders.txn_id.AdvancedSearch.SearchOperator = Request.QueryString("z_txn_id")
		Orders.payment_gross.AdvancedSearch.SearchValue = Request.QueryString("x_payment_gross")
		Orders.payment_gross.AdvancedSearch.SearchOperator = Request.QueryString("z_payment_gross")
		Orders.payment_fee.AdvancedSearch.SearchValue = Request.QueryString("x_payment_fee")
		Orders.payment_fee.AdvancedSearch.SearchOperator = Request.QueryString("z_payment_fee")
		Orders.payment_type.AdvancedSearch.SearchValue = Request.QueryString("x_payment_type")
		Orders.payment_type.AdvancedSearch.SearchOperator = Request.QueryString("z_payment_type")
		Orders.txn_type.AdvancedSearch.SearchValue = Request.QueryString("x_txn_type")
		Orders.txn_type.AdvancedSearch.SearchOperator = Request.QueryString("z_txn_type")
		Orders.receiver_email.AdvancedSearch.SearchValue = Request.QueryString("x_receiver_email")
		Orders.receiver_email.AdvancedSearch.SearchOperator = Request.QueryString("z_receiver_email")
		Orders.pShip_Name.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_Name")
		Orders.pShip_Name.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_Name")
		Orders.pShip_Address.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_Address")
		Orders.pShip_Address.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_Address")
		Orders.pShip_City.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_City")
		Orders.pShip_City.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_City")
		Orders.pShip_Province.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_Province")
		Orders.pShip_Province.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_Province")
		Orders.pShip_Postal.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_Postal")
		Orders.pShip_Postal.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_Postal")
		Orders.pShip_Country.AdvancedSearch.SearchValue = Request.QueryString("x_pShip_Country")
		Orders.pShip_Country.AdvancedSearch.SearchOperator = Request.QueryString("z_pShip_Country")
		Orders.Tax.AdvancedSearch.SearchValue = Request.QueryString("x_Tax")
		Orders.Tax.AdvancedSearch.SearchOperator = Request.QueryString("z_Tax")
		Orders.Shipping.AdvancedSearch.SearchValue = Request.QueryString("x_Shipping")
		Orders.Shipping.AdvancedSearch.SearchOperator = Request.QueryString("z_Shipping")
		Orders.EmailSent.AdvancedSearch.SearchValue = Request.QueryString("x_EmailSent")
		Orders.EmailSent.AdvancedSearch.SearchOperator = Request.QueryString("z_EmailSent")
		Orders.EmailDate.AdvancedSearch.SearchValue = Request.QueryString("x_EmailDate")
		Orders.EmailDate.AdvancedSearch.SearchOperator = Request.QueryString("z_EmailDate")
		Orders.PromoCodeUsed.AdvancedSearch.SearchValue = Request.QueryString("x_PromoCodeUsed")
		Orders.PromoCodeUsed.AdvancedSearch.SearchOperator = Request.QueryString("z_PromoCodeUsed")
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Orders.CurrentFilter
		Call Orders.Recordset_Selecting(sFilter)
		Orders.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Orders.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Orders.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Orders.KeyFilter

		' Call Row Selecting event
		Call Orders.Row_Selecting(sFilter)

		' Load sql based on filter
		Orders.CurrentFilter = sFilter
		sSql = Orders.SQL
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
		Call Orders.Row_Selected(RsRow)
		Orders.OrderId.DbValue = RsRow("OrderId")
		Orders.CustomerId.DbValue = RsRow("CustomerId")
		Orders.InvoiceId.DbValue = RsRow("InvoiceId")
		Orders.Amount.DbValue = RsRow("Amount")
		Orders.Ship_FirstName.DbValue = RsRow("Ship_FirstName")
		Orders.Ship_LastName.DbValue = RsRow("Ship_LastName")
		Orders.Ship_Address.DbValue = RsRow("Ship_Address")
		Orders.Ship_Address2.DbValue = RsRow("Ship_Address2")
		Orders.Ship_City.DbValue = RsRow("Ship_City")
		Orders.Ship_Province.DbValue = RsRow("Ship_Province")
		Orders.Ship_Postal.DbValue = RsRow("Ship_Postal")
		Orders.Ship_Country.DbValue = RsRow("Ship_Country")
		Orders.Ship_Phone.DbValue = RsRow("Ship_Phone")
		Orders.Ship_Email.DbValue = RsRow("Ship_Email")
		Orders.payment_status.DbValue = RsRow("payment_status")
		Orders.Ordered_Date.DbValue = RsRow("Ordered_Date")
		Orders.payment_date.DbValue = RsRow("payment_date")
		Orders.pfirst_name.DbValue = RsRow("pfirst_name")
		Orders.plast_name.DbValue = RsRow("plast_name")
		Orders.payer_email.DbValue = RsRow("payer_email")
		Orders.txn_id.DbValue = RsRow("txn_id")
		Orders.payment_gross.DbValue = RsRow("payment_gross")
		Orders.payment_fee.DbValue = RsRow("payment_fee")
		Orders.payment_type.DbValue = RsRow("payment_type")
		Orders.txn_type.DbValue = RsRow("txn_type")
		Orders.receiver_email.DbValue = RsRow("receiver_email")
		Orders.pShip_Name.DbValue = RsRow("pShip_Name")
		Orders.pShip_Address.DbValue = RsRow("pShip_Address")
		Orders.pShip_City.DbValue = RsRow("pShip_City")
		Orders.pShip_Province.DbValue = RsRow("pShip_Province")
		Orders.pShip_Postal.DbValue = RsRow("pShip_Postal")
		Orders.pShip_Country.DbValue = RsRow("pShip_Country")
		Orders.Tax.DbValue = RsRow("Tax")
		Orders.Shipping.DbValue = RsRow("Shipping")
		Orders.EmailSent.DbValue = RsRow("EmailSent")
		Orders.EmailDate.DbValue = RsRow("EmailDate")
		Orders.PromoCodeUsed.DbValue = RsRow("PromoCodeUsed")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Orders.GetKey("OrderId")&"" <> "" Then
			Orders.OrderId.CurrentValue = Orders.GetKey("OrderId") ' OrderId
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Orders.CurrentFilter = Orders.KeyFilter
			Dim sSql
			sSql = Orders.SQL
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
		ViewUrl = Orders.ViewUrl
		EditUrl = Orders.EditUrl("")
		InlineEditUrl = Orders.InlineEditUrl
		CopyUrl = Orders.CopyUrl("")
		InlineCopyUrl = Orders.InlineCopyUrl
		DeleteUrl = Orders.DeleteUrl

		' Call Row Rendering event
		Call Orders.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' OrderId
		' CustomerId
		' InvoiceId
		' Amount
		' Ship_FirstName
		' Ship_LastName
		' Ship_Address
		' Ship_Address2
		' Ship_City
		' Ship_Province
		' Ship_Postal
		' Ship_Country
		' Ship_Phone
		' Ship_Email
		' payment_status
		' Ordered_Date
		' payment_date
		' pfirst_name
		' plast_name
		' payer_email
		' txn_id
		' payment_gross
		' payment_fee
		' payment_type
		' txn_type
		' receiver_email
		' pShip_Name
		' pShip_Address
		' pShip_City
		' pShip_Province
		' pShip_Postal
		' pShip_Country
		' Tax
		' Shipping
		' EmailSent
		' EmailDate
		' PromoCodeUsed
		' -----------
		'  View  Row
		' -----------

		If Orders.RowType = EW_ROWTYPE_VIEW Then ' View row

			' OrderId
			Orders.OrderId.ViewValue = Orders.OrderId.CurrentValue
			Orders.OrderId.ViewCustomAttributes = ""

			' CustomerId
			If Orders.CustomerId.CurrentValue & "" <> "" Then
				sFilterWrk = "[CustomerID] = " & ew_AdjustSql(Orders.CustomerId.CurrentValue) & ""
			sSqlWrk = "SELECT DISTINCT [Inv_FirstName], [Inv_LastName] FROM [Customers]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Orders.CustomerId.ViewValue = RsWrk("Inv_FirstName")
					Orders.CustomerId.ViewValue = Orders.CustomerId.ViewValue & ew_ValueSeparator(0,1,Orders.CustomerId) & RsWrk("Inv_LastName")
				Else
					Orders.CustomerId.ViewValue = Orders.CustomerId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Orders.CustomerId.ViewValue = Null
			End If
			Orders.CustomerId.ViewCustomAttributes = ""

			' InvoiceId
			Orders.InvoiceId.ViewValue = Orders.InvoiceId.CurrentValue
			Orders.InvoiceId.ViewCustomAttributes = ""

			' Amount
			Orders.Amount.ViewValue = Orders.Amount.CurrentValue
			Orders.Amount.ViewCustomAttributes = ""

			' Ship_FirstName
			Orders.Ship_FirstName.ViewValue = Orders.Ship_FirstName.CurrentValue
			Orders.Ship_FirstName.ViewCustomAttributes = ""

			' Ship_LastName
			Orders.Ship_LastName.ViewValue = Orders.Ship_LastName.CurrentValue
			Orders.Ship_LastName.ViewCustomAttributes = ""

			' Ship_Address
			Orders.Ship_Address.ViewValue = Orders.Ship_Address.CurrentValue
			Orders.Ship_Address.ViewCustomAttributes = ""

			' Ship_Address2
			Orders.Ship_Address2.ViewValue = Orders.Ship_Address2.CurrentValue
			Orders.Ship_Address2.ViewCustomAttributes = ""

			' Ship_City
			Orders.Ship_City.ViewValue = Orders.Ship_City.CurrentValue
			Orders.Ship_City.ViewCustomAttributes = ""

			' Ship_Province
			Orders.Ship_Province.ViewValue = Orders.Ship_Province.CurrentValue
			Orders.Ship_Province.ViewCustomAttributes = ""

			' Ship_Postal
			Orders.Ship_Postal.ViewValue = Orders.Ship_Postal.CurrentValue
			Orders.Ship_Postal.ViewCustomAttributes = ""

			' Ship_Country
			Orders.Ship_Country.ViewValue = Orders.Ship_Country.CurrentValue
			Orders.Ship_Country.ViewCustomAttributes = ""

			' Ship_Phone
			Orders.Ship_Phone.ViewValue = Orders.Ship_Phone.CurrentValue
			Orders.Ship_Phone.ViewCustomAttributes = ""

			' Ship_Email
			Orders.Ship_Email.ViewValue = Orders.Ship_Email.CurrentValue
			Orders.Ship_Email.ViewCustomAttributes = ""

			' payment_status
			If Not IsNull(Orders.payment_status.CurrentValue) Then
				Select Case Orders.payment_status.CurrentValue
					Case "Completed"
						Orders.payment_status.ViewValue = ew_IIf(Orders.payment_status.FldTagCaption(1) <> "", Orders.payment_status.FldTagCaption(1), "Completed")
					Case "WIP"
						Orders.payment_status.ViewValue = ew_IIf(Orders.payment_status.FldTagCaption(2) <> "", Orders.payment_status.FldTagCaption(2), "WIP")
					Case "Pending"
						Orders.payment_status.ViewValue = ew_IIf(Orders.payment_status.FldTagCaption(3) <> "", Orders.payment_status.FldTagCaption(3), "Pending")
					Case "Failed"
						Orders.payment_status.ViewValue = ew_IIf(Orders.payment_status.FldTagCaption(4) <> "", Orders.payment_status.FldTagCaption(4), "Failed")
					Case "Cancelled"
						Orders.payment_status.ViewValue = ew_IIf(Orders.payment_status.FldTagCaption(5) <> "", Orders.payment_status.FldTagCaption(5), "Cancelled")
					Case Else
						Orders.payment_status.ViewValue = Orders.payment_status.CurrentValue
				End Select
			Else
				Orders.payment_status.ViewValue = Null
			End If
			Orders.payment_status.ViewCustomAttributes = ""

			' Ordered_Date
			Orders.Ordered_Date.ViewValue = Orders.Ordered_Date.CurrentValue
			Orders.Ordered_Date.ViewCustomAttributes = ""

			' payment_date
			Orders.payment_date.ViewValue = Orders.payment_date.CurrentValue
			Orders.payment_date.ViewCustomAttributes = ""

			' pfirst_name
			Orders.pfirst_name.ViewValue = Orders.pfirst_name.CurrentValue
			Orders.pfirst_name.ViewCustomAttributes = ""

			' plast_name
			Orders.plast_name.ViewValue = Orders.plast_name.CurrentValue
			Orders.plast_name.ViewCustomAttributes = ""

			' payer_email
			Orders.payer_email.ViewValue = Orders.payer_email.CurrentValue
			Orders.payer_email.ViewCustomAttributes = ""

			' txn_id
			Orders.txn_id.ViewValue = Orders.txn_id.CurrentValue
			Orders.txn_id.ViewCustomAttributes = ""

			' payment_gross
			Orders.payment_gross.ViewValue = Orders.payment_gross.CurrentValue
			Orders.payment_gross.ViewCustomAttributes = ""

			' payment_fee
			Orders.payment_fee.ViewValue = Orders.payment_fee.CurrentValue
			Orders.payment_fee.ViewCustomAttributes = ""

			' payment_type
			Orders.payment_type.ViewValue = Orders.payment_type.CurrentValue
			Orders.payment_type.ViewCustomAttributes = ""

			' txn_type
			Orders.txn_type.ViewValue = Orders.txn_type.CurrentValue
			Orders.txn_type.ViewCustomAttributes = ""

			' receiver_email
			Orders.receiver_email.ViewValue = Orders.receiver_email.CurrentValue
			Orders.receiver_email.ViewCustomAttributes = ""

			' pShip_Name
			Orders.pShip_Name.ViewValue = Orders.pShip_Name.CurrentValue
			Orders.pShip_Name.ViewCustomAttributes = ""

			' pShip_Address
			Orders.pShip_Address.ViewValue = Orders.pShip_Address.CurrentValue
			Orders.pShip_Address.ViewCustomAttributes = ""

			' pShip_City
			Orders.pShip_City.ViewValue = Orders.pShip_City.CurrentValue
			Orders.pShip_City.ViewCustomAttributes = ""

			' pShip_Province
			Orders.pShip_Province.ViewValue = Orders.pShip_Province.CurrentValue
			Orders.pShip_Province.ViewCustomAttributes = ""

			' pShip_Postal
			Orders.pShip_Postal.ViewValue = Orders.pShip_Postal.CurrentValue
			Orders.pShip_Postal.ViewCustomAttributes = ""

			' pShip_Country
			Orders.pShip_Country.ViewValue = Orders.pShip_Country.CurrentValue
			Orders.pShip_Country.ViewCustomAttributes = ""

			' Tax
			Orders.Tax.ViewValue = Orders.Tax.CurrentValue
			Orders.Tax.ViewCustomAttributes = ""

			' Shipping
			Orders.Shipping.ViewValue = Orders.Shipping.CurrentValue
			Orders.Shipping.ViewCustomAttributes = ""

			' EmailSent
			If Not IsNull(Orders.EmailSent.CurrentValue) Then
				Select Case Orders.EmailSent.CurrentValue
					Case "confirm"
						Orders.EmailSent.ViewValue = ew_IIf(Orders.EmailSent.FldTagCaption(1) <> "", Orders.EmailSent.FldTagCaption(1), "Confirm")
					Case Else
						Orders.EmailSent.ViewValue = Orders.EmailSent.CurrentValue
				End Select
			Else
				Orders.EmailSent.ViewValue = Null
			End If
			Orders.EmailSent.ViewCustomAttributes = ""

			' EmailDate
			Orders.EmailDate.ViewValue = Orders.EmailDate.CurrentValue
			Orders.EmailDate.ViewCustomAttributes = ""

			' PromoCodeUsed
			Orders.PromoCodeUsed.ViewValue = Orders.PromoCodeUsed.CurrentValue
			Orders.PromoCodeUsed.ViewCustomAttributes = ""

			' View refer script
			' OrderId

			Orders.OrderId.LinkCustomAttributes = ""
			Orders.OrderId.HrefValue = ""
			Orders.OrderId.TooltipValue = ""

			' CustomerId
			Orders.CustomerId.LinkCustomAttributes = ""
			If Not ew_Empty(Orders.CustomerId.CurrentValue) Then
				Orders.CustomerId.HrefValue = "Customersedit.asp?CustomerID=" & Orders.CustomerId.CurrentValue
				Orders.CustomerId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Orders.Export <> "" Then Orders.CustomerId.HrefValue = ew_ConvertFullUrl(Orders.CustomerId.HrefValue)
			Else
				Orders.CustomerId.HrefValue = ""
			End If
			Orders.CustomerId.TooltipValue = ""

			' Amount
			Orders.Amount.LinkCustomAttributes = ""
			Orders.Amount.HrefValue = ""
			Orders.Amount.TooltipValue = ""

			' Ship_FirstName
			Orders.Ship_FirstName.LinkCustomAttributes = ""
			Orders.Ship_FirstName.HrefValue = ""
			Orders.Ship_FirstName.TooltipValue = ""

			' Ship_LastName
			Orders.Ship_LastName.LinkCustomAttributes = ""
			Orders.Ship_LastName.HrefValue = ""
			Orders.Ship_LastName.TooltipValue = ""

			' payment_status
			Orders.payment_status.LinkCustomAttributes = ""
			Orders.payment_status.HrefValue = ""
			Orders.payment_status.TooltipValue = ""

			' Ordered_Date
			Orders.Ordered_Date.LinkCustomAttributes = ""
			Orders.Ordered_Date.HrefValue = ""
			Orders.Ordered_Date.TooltipValue = ""

			' payer_email
			Orders.payer_email.LinkCustomAttributes = ""
			Orders.payer_email.HrefValue = ""
			Orders.payer_email.TooltipValue = ""

			' payment_gross
			Orders.payment_gross.LinkCustomAttributes = ""
			Orders.payment_gross.HrefValue = ""
			Orders.payment_gross.TooltipValue = ""

			' payment_fee
			Orders.payment_fee.LinkCustomAttributes = ""
			Orders.payment_fee.HrefValue = ""
			Orders.payment_fee.TooltipValue = ""

			' Tax
			Orders.Tax.LinkCustomAttributes = ""
			Orders.Tax.HrefValue = ""
			Orders.Tax.TooltipValue = ""

			' Shipping
			Orders.Shipping.LinkCustomAttributes = ""
			Orders.Shipping.HrefValue = ""
			Orders.Shipping.TooltipValue = ""

			' EmailSent
			Orders.EmailSent.LinkCustomAttributes = ""
			Orders.EmailSent.HrefValue = ""
			Orders.EmailSent.TooltipValue = ""

			' EmailDate
			Orders.EmailDate.LinkCustomAttributes = ""
			Orders.EmailDate.HrefValue = ""
			Orders.EmailDate.TooltipValue = ""

			' PromoCodeUsed
			Orders.PromoCodeUsed.LinkCustomAttributes = ""
			Orders.PromoCodeUsed.HrefValue = ""
			Orders.PromoCodeUsed.TooltipValue = ""

		' ------------
		'  Search Row
		' ------------

		ElseIf Orders.RowType = EW_ROWTYPE_SEARCH Then ' Search row

			' OrderId
			Orders.OrderId.EditCustomAttributes = ""
			Orders.OrderId.EditValue = ew_HtmlEncode(Orders.OrderId.AdvancedSearch.SearchValue)

			' CustomerId
			Orders.CustomerId.EditCustomAttributes = ""
			If Orders.CustomerId.SessionValue <> "" Then
				Orders.CustomerId.AdvancedSearch.SearchValue = Orders.CustomerId.SessionValue
			If Orders.CustomerId.CurrentValue & "" <> "" Then
				sFilterWrk = "[CustomerID] = " & ew_AdjustSql(Orders.CustomerId.CurrentValue) & ""
			sSqlWrk = "SELECT DISTINCT [Inv_FirstName], [Inv_LastName] FROM [Customers]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Orders.CustomerId.ViewValue = RsWrk("Inv_FirstName")
					Orders.CustomerId.ViewValue = Orders.CustomerId.ViewValue & ew_ValueSeparator(0,1,Orders.CustomerId) & RsWrk("Inv_LastName")
				Else
					Orders.CustomerId.ViewValue = Orders.CustomerId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Orders.CustomerId.ViewValue = Null
			End If
			Orders.CustomerId.ViewCustomAttributes = ""
			Else
				sFilterWrk = ""
			sSqlWrk = "SELECT DISTINCT [CustomerID], [Inv_FirstName] AS [DispFld], [Inv_LastName] AS [Disp2Fld], '' AS [Disp3Fld], '' AS [Disp4Fld], '' AS [SelectFilterFld] FROM [Customers]"
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
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect"), ""))
			Orders.CustomerId.EditValue = arwrk
			End If

			' Amount
			Orders.Amount.EditCustomAttributes = ""
			Orders.Amount.EditValue = ew_HtmlEncode(Orders.Amount.AdvancedSearch.SearchValue)

			' Ship_FirstName
			Orders.Ship_FirstName.EditCustomAttributes = ""
			Orders.Ship_FirstName.EditValue = ew_HtmlEncode(Orders.Ship_FirstName.AdvancedSearch.SearchValue)

			' Ship_LastName
			Orders.Ship_LastName.EditCustomAttributes = ""
			Orders.Ship_LastName.EditValue = ew_HtmlEncode(Orders.Ship_LastName.AdvancedSearch.SearchValue)

			' payment_status
			Orders.payment_status.EditCustomAttributes = ""
			Redim arwrk(1, 4)
			arwrk(0, 0) = "Completed"
			arwrk(1, 0) = ew_IIf(Orders.payment_status.FldTagCaption(1) <> "", Orders.payment_status.FldTagCaption(1), "Completed")
			arwrk(0, 1) = "WIP"
			arwrk(1, 1) = ew_IIf(Orders.payment_status.FldTagCaption(2) <> "", Orders.payment_status.FldTagCaption(2), "WIP")
			arwrk(0, 2) = "Pending"
			arwrk(1, 2) = ew_IIf(Orders.payment_status.FldTagCaption(3) <> "", Orders.payment_status.FldTagCaption(3), "Pending")
			arwrk(0, 3) = "Failed"
			arwrk(1, 3) = ew_IIf(Orders.payment_status.FldTagCaption(4) <> "", Orders.payment_status.FldTagCaption(4), "Failed")
			arwrk(0, 4) = "Cancelled"
			arwrk(1, 4) = ew_IIf(Orders.payment_status.FldTagCaption(5) <> "", Orders.payment_status.FldTagCaption(5), "Cancelled")
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect")))
			Orders.payment_status.EditValue = arwrk

			' Ordered_Date
			Orders.Ordered_Date.EditCustomAttributes = ""
			Orders.Ordered_Date.EditValue = Orders.Ordered_Date.AdvancedSearch.SearchValue

			' payer_email
			Orders.payer_email.EditCustomAttributes = ""
			Orders.payer_email.EditValue = ew_HtmlEncode(Orders.payer_email.AdvancedSearch.SearchValue)

			' payment_gross
			Orders.payment_gross.EditCustomAttributes = ""
			Orders.payment_gross.EditValue = ew_HtmlEncode(Orders.payment_gross.AdvancedSearch.SearchValue)

			' payment_fee
			Orders.payment_fee.EditCustomAttributes = ""
			Orders.payment_fee.EditValue = ew_HtmlEncode(Orders.payment_fee.AdvancedSearch.SearchValue)

			' Tax
			Orders.Tax.EditCustomAttributes = ""
			Orders.Tax.EditValue = ew_HtmlEncode(Orders.Tax.AdvancedSearch.SearchValue)

			' Shipping
			Orders.Shipping.EditCustomAttributes = ""
			Orders.Shipping.EditValue = ew_HtmlEncode(Orders.Shipping.AdvancedSearch.SearchValue)

			' EmailSent
			Orders.EmailSent.EditCustomAttributes = ""
			Redim arwrk(1, 0)
			arwrk(0, 0) = "confirm"
			arwrk(1, 0) = ew_IIf(Orders.EmailSent.FldTagCaption(1) <> "", Orders.EmailSent.FldTagCaption(1), "Confirm")
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect")))
			Orders.EmailSent.EditValue = arwrk

			' EmailDate
			Orders.EmailDate.EditCustomAttributes = ""
			Orders.EmailDate.EditValue = Orders.EmailDate.AdvancedSearch.SearchValue

			' PromoCodeUsed
			Orders.PromoCodeUsed.EditCustomAttributes = ""
			Orders.PromoCodeUsed.EditValue = ew_HtmlEncode(Orders.PromoCodeUsed.AdvancedSearch.SearchValue)
		End If
		If Orders.RowType = EW_ROWTYPE_ADD Or Orders.RowType = EW_ROWTYPE_EDIT Or Orders.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Orders.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Orders.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Orders.Row_Rendered()
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
		If Not ew_CheckInteger(Orders.OrderId.AdvancedSearch.SearchValue) Then
			Call ew_AddMessage(gsSearchError, Orders.OrderId.FldErrMsg)
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
		Orders.OrderId.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_OrderId")
		Orders.CustomerId.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_CustomerId")
		Orders.InvoiceId.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_InvoiceId")
		Orders.Amount.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Amount")
		Orders.Ship_FirstName.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_FirstName")
		Orders.Ship_LastName.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_LastName")
		Orders.Ship_Address.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Address")
		Orders.Ship_Address2.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Address2")
		Orders.Ship_City.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_City")
		Orders.Ship_Province.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Province")
		Orders.Ship_Postal.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Postal")
		Orders.Ship_Country.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Country")
		Orders.Ship_Phone.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Phone")
		Orders.Ship_Email.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ship_Email")
		Orders.payment_status.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payment_status")
		Orders.Ordered_Date.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Ordered_Date")
		Orders.payment_date.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payment_date")
		Orders.pfirst_name.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pfirst_name")
		Orders.plast_name.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_plast_name")
		Orders.payer_email.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payer_email")
		Orders.txn_id.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_txn_id")
		Orders.payment_gross.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payment_gross")
		Orders.payment_fee.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payment_fee")
		Orders.payment_type.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_payment_type")
		Orders.txn_type.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_txn_type")
		Orders.receiver_email.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_receiver_email")
		Orders.pShip_Name.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_Name")
		Orders.pShip_Address.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_Address")
		Orders.pShip_City.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_City")
		Orders.pShip_Province.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_Province")
		Orders.pShip_Postal.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_Postal")
		Orders.pShip_Country.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_pShip_Country")
		Orders.Tax.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Tax")
		Orders.Shipping.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_Shipping")
		Orders.EmailSent.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_EmailSent")
		Orders.EmailDate.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_EmailDate")
		Orders.PromoCodeUsed.AdvancedSearch.SearchValue = Orders.GetAdvancedSearch("x_PromoCodeUsed")
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
			If sMasterTblVar = "Customers" Then
				bValidMaster = True
				If Request.QueryString("CustomerID").Count > 0 Then
					Customers.CustomerID.QueryStringValue = Request.QueryString("CustomerID")
					Orders.CustomerId.QueryStringValue = Customers.CustomerID.QueryStringValue
					Orders.CustomerId.SessionValue = Orders.CustomerId.QueryStringValue
					If Not IsNumeric(Customers.CustomerID.QueryStringValue) Then bValidMaster = False
				Else
					bValidMaster = False
				End If
			End If
		End If
		If bValidMaster Then

			' Save current master table
			Orders.CurrentMasterTable = sMasterTblVar

			' Reset start record counter (new master key)
			StartRec = 1
			Orders.StartRecordNumber = StartRec

			' Clear previous master session values
			If sMasterTblVar <> "Customers" Then
				If Orders.CustomerId.QueryStringValue = "" Then Orders.CustomerId.SessionValue = ""
			End If
		End If
		DbMasterFilter = Orders.MasterFilter '  Get master filter
		DbDetailFilter = Orders.DetailFilter ' Get detail filter
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
	 Dim opt
	 Set opt = ListOptions.Add("Report")
	 opt.OnLeft = True ' Link on left
	' opt.MoveTo 0 ' Move to first column
End Sub                                

		' ListOptions Rendered event
Sub ListOptions_Rendered()
'Example: 
		if((Orders.payment_status.CurrentValue <> "WIP") AND  (Orders.payment_status.CurrentValue<>"Cancelled")) then    
			ListOptions.GetItem("Report").Body = "<a href=""sendemail.asp?orderid=" & Orders.OrderId.CurrentValue &""">Notify</a>"
		else 
			ListOptions.GetItem("Report").Body = ""
		end if
End Sub                                                                                   


End Class
%>
