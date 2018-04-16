<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Productsinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Products_list
Set Products_list = New cProducts_list
Set Page = Products_list

' Page init processing
Call Products_list.Page_Init()

' Page main processing
Call Products_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Products.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Products_list = new ew_Page("Products_list");
// page properties
Products_list.PageID = "list"; // page ID
Products_list.FormID = "fProductslist"; // form ID
var EW_PAGE_ID = Products_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Products_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Products_list.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Products_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Products_list.ValidateRequired = false; // no JavaScript validation
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
<% If (Products.Export = "") Or (EW_EXPORT_MASTER_RECORD And Products.Export = "print") Then %>
<% End If %>
<% Products_list.ShowPageHeader() %>
<%

' Load recordset
Set Products_list.Recordset = Products_list.LoadRecordset()
	Products_list.TotalRecs = Products_list.Recordset.RecordCount
	Products_list.StartRec = 1
	If Products_list.DisplayRecs <= 0 Then ' Display all records
		Products_list.DisplayRecs = Products_list.TotalRecs
	End If
	If Not (Products.ExportAll And Products.Export <> "") Then
		Products_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Products.TableCaption %>
&nbsp;&nbsp;<% Products_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Products.Export = "" And Products.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Products_list);" style="text-decoration: none;"><img id="Products_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Products_list_SearchPanel">
<form name="fProductslistsrch" id="fProductslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Products">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Products.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Products_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Products.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Products.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Products.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Products_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If Products.Export = "" Then %>
<div class="ewGridUpperPanel">
<% If Products.CurrentAction <> "gridadd" And Products.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Products_list.Pager) Then Set Products_list.Pager = ew_NewNumericPager(Products_list.StartRec, Products_list.DisplayRecs, Products_list.TotalRecs, Products_list.RecRange) %>
<% If Products_list.Pager.RecordCount > 0 Then %>
	<% If Products_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Products_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Products_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Products_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Products_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Products_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Products_list.Pager.RecordCount %>
<% Else %>
	<% If Products_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Products_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Products_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fProductslist, '<%= Products_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<form name="fProductslist" id="fProductslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Products">
<div id="gmp_Products" class="ewGridMiddlePanel">
<% If Products_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Products.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Products_list.RenderListOptions()

' Render list options (header, left)
Products_list.ListOptions.Render "header", "left"
%>
<% If Products.Description.Visible Then ' Description %>
	<% If Products.SortUrl(Products.Description) = "" Then %>
		<td><%= Products.Description.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.Description) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.Description.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.Description.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.Description.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.Price.Visible Then ' Price %>
	<% If Products.SortUrl(Products.Price) = "" Then %>
		<td><%= Products.Price.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.Price) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.Price.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.Price.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.Price.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.Active.Visible Then ' Active %>
	<% If Products.SortUrl(Products.Active) = "" Then %>
		<td><%= Products.Active.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.Active) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.Active.FldCaption %></td><td style="width: 10px;"><% If Products.Active.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.Active.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.Sizes.Visible Then ' Sizes %>
	<% If Products.SortUrl(Products.Sizes) = "" Then %>
		<td><%= Products.Sizes.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.Sizes) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.Sizes.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.Sizes.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.Sizes.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.Image_Thumb.Visible Then ' Image_Thumb %>
	<% If Products.SortUrl(Products.Image_Thumb) = "" Then %>
		<td><%= Products.Image_Thumb.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.Image_Thumb) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.Image_Thumb.FldCaption %></td><td style="width: 10px;"><% If Products.Image_Thumb.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.Image_Thumb.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.ProductName.Visible Then ' ProductName %>
	<% If Products.SortUrl(Products.ProductName) = "" Then %>
		<td><%= Products.ProductName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.ProductName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.ProductName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.ProductName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.ProductName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.ItemNo.Visible Then ' ItemNo %>
	<% If Products.SortUrl(Products.ItemNo) = "" Then %>
		<td><%= Products.ItemNo.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.ItemNo) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.ItemNo.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.ItemNo.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.ItemNo.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Products.UPC.Visible Then ' UPC %>
	<% If Products.SortUrl(Products.UPC) = "" Then %>
		<td><%= Products.UPC.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Products.SortUrl(Products.UPC) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Products.UPC.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Products.UPC.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Products.UPC.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Products_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Products.ExportAll And Products.Export <> "") Then
	Products_list.StopRec = Products_list.TotalRecs
Else

	' Set the last record to display
	If Products_list.TotalRecs > Products_list.StartRec + Products_list.DisplayRecs - 1 Then
		Products_list.StopRec = Products_list.StartRec + Products_list.DisplayRecs - 1
	Else
		Products_list.StopRec = Products_list.TotalRecs
	End If
End If

' Move to first record
Products_list.RecCnt = Products_list.StartRec - 1
If Not Products_list.Recordset.Eof Then
	Products_list.Recordset.MoveFirst
	If Products_list.StartRec > 1 Then Products_list.Recordset.Move Products_list.StartRec - 1
ElseIf Not Products.AllowAddDeleteRow And Products_list.StopRec = 0 Then
	Products_list.StopRec = Products.GridAddRowCount
End If

' Initialize Aggregate
Products.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Products.ResetAttrs()
Call Products_list.RenderRow()
Products_list.RowCnt = 0

' Output date rows
Do While CLng(Products_list.RecCnt) < CLng(Products_list.StopRec)
	Products_list.RecCnt = Products_list.RecCnt + 1
	If CLng(Products_list.RecCnt) >= CLng(Products_list.StartRec) Then
		Products_list.RowCnt = Products_list.RowCnt + 1

	' Set up key count
	Products_list.KeyCount = Products_list.RowIndex
	Call Products.ResetAttrs()
	Products.CssClass = ""
	If Products.CurrentAction = "gridadd" Then
	Else
		Call Products_list.LoadRowValues(Products_list.Recordset) ' Load row values
	End If
	Products.RowType = EW_ROWTYPE_VIEW ' Render view
	Products.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Products_list.RenderRow()

	' Render list options
	Call Products_list.RenderListOptions()
%>
	<tr<%= Products.RowAttributes %>>
<%

' Render list options (body, left)
Products_list.ListOptions.Render "body", "left"
%>
	<% If Products.Description.Visible Then ' Description %>
		<td<%= Products.Description.CellAttributes %>>
<div<%= Products.Description.ViewAttributes %>><%= Products.Description.ListViewValue %></div>
<a name="<%= Products_list.PageObjName & "_row_" & Products_list.RowCnt %>" id="<%= Products_list.PageObjName & "_row_" & Products_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Products.Price.Visible Then ' Price %>
		<td<%= Products.Price.CellAttributes %>>
<div<%= Products.Price.ViewAttributes %>><%= Products.Price.ListViewValue %></div>
</td>
	<% End If %>
	<% If Products.Active.Visible Then ' Active %>
		<td<%= Products.Active.CellAttributes %>>
<% If ew_ConvertToBool(Products.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Products.Active.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Products.Active.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
	<% If Products.Sizes.Visible Then ' Sizes %>
		<td<%= Products.Sizes.CellAttributes %>>
<div<%= Products.Sizes.ViewAttributes %>><%= Products.Sizes.ListViewValue %></div>
</td>
	<% End If %>
	<% If Products.Image_Thumb.Visible Then ' Image_Thumb %>
		<td<%= Products.Image_Thumb.CellAttributes %>>
<% If Products.Image_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% ElseIf Products.CurrentAction <> "edit" And Products.CurrentAction <> "gridedit" Then %>
&nbsp;
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% ElseIf Products.CurrentAction <> "edit" And Products.CurrentAction <> "gridedit" Then %>
&nbsp;
<% End If %>
<% End If %>
</td>
	<% End If %>
	<% If Products.ProductName.Visible Then ' ProductName %>
		<td<%= Products.ProductName.CellAttributes %>>
<div<%= Products.ProductName.ViewAttributes %>><%= Products.ProductName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Products.ItemNo.Visible Then ' ItemNo %>
		<td<%= Products.ItemNo.CellAttributes %>>
<div<%= Products.ItemNo.ViewAttributes %>><%= Products.ItemNo.ListViewValue %></div>
</td>
	<% End If %>
	<% If Products.UPC.Visible Then ' UPC %>
		<td<%= Products.UPC.CellAttributes %>>
<div<%= Products.UPC.ViewAttributes %>><%= Products.UPC.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Products_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Products.CurrentAction <> "gridadd" Then
		Products_list.Recordset.MoveNext()
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
Products_list.Recordset.Close
Set Products_list.Recordset = Nothing
%>
<% If Products_list.TotalRecs > 0 Then %>
<% If Products.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Products.CurrentAction <> "gridadd" And Products.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Products_list.Pager) Then Set Products_list.Pager = ew_NewNumericPager(Products_list.StartRec, Products_list.DisplayRecs, Products_list.TotalRecs, Products_list.RecRange) %>
<% If Products_list.Pager.RecordCount > 0 Then %>
	<% If Products_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Products_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Products_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Products_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Products_list.PageUrl %>start=<%= Products_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Products_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Products_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Products_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Products_list.Pager.RecordCount %>
<% Else %>
	<% If Products_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Products_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
<% If Products_list.TotalRecs > 0 Then %>
<% If Security.IsLoggedIn() Then %>
<a class="ewGridLink" href="" onclick="ew_SubmitSelected(document.fProductslist, '<%= Products_list.MultiDeleteUrl %>');return false;"><%= Language.Phrase("DeleteSelectedLink") %></a>&nbsp;&nbsp;
<% End If %>
<% End If %>
</span>
</div>
<% End If %>
<% End If %>
</td></tr></table>
<% If Products.Export = "" And Products.CurrentAction = "" Then %>
<% End If %>
<%
Products_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Products.Export = "" Then %>
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
Set Products_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProducts_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Products"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Products_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Products.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Products.TableVar & "&" ' add page token
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
		If Products.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Products.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Products.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Products) Then Set Products = New cProducts
		Set Table = Products

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Productsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Productsdelete.asp"
		MultiUpdateUrl = "Productsupdate.asp"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Products"

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
				Products.GridAddRowCount = gridaddcnt
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
		Set Products = Nothing
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
			If Products.Export <> "" Or Products.CurrentAction = "gridadd" Or Products.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Products.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Products.RecordsPerPage <> "" Then
			DisplayRecs = Products.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Products.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Products.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Products.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Products.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Products.SessionWhere = sFilter
		Products.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Products.Description, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.Price, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.Image, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.Sizes, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.Image_Thumb, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.ProductName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.ItemNo, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.UPC, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.Price_rebate, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.fDescription, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.fImage, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.fSizes, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.fImage_Thumb, Keyword)
			Call BuildBasicSearchSQL(sWhere, Products.fProductName, Keyword)
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
		sSearchKeyword = Products.BasicSearchKeyword
		sSearchType = Products.BasicSearchType
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
			Products.SessionBasicSearchKeyword = sSearchKeyword
			Products.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Products.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Products.SessionBasicSearchKeyword = ""
		Products.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Products.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Products.BasicSearchKeyword = Products.SessionBasicSearchKeyword
			Products.BasicSearchType = Products.SessionBasicSearchType
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
			Products.CurrentOrder = Request.QueryString("order")
			Products.CurrentOrderType = Request.QueryString("ordertype")

			' Field Description
			Call Products.UpdateSort(Products.Description)

			' Field Price
			Call Products.UpdateSort(Products.Price)

			' Field Active
			Call Products.UpdateSort(Products.Active)

			' Field Sizes
			Call Products.UpdateSort(Products.Sizes)

			' Field Image_Thumb
			Call Products.UpdateSort(Products.Image_Thumb)

			' Field ProductName
			Call Products.UpdateSort(Products.ProductName)

			' Field ItemNo
			Call Products.UpdateSort(Products.ItemNo)

			' Field UPC
			Call Products.UpdateSort(Products.UPC)
			Products.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Products.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Products.SqlOrderBy <> "" Then
				sOrderBy = Products.SqlOrderBy
				Products.SessionOrderBy = sOrderBy
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
				Products.SessionOrderBy = sOrderBy
				Products.Description.Sort = ""
				Products.Price.Sort = ""
				Products.Active.Sort = ""
				Products.Sizes.Sort = ""
				Products.Image_Thumb.Sort = ""
				Products.ProductName.Sort = ""
				Products.ItemNo.Sort = ""
				Products.UPC.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Products.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Header = "<input type=""checkbox"" name=""key"" id=""key"" class=""aspmaker"" onclick=""Products_list.SelectAllKey(this);"">"
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
			ListOptions.GetItem("checkbox").Body = "<input type=""checkbox"" name=""key_m"" id=""key_m"" value=""" & ew_HtmlEncode(Products.ItemId.CurrentValue) & """ class=""aspmaker"" onclick='ew_ClickMultiCheckbox(this);'>"
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
				Products.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Products.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Products.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Products.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Products.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Products.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Products.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Products.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Products.CurrentFilter
		Call Products.Recordset_Selecting(sFilter)
		Products.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Products.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Products.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Products.KeyFilter

		' Call Row Selecting event
		Call Products.Row_Selecting(sFilter)

		' Load sql based on filter
		Products.CurrentFilter = sFilter
		sSql = Products.SQL
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
		Call Products.Row_Selected(RsRow)
		Products.ItemId.DbValue = RsRow("ItemId")
		Products.Description.DbValue = RsRow("Description")
		Products.Price.DbValue = RsRow("Price")
		Products.Active.DbValue = ew_IIf(RsRow("Active"), "1", "0")
		Products.Image.Upload.DbValue = RsRow("Image")
		Products.Sizes.DbValue = RsRow("Sizes")
		Products.Image_Thumb.Upload.DbValue = RsRow("Image_Thumb")
		Products.ProductName.DbValue = RsRow("ProductName")
		Products.ItemNo.DbValue = RsRow("ItemNo")
		Products.UPC.DbValue = RsRow("UPC")
		Products.Price_rebate.DbValue = RsRow("Price_rebate")
		Products.fDescription.DbValue = RsRow("fDescription")
		Products.fImage.Upload.DbValue = RsRow("fImage")
		Products.fSizes.DbValue = RsRow("fSizes")
		Products.fImage_Thumb.Upload.DbValue = RsRow("fImage_Thumb")
		Products.fProductName.DbValue = RsRow("fProductName")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Products.GetKey("ItemId")&"" <> "" Then
			Products.ItemId.CurrentValue = Products.GetKey("ItemId") ' ItemId
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Products.CurrentFilter = Products.KeyFilter
			Dim sSql
			sSql = Products.SQL
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
		ViewUrl = Products.ViewUrl
		EditUrl = Products.EditUrl("")
		InlineEditUrl = Products.InlineEditUrl
		CopyUrl = Products.CopyUrl("")
		InlineCopyUrl = Products.InlineCopyUrl
		DeleteUrl = Products.DeleteUrl

		' Call Row Rendering event
		Call Products.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' ItemId
		' Description
		' Price
		' Active
		' Image
		' Sizes
		' Image_Thumb
		' ProductName
		' ItemNo
		' UPC
		' Price_rebate
		' fDescription
		' fImage
		' fSizes
		' fImage_Thumb
		' fProductName
		' -----------
		'  View  Row
		' -----------

		If Products.RowType = EW_ROWTYPE_VIEW Then ' View row

			' ItemId
			Products.ItemId.ViewValue = Products.ItemId.CurrentValue
			Products.ItemId.ViewCustomAttributes = ""

			' Description
			Products.Description.ViewValue = Products.Description.CurrentValue
			Products.Description.ViewCustomAttributes = ""

			' Price
			Products.Price.ViewValue = Products.Price.CurrentValue
			Products.Price.ViewCustomAttributes = ""

			' Active
			If ew_ConvertToBool(Products.Active.CurrentValue) Then
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(1) <> "", Products.Active.FldTagCaption(1), "Yes")
			Else
				Products.Active.ViewValue = ew_IIf(Products.Active.FldTagCaption(2) <> "", Products.Active.FldTagCaption(2), "No")
			End If
			Products.Active.ViewCustomAttributes = ""

			' Image
			If Not ew_Empty(Products.Image.Upload.DbValue) Then
				Products.Image.ViewValue = Products.Image.Upload.DbValue
				Products.Image.ImageAlt = Products.Image.FldAlt
			Else
				Products.Image.ViewValue = ""
			End If
			Products.Image.ViewCustomAttributes = ""

			' Sizes
			Products.Sizes.ViewValue = Products.Sizes.CurrentValue
			Products.Sizes.ViewCustomAttributes = ""

			' Image_Thumb
			If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then
				Products.Image_Thumb.ViewValue = Products.Image_Thumb.Upload.DbValue
				Products.Image_Thumb.ImageAlt = Products.Image_Thumb.FldAlt
			Else
				Products.Image_Thumb.ViewValue = ""
			End If
			Products.Image_Thumb.ViewCustomAttributes = ""

			' ProductName
			Products.ProductName.ViewValue = Products.ProductName.CurrentValue
			Products.ProductName.ViewCustomAttributes = ""

			' ItemNo
			Products.ItemNo.ViewValue = Products.ItemNo.CurrentValue
			Products.ItemNo.ViewCustomAttributes = ""

			' UPC
			Products.UPC.ViewValue = Products.UPC.CurrentValue
			Products.UPC.ViewCustomAttributes = ""

			' Price_rebate
			Products.Price_rebate.ViewValue = Products.Price_rebate.CurrentValue
			Products.Price_rebate.ViewCustomAttributes = ""

			' fDescription
			Products.fDescription.ViewValue = Products.fDescription.CurrentValue
			Products.fDescription.ViewCustomAttributes = ""

			' fImage
			If Not ew_Empty(Products.fImage.Upload.DbValue) Then
				Products.fImage.ViewValue = Products.fImage.Upload.DbValue
				Products.fImage.ImageAlt = Products.fImage.FldAlt
			Else
				Products.fImage.ViewValue = ""
			End If
			Products.fImage.ViewCustomAttributes = ""

			' fSizes
			Products.fSizes.ViewValue = Products.fSizes.CurrentValue
			Products.fSizes.ViewCustomAttributes = ""

			' fImage_Thumb
			If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then
				Products.fImage_Thumb.ViewValue = Products.fImage_Thumb.Upload.DbValue
				Products.fImage_Thumb.ImageAlt = Products.fImage_Thumb.FldAlt
			Else
				Products.fImage_Thumb.ViewValue = ""
			End If
			Products.fImage_Thumb.ViewCustomAttributes = ""

			' fProductName
			Products.fProductName.ViewValue = Products.fProductName.CurrentValue
			Products.fProductName.ViewCustomAttributes = ""

			' View refer script
			' Description

			Products.Description.LinkCustomAttributes = ""
			Products.Description.HrefValue = ""
			Products.Description.TooltipValue = ""

			' Price
			Products.Price.LinkCustomAttributes = ""
			Products.Price.HrefValue = ""
			Products.Price.TooltipValue = ""

			' Active
			Products.Active.LinkCustomAttributes = ""
			Products.Active.HrefValue = ""
			Products.Active.TooltipValue = ""

			' Sizes
			Products.Sizes.LinkCustomAttributes = ""
			Products.Sizes.HrefValue = ""
			Products.Sizes.TooltipValue = ""

			' Image_Thumb
			Products.Image_Thumb.LinkCustomAttributes = ""
			Products.Image_Thumb.HrefValue = ""
			Products.Image_Thumb.TooltipValue = ""

			' ProductName
			Products.ProductName.LinkCustomAttributes = ""
			Products.ProductName.HrefValue = ""
			Products.ProductName.TooltipValue = ""

			' ItemNo
			Products.ItemNo.LinkCustomAttributes = ""
			Products.ItemNo.HrefValue = ""
			Products.ItemNo.TooltipValue = ""

			' UPC
			Products.UPC.LinkCustomAttributes = ""
			Products.UPC.HrefValue = ""
			Products.UPC.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Products.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Products.Row_Rendered()
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
