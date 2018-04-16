<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Shippinginfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Shipping_list
Set Shipping_list = New cShipping_list
Set Page = Shipping_list

' Page init processing
Call Shipping_list.Page_Init()

' Page main processing
Call Shipping_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Shipping.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Shipping_list = new ew_Page("Shipping_list");
// page properties
Shipping_list.PageID = "list"; // page ID
Shipping_list.FormID = "fShippinglist"; // form ID
var EW_PAGE_ID = Shipping_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Shipping_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
Shipping_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Shipping_list.ValidateRequired = false; // no JavaScript validation
<% End If %>
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
<% If (Shipping.Export = "") Or (EW_EXPORT_MASTER_RECORD And Shipping.Export = "print") Then %>
<% End If %>
<% Shipping_list.ShowPageHeader() %>
<%

' Load recordset
Set Shipping_list.Recordset = Shipping_list.LoadRecordset()
	Shipping_list.TotalRecs = Shipping_list.Recordset.RecordCount
	Shipping_list.StartRec = 1
	If Shipping_list.DisplayRecs <= 0 Then ' Display all records
		Shipping_list.DisplayRecs = Shipping_list.TotalRecs
	End If
	If Not (Shipping.ExportAll And Shipping.Export <> "") Then
		Shipping_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeTABLE") %><%= Shipping.TableCaption %>
&nbsp;&nbsp;<% Shipping_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If Shipping.Export = "" And Shipping.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(Shipping_list);" style="text-decoration: none;"><img id="Shipping_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="Shipping_list_SearchPanel">
<form name="fShippinglistsrch" id="fShippinglistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="Shipping">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(Shipping.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= Shipping_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If Shipping.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Shipping.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Shipping.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% Shipping_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<form name="fShippinglist" id="fShippinglist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="Shipping">
<div id="gmp_Shipping" class="ewGridMiddlePanel">
<% If Shipping_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Shipping.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Shipping_list.RenderListOptions()

' Render list options (header, left)
Shipping_list.ListOptions.Render "header", "left"
%>
<% If Shipping.CustomerId.Visible Then ' CustomerId %>
	<% If Shipping.SortUrl(Shipping.CustomerId) = "" Then %>
		<td><%= Shipping.CustomerId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.CustomerId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.CustomerId.FldCaption %></td><td style="width: 10px;"><% If Shipping.CustomerId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.CustomerId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_FirstName.Visible Then ' ship_FirstName %>
	<% If Shipping.SortUrl(Shipping.ship_FirstName) = "" Then %>
		<td><%= Shipping.ship_FirstName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_FirstName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_FirstName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_LastName.Visible Then ' ship_LastName %>
	<% If Shipping.SortUrl(Shipping.ship_LastName) = "" Then %>
		<td><%= Shipping.ship_LastName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_LastName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_LastName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_Address.Visible Then ' ship_Address %>
	<% If Shipping.SortUrl(Shipping.ship_Address) = "" Then %>
		<td><%= Shipping.ship_Address.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_Address) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_Address.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_Address.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_Address.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_City.Visible Then ' ship_City %>
	<% If Shipping.SortUrl(Shipping.ship_City) = "" Then %>
		<td><%= Shipping.ship_City.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_City) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_City.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_City.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_City.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_Province.Visible Then ' ship_Province %>
	<% If Shipping.SortUrl(Shipping.ship_Province) = "" Then %>
		<td><%= Shipping.ship_Province.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_Province) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_Province.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_Province.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_Province.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_PostalCode.Visible Then ' ship_PostalCode %>
	<% If Shipping.SortUrl(Shipping.ship_PostalCode) = "" Then %>
		<td><%= Shipping.ship_PostalCode.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_PostalCode) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_PostalCode.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_PostalCode.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_PostalCode.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_Country.Visible Then ' ship_Country %>
	<% If Shipping.SortUrl(Shipping.ship_Country) = "" Then %>
		<td><%= Shipping.ship_Country.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_Country) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_Country.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_Country.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_Country.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
	<% If Shipping.SortUrl(Shipping.ship_EmailAddress) = "" Then %>
		<td><%= Shipping.ship_EmailAddress.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_EmailAddress) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_EmailAddress.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_EmailAddress.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_EmailAddress.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.HomePhone.Visible Then ' HomePhone %>
	<% If Shipping.SortUrl(Shipping.HomePhone) = "" Then %>
		<td><%= Shipping.HomePhone.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.HomePhone) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.HomePhone.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.HomePhone.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.HomePhone.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.WorkPhone.Visible Then ' WorkPhone %>
	<% If Shipping.SortUrl(Shipping.WorkPhone) = "" Then %>
		<td><%= Shipping.WorkPhone.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.WorkPhone) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.WorkPhone.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.WorkPhone.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.WorkPhone.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Shipping.ship_Address2.Visible Then ' ship_Address2 %>
	<% If Shipping.SortUrl(Shipping.ship_Address2) = "" Then %>
		<td><%= Shipping.ship_Address2.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= Shipping.SortUrl(Shipping.ship_Address2) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Shipping.ship_Address2.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If Shipping.ship_Address2.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Shipping.ship_Address2.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Shipping_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Shipping.ExportAll And Shipping.Export <> "") Then
	Shipping_list.StopRec = Shipping_list.TotalRecs
Else

	' Set the last record to display
	If Shipping_list.TotalRecs > Shipping_list.StartRec + Shipping_list.DisplayRecs - 1 Then
		Shipping_list.StopRec = Shipping_list.StartRec + Shipping_list.DisplayRecs - 1
	Else
		Shipping_list.StopRec = Shipping_list.TotalRecs
	End If
End If

' Move to first record
Shipping_list.RecCnt = Shipping_list.StartRec - 1
If Not Shipping_list.Recordset.Eof Then
	Shipping_list.Recordset.MoveFirst
	If Shipping_list.StartRec > 1 Then Shipping_list.Recordset.Move Shipping_list.StartRec - 1
ElseIf Not Shipping.AllowAddDeleteRow And Shipping_list.StopRec = 0 Then
	Shipping_list.StopRec = Shipping.GridAddRowCount
End If

' Initialize Aggregate
Shipping.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Shipping.ResetAttrs()
Call Shipping_list.RenderRow()
Shipping_list.RowCnt = 0

' Output date rows
Do While CLng(Shipping_list.RecCnt) < CLng(Shipping_list.StopRec)
	Shipping_list.RecCnt = Shipping_list.RecCnt + 1
	If CLng(Shipping_list.RecCnt) >= CLng(Shipping_list.StartRec) Then
		Shipping_list.RowCnt = Shipping_list.RowCnt + 1

	' Set up key count
	Shipping_list.KeyCount = Shipping_list.RowIndex
	Call Shipping.ResetAttrs()
	Shipping.CssClass = ""
	If Shipping.CurrentAction = "gridadd" Then
	Else
		Call Shipping_list.LoadRowValues(Shipping_list.Recordset) ' Load row values
	End If
	Shipping.RowType = EW_ROWTYPE_VIEW ' Render view
	Shipping.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call Shipping_list.RenderRow()

	' Render list options
	Call Shipping_list.RenderListOptions()
%>
	<tr<%= Shipping.RowAttributes %>>
<%

' Render list options (body, left)
Shipping_list.ListOptions.Render "body", "left"
%>
	<% If Shipping.CustomerId.Visible Then ' CustomerId %>
		<td<%= Shipping.CustomerId.CellAttributes %>>
<div<%= Shipping.CustomerId.ViewAttributes %>><%= Shipping.CustomerId.ListViewValue %></div>
<a name="<%= Shipping_list.PageObjName & "_row_" & Shipping_list.RowCnt %>" id="<%= Shipping_list.PageObjName & "_row_" & Shipping_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Shipping.ship_FirstName.Visible Then ' ship_FirstName %>
		<td<%= Shipping.ship_FirstName.CellAttributes %>>
<div<%= Shipping.ship_FirstName.ViewAttributes %>><%= Shipping.ship_FirstName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_LastName.Visible Then ' ship_LastName %>
		<td<%= Shipping.ship_LastName.CellAttributes %>>
<div<%= Shipping.ship_LastName.ViewAttributes %>><%= Shipping.ship_LastName.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_Address.Visible Then ' ship_Address %>
		<td<%= Shipping.ship_Address.CellAttributes %>>
<div<%= Shipping.ship_Address.ViewAttributes %>><%= Shipping.ship_Address.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_City.Visible Then ' ship_City %>
		<td<%= Shipping.ship_City.CellAttributes %>>
<div<%= Shipping.ship_City.ViewAttributes %>><%= Shipping.ship_City.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_Province.Visible Then ' ship_Province %>
		<td<%= Shipping.ship_Province.CellAttributes %>>
<div<%= Shipping.ship_Province.ViewAttributes %>><%= Shipping.ship_Province.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_PostalCode.Visible Then ' ship_PostalCode %>
		<td<%= Shipping.ship_PostalCode.CellAttributes %>>
<div<%= Shipping.ship_PostalCode.ViewAttributes %>><%= Shipping.ship_PostalCode.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_Country.Visible Then ' ship_Country %>
		<td<%= Shipping.ship_Country.CellAttributes %>>
<div<%= Shipping.ship_Country.ViewAttributes %>><%= Shipping.ship_Country.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
		<td<%= Shipping.ship_EmailAddress.CellAttributes %>>
<div<%= Shipping.ship_EmailAddress.ViewAttributes %>><%= Shipping.ship_EmailAddress.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.HomePhone.Visible Then ' HomePhone %>
		<td<%= Shipping.HomePhone.CellAttributes %>>
<div<%= Shipping.HomePhone.ViewAttributes %>><%= Shipping.HomePhone.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.WorkPhone.Visible Then ' WorkPhone %>
		<td<%= Shipping.WorkPhone.CellAttributes %>>
<div<%= Shipping.WorkPhone.ViewAttributes %>><%= Shipping.WorkPhone.ListViewValue %></div>
</td>
	<% End If %>
	<% If Shipping.ship_Address2.Visible Then ' ship_Address2 %>
		<td<%= Shipping.ship_Address2.CellAttributes %>>
<div<%= Shipping.ship_Address2.ViewAttributes %>><%= Shipping.ship_Address2.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
Shipping_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If Shipping.CurrentAction <> "gridadd" Then
		Shipping_list.Recordset.MoveNext()
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
Shipping_list.Recordset.Close
Set Shipping_list.Recordset = Nothing
%>
<% If Shipping.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Shipping.CurrentAction <> "gridadd" And Shipping.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(Shipping_list.Pager) Then Set Shipping_list.Pager = ew_NewNumericPager(Shipping_list.StartRec, Shipping_list.DisplayRecs, Shipping_list.TotalRecs, Shipping_list.RecRange) %>
<% If Shipping_list.Pager.RecordCount > 0 Then %>
	<% If Shipping_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= Shipping_list.PageUrl %>start=<%= Shipping_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If Shipping_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= Shipping_list.PageUrl %>start=<%= Shipping_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Shipping_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= Shipping_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Shipping_list.Pager.NextButton.Enabled Then %>
	<a href="<%= Shipping_list.PageUrl %>start=<%= Shipping_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If Shipping_list.Pager.LastButton.Enabled Then %>
	<a href="<%= Shipping_list.PageUrl %>start=<%= Shipping_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If Shipping_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= Shipping_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Shipping_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Shipping_list.Pager.RecordCount %>
<% Else %>
	<% If Shipping_list.SearchWhere = "0=101" Then %>
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
<a class="ewGridLink" href="<%= Shipping_list.AddUrl %>"><%= Language.Phrase("AddLink") %></a>&nbsp;&nbsp;
<% End If %>
</span>
</div>
<% End If %>
</td></tr></table>
<% If Shipping.Export = "" And Shipping.CurrentAction = "" Then %>
<% End If %>
<%
Shipping_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Shipping.Export = "" Then %>
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
Set Shipping_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cShipping_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Shipping"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Shipping_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If Shipping.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Shipping.TableVar & "&" ' add page token
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
		If Shipping.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Shipping.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Shipping.TableVar = Request.QueryString("t"))
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
		If IsEmpty(Shipping) Then Set Shipping = New cShipping
		Set Table = Shipping

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "Shippingadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "Shippingdelete.asp"
		MultiUpdateUrl = "Shippingupdate.asp"

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Shipping"

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
				Shipping.GridAddRowCount = gridaddcnt
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
		Set Shipping = Nothing
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
			If Shipping.Export <> "" Or Shipping.CurrentAction = "gridadd" Or Shipping.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call Shipping.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Shipping.RecordsPerPage <> "" Then
			DisplayRecs = Shipping.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Shipping.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			Shipping.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				Shipping.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = Shipping.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Shipping.SessionWhere = sFilter
		Shipping.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_FirstName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_LastName, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_PostalCode, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_EmailAddress, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.HomePhone, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.WorkPhone, Keyword)
			Call BuildBasicSearchSQL(sWhere, Shipping.ship_Address2, Keyword)
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
		sSearchKeyword = Shipping.BasicSearchKeyword
		sSearchType = Shipping.BasicSearchType
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
			Shipping.SessionBasicSearchKeyword = sSearchKeyword
			Shipping.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Shipping.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		Shipping.SessionBasicSearchKeyword = ""
		Shipping.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If Shipping.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			Shipping.BasicSearchKeyword = Shipping.SessionBasicSearchKeyword
			Shipping.BasicSearchType = Shipping.SessionBasicSearchType
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
			Shipping.CurrentOrder = Request.QueryString("order")
			Shipping.CurrentOrderType = Request.QueryString("ordertype")

			' Field CustomerId
			Call Shipping.UpdateSort(Shipping.CustomerId)

			' Field ship_FirstName
			Call Shipping.UpdateSort(Shipping.ship_FirstName)

			' Field ship_LastName
			Call Shipping.UpdateSort(Shipping.ship_LastName)

			' Field ship_Address
			Call Shipping.UpdateSort(Shipping.ship_Address)

			' Field ship_City
			Call Shipping.UpdateSort(Shipping.ship_City)

			' Field ship_Province
			Call Shipping.UpdateSort(Shipping.ship_Province)

			' Field ship_PostalCode
			Call Shipping.UpdateSort(Shipping.ship_PostalCode)

			' Field ship_Country
			Call Shipping.UpdateSort(Shipping.ship_Country)

			' Field ship_EmailAddress
			Call Shipping.UpdateSort(Shipping.ship_EmailAddress)

			' Field HomePhone
			Call Shipping.UpdateSort(Shipping.HomePhone)

			' Field WorkPhone
			Call Shipping.UpdateSort(Shipping.WorkPhone)

			' Field ship_Address2
			Call Shipping.UpdateSort(Shipping.ship_Address2)
			Shipping.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Shipping.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Shipping.SqlOrderBy <> "" Then
				sOrderBy = Shipping.SqlOrderBy
				Shipping.SessionOrderBy = sOrderBy
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
				Shipping.SessionOrderBy = sOrderBy
				Shipping.CustomerId.Sort = ""
				Shipping.ship_FirstName.Sort = ""
				Shipping.ship_LastName.Sort = ""
				Shipping.ship_Address.Sort = ""
				Shipping.ship_City.Sort = ""
				Shipping.ship_Province.Sort = ""
				Shipping.ship_PostalCode.Sort = ""
				Shipping.ship_Country.Sort = ""
				Shipping.ship_EmailAddress.Sort = ""
				Shipping.HomePhone.Sort = ""
				Shipping.WorkPhone.Sort = ""
				Shipping.ship_Address2.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Shipping.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		ListOptions.Add("view")
		ListOptions.GetItem("view").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("view").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("view").OnLeft = False
		ListOptions.Add("edit")
		ListOptions.GetItem("edit").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("edit").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("edit").OnLeft = False
		ListOptions.Add("copy")
		ListOptions.GetItem("copy").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("copy").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("copy").OnLeft = False
		ListOptions.Add("delete")
		ListOptions.GetItem("delete").CssStyle = "white-space: nowrap;"
		ListOptions.GetItem("delete").Visible = Security.IsLoggedIn()
		ListOptions.GetItem("delete").OnLeft = False
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
		If Security.IsLoggedIn() And ListOptions.GetItem("delete").Visible Then
			ListOptions.GetItem("delete").Body = "<a class=""ewRowLink""" & "" & " href=""" & DeleteUrl & """>" & "<img src=""images/delete.gif"" alt=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ title=""" & ew_HtmlEncode(Language.Phrase("DeleteLink")) & """ width=""16"" height=""16"" border=""0"">" & "</a>"
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
				Shipping.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Shipping.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Shipping.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Shipping.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Shipping.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Shipping.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Shipping.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		Shipping.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Shipping.CurrentFilter
		Call Shipping.Recordset_Selecting(sFilter)
		Shipping.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Shipping.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Shipping.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Shipping.KeyFilter

		' Call Row Selecting event
		Call Shipping.Row_Selecting(sFilter)

		' Load sql based on filter
		Shipping.CurrentFilter = sFilter
		sSql = Shipping.SQL
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
		Call Shipping.Row_Selected(RsRow)
		Shipping.AddressID.DbValue = RsRow("AddressID")
		Shipping.CustomerId.DbValue = RsRow("CustomerId")
		Shipping.ship_FirstName.DbValue = RsRow("ship_FirstName")
		Shipping.ship_LastName.DbValue = RsRow("ship_LastName")
		Shipping.ship_Address.DbValue = RsRow("ship_Address")
		Shipping.ship_City.DbValue = RsRow("ship_City")
		Shipping.ship_Province.DbValue = RsRow("ship_Province")
		Shipping.ship_PostalCode.DbValue = RsRow("ship_PostalCode")
		Shipping.ship_Country.DbValue = RsRow("ship_Country")
		Shipping.ship_EmailAddress.DbValue = RsRow("ship_EmailAddress")
		Shipping.HomePhone.DbValue = RsRow("HomePhone")
		Shipping.WorkPhone.DbValue = RsRow("WorkPhone")
		Shipping.ship_Address2.DbValue = RsRow("ship_Address2")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If Shipping.GetKey("AddressID")&"" <> "" Then
			Shipping.AddressID.CurrentValue = Shipping.GetKey("AddressID") ' AddressID
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			Shipping.CurrentFilter = Shipping.KeyFilter
			Dim sSql
			sSql = Shipping.SQL
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
		ViewUrl = Shipping.ViewUrl
		EditUrl = Shipping.EditUrl("")
		InlineEditUrl = Shipping.InlineEditUrl
		CopyUrl = Shipping.CopyUrl("")
		InlineCopyUrl = Shipping.InlineCopyUrl
		DeleteUrl = Shipping.DeleteUrl

		' Call Row Rendering event
		Call Shipping.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' AddressID
		' CustomerId
		' ship_FirstName
		' ship_LastName
		' ship_Address
		' ship_City
		' ship_Province
		' ship_PostalCode
		' ship_Country
		' ship_EmailAddress
		' HomePhone
		' WorkPhone
		' ship_Address2
		' -----------
		'  View  Row
		' -----------

		If Shipping.RowType = EW_ROWTYPE_VIEW Then ' View row

			' AddressID
			Shipping.AddressID.ViewValue = Shipping.AddressID.CurrentValue
			Shipping.AddressID.ViewCustomAttributes = ""

			' CustomerId
			If Shipping.CustomerId.CurrentValue & "" <> "" Then
				sFilterWrk = "[CustomerID] = " & ew_AdjustSql(Shipping.CustomerId.CurrentValue) & ""
			sSqlWrk = "SELECT [Inv_FirstName], [Inv_LastName] FROM [Customers]"
			sWhereWrk = ""
			Call ew_AddFilter(sWhereWrk, sFilterWrk)
			If sWhereWrk <> "" Then sSqlWrk = sSqlWrk & " WHERE " & sWhereWrk
				Set RsWrk = Conn.Execute(sSqlWrk)
				If Not RsWrk.Eof Then
					Shipping.CustomerId.ViewValue = RsWrk("Inv_FirstName")
					Shipping.CustomerId.ViewValue = Shipping.CustomerId.ViewValue & ew_ValueSeparator(0,1,Shipping.CustomerId) & RsWrk("Inv_LastName")
				Else
					Shipping.CustomerId.ViewValue = Shipping.CustomerId.CurrentValue
				End If
				RsWrk.Close
				Set RsWrk = Nothing
			Else
				Shipping.CustomerId.ViewValue = Null
			End If
			Shipping.CustomerId.ViewCustomAttributes = ""

			' ship_FirstName
			Shipping.ship_FirstName.ViewValue = Shipping.ship_FirstName.CurrentValue
			Shipping.ship_FirstName.ViewCustomAttributes = ""

			' ship_LastName
			Shipping.ship_LastName.ViewValue = Shipping.ship_LastName.CurrentValue
			Shipping.ship_LastName.ViewCustomAttributes = ""

			' ship_Address
			Shipping.ship_Address.ViewValue = Shipping.ship_Address.CurrentValue
			Shipping.ship_Address.ViewCustomAttributes = ""

			' ship_City
			Shipping.ship_City.ViewValue = Shipping.ship_City.CurrentValue
			Shipping.ship_City.ViewCustomAttributes = ""

			' ship_Province
			Shipping.ship_Province.ViewValue = Shipping.ship_Province.CurrentValue
			Shipping.ship_Province.ViewCustomAttributes = ""

			' ship_PostalCode
			Shipping.ship_PostalCode.ViewValue = Shipping.ship_PostalCode.CurrentValue
			Shipping.ship_PostalCode.ViewCustomAttributes = ""

			' ship_Country
			Shipping.ship_Country.ViewValue = Shipping.ship_Country.CurrentValue
			Shipping.ship_Country.ViewCustomAttributes = ""

			' ship_EmailAddress
			Shipping.ship_EmailAddress.ViewValue = Shipping.ship_EmailAddress.CurrentValue
			Shipping.ship_EmailAddress.ViewCustomAttributes = ""

			' HomePhone
			Shipping.HomePhone.ViewValue = Shipping.HomePhone.CurrentValue
			Shipping.HomePhone.ViewCustomAttributes = ""

			' WorkPhone
			Shipping.WorkPhone.ViewValue = Shipping.WorkPhone.CurrentValue
			Shipping.WorkPhone.ViewCustomAttributes = ""

			' ship_Address2
			Shipping.ship_Address2.ViewValue = Shipping.ship_Address2.CurrentValue
			Shipping.ship_Address2.ViewCustomAttributes = ""

			' View refer script
			' CustomerId

			Shipping.CustomerId.LinkCustomAttributes = ""
			Shipping.CustomerId.HrefValue = ""
			Shipping.CustomerId.TooltipValue = ""

			' ship_FirstName
			Shipping.ship_FirstName.LinkCustomAttributes = ""
			Shipping.ship_FirstName.HrefValue = ""
			Shipping.ship_FirstName.TooltipValue = ""

			' ship_LastName
			Shipping.ship_LastName.LinkCustomAttributes = ""
			Shipping.ship_LastName.HrefValue = ""
			Shipping.ship_LastName.TooltipValue = ""

			' ship_Address
			Shipping.ship_Address.LinkCustomAttributes = ""
			Shipping.ship_Address.HrefValue = ""
			Shipping.ship_Address.TooltipValue = ""

			' ship_City
			Shipping.ship_City.LinkCustomAttributes = ""
			Shipping.ship_City.HrefValue = ""
			Shipping.ship_City.TooltipValue = ""

			' ship_Province
			Shipping.ship_Province.LinkCustomAttributes = ""
			Shipping.ship_Province.HrefValue = ""
			Shipping.ship_Province.TooltipValue = ""

			' ship_PostalCode
			Shipping.ship_PostalCode.LinkCustomAttributes = ""
			Shipping.ship_PostalCode.HrefValue = ""
			Shipping.ship_PostalCode.TooltipValue = ""

			' ship_Country
			Shipping.ship_Country.LinkCustomAttributes = ""
			Shipping.ship_Country.HrefValue = ""
			Shipping.ship_Country.TooltipValue = ""

			' ship_EmailAddress
			Shipping.ship_EmailAddress.LinkCustomAttributes = ""
			Shipping.ship_EmailAddress.HrefValue = ""
			Shipping.ship_EmailAddress.TooltipValue = ""

			' HomePhone
			Shipping.HomePhone.LinkCustomAttributes = ""
			Shipping.HomePhone.HrefValue = ""
			Shipping.HomePhone.TooltipValue = ""

			' WorkPhone
			Shipping.WorkPhone.LinkCustomAttributes = ""
			Shipping.WorkPhone.HrefValue = ""
			Shipping.WorkPhone.TooltipValue = ""

			' ship_Address2
			Shipping.ship_Address2.LinkCustomAttributes = ""
			Shipping.ship_Address2.HrefValue = ""
			Shipping.ship_Address2.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Shipping.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Shipping.Row_Rendered()
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
