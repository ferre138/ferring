<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="QAddressinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim QAddress_list
Set QAddress_list = New cQAddress_list
Set Page = QAddress_list

' Page init processing
Call QAddress_list.Page_Init()

' Page main processing
Call QAddress_list.Page_Main()
%>
<!--#include file="header.asp"-->
<% If QAddress.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var QAddress_list = new ew_Page("QAddress_list");
// page properties
QAddress_list.PageID = "list"; // page ID
QAddress_list.FormID = "fQAddresslist"; // form ID
var EW_PAGE_ID = QAddress_list.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
QAddress_list.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
QAddress_list.ValidateRequired = true; // uses JavaScript validation
<% Else %>
QAddress_list.ValidateRequired = false; // no JavaScript validation
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
<% If (QAddress.Export = "") Or (EW_EXPORT_MASTER_RECORD And QAddress.Export = "print") Then %>
<% End If %>
<% QAddress_list.ShowPageHeader() %>
<%

' Load recordset
Set QAddress_list.Recordset = QAddress_list.LoadRecordset()
	QAddress_list.TotalRecs = QAddress_list.Recordset.RecordCount
	QAddress_list.StartRec = 1
	If QAddress_list.DisplayRecs <= 0 Then ' Display all records
		QAddress_list.DisplayRecs = QAddress_list.TotalRecs
	End If
	If Not (QAddress.ExportAll And QAddress.Export <> "") Then
		QAddress_list.SetUpStartRec() ' Set up start record position
	End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><%= Language.Phrase("TblTypeVIEW") %><%= QAddress.TableCaption %>
&nbsp;&nbsp;<% QAddress_list.ExportOptions.Render "body", "" %>
</p>
<% If Security.IsLoggedIn() Then %>
<% If QAddress.Export = "" And QAddress.CurrentAction = "" Then %>
<a href="javascript:ew_ToggleSearchPanel(QAddress_list);" style="text-decoration: none;"><img id="QAddress_list_SearchImage" src="images/collapse.gif" alt="" width="9" height="9" border="0"></a><span class="aspmaker">&nbsp;<%= Language.Phrase("Search") %></span><br>
<div id="QAddress_list_SearchPanel">
<form name="fQAddresslistsrch" id="fQAddresslistsrch" class="ewForm" action="<%= ew_CurrentPage %>">
<input type="hidden" id="t" name="t" value="QAddress">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewCssTableRow">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= ew_HtmlEncode(QAddress.SessionBasicSearchKeyword) %>">
	<input type="Submit" name="Submit" id="Submit" value="<%= ew_BtnCaption(Language.Phrase("QuickSearchBtn")) %>">&nbsp;
	<a href="<%= QAddress_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>&nbsp;
</div>
<div id="xsr_2" class="ewCssTableRow">
	<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value=""<% If QAddress.SessionBasicSearchType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If QAddress.SessionBasicSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>&nbsp;&nbsp;<label><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If QAddress.SessionBasicSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</form>
</div>
<% End If %>
<% End If %>
<% QAddress_list.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<form name="fQAddresslist" id="fQAddresslist" class="ewForm" action="" method="post">
<input type="hidden" name="t" id="t" value="QAddress">
<div id="gmp_QAddress" class="ewGridMiddlePanel">
<% If QAddress_list.TotalRecs > 0 Then %>
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= QAddress.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call QAddress_list.RenderListOptions()

' Render list options (header, left)
QAddress_list.ListOptions.Render "header", "left"
%>
<% If QAddress.Customers2ECustomerId.Visible Then ' Customers.CustomerId %>
	<% If QAddress.SortUrl(QAddress.Customers2ECustomerId) = "" Then %>
		<td><%= QAddress.Customers2ECustomerId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Customers2ECustomerId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Customers2ECustomerId.FldCaption %></td><td style="width: 10px;"><% If QAddress.Customers2ECustomerId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Customers2ECustomerId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.Inv_FirstName.Visible Then ' Inv_FirstName %>
	<% If QAddress.SortUrl(QAddress.Inv_FirstName) = "" Then %>
		<td><%= QAddress.Inv_FirstName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Inv_FirstName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Inv_FirstName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.Inv_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Inv_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.Inv_LastName.Visible Then ' Inv_LastName %>
	<% If QAddress.SortUrl(QAddress.Inv_LastName) = "" Then %>
		<td><%= QAddress.Inv_LastName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Inv_LastName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Inv_LastName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.Inv_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Inv_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.Inv_Address.Visible Then ' Inv_Address %>
	<% If QAddress.SortUrl(QAddress.Inv_Address) = "" Then %>
		<td><%= QAddress.Inv_Address.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Inv_Address) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Inv_Address.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.Inv_Address.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Inv_Address.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_City.Visible Then ' inv_City %>
	<% If QAddress.SortUrl(QAddress.inv_City) = "" Then %>
		<td><%= QAddress.inv_City.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_City) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_City.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_City.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_City.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_Province.Visible Then ' inv_Province %>
	<% If QAddress.SortUrl(QAddress.inv_Province) = "" Then %>
		<td><%= QAddress.inv_Province.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_Province) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_Province.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_Province.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_Province.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_PostalCode.Visible Then ' inv_PostalCode %>
	<% If QAddress.SortUrl(QAddress.inv_PostalCode) = "" Then %>
		<td><%= QAddress.inv_PostalCode.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_PostalCode) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_PostalCode.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_PostalCode.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_PostalCode.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_Country.Visible Then ' inv_Country %>
	<% If QAddress.SortUrl(QAddress.inv_Country) = "" Then %>
		<td><%= QAddress.inv_Country.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_Country) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_Country.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_Country.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_Country.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_PhoneNumber.Visible Then ' inv_PhoneNumber %>
	<% If QAddress.SortUrl(QAddress.inv_PhoneNumber) = "" Then %>
		<td><%= QAddress.inv_PhoneNumber.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_PhoneNumber) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_PhoneNumber.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_PhoneNumber.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_PhoneNumber.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
	<% If QAddress.SortUrl(QAddress.inv_EmailAddress) = "" Then %>
		<td><%= QAddress.inv_EmailAddress.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_EmailAddress) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_EmailAddress.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_EmailAddress.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_EmailAddress.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.inv_Fax.Visible Then ' inv_Fax %>
	<% If QAddress.SortUrl(QAddress.inv_Fax) = "" Then %>
		<td><%= QAddress.inv_Fax.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.inv_Fax) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.inv_Fax.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.inv_Fax.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.inv_Fax.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.Inv_Address2.Visible Then ' Inv_Address2 %>
	<% If QAddress.SortUrl(QAddress.Inv_Address2) = "" Then %>
		<td><%= QAddress.Inv_Address2.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Inv_Address2) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Inv_Address2.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.Inv_Address2.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Inv_Address2.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.UserName.Visible Then ' UserName %>
	<% If QAddress.SortUrl(QAddress.UserName) = "" Then %>
		<td><%= QAddress.UserName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.UserName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.UserName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.UserName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.UserName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.passwrd.Visible Then ' passwrd %>
	<% If QAddress.SortUrl(QAddress.passwrd) = "" Then %>
		<td><%= QAddress.passwrd.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.passwrd) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.passwrd.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.passwrd.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.passwrd.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.NewCustomer.Visible Then ' NewCustomer %>
	<% If QAddress.SortUrl(QAddress.NewCustomer) = "" Then %>
		<td><%= QAddress.NewCustomer.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.NewCustomer) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.NewCustomer.FldCaption %></td><td style="width: 10px;"><% If QAddress.NewCustomer.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.NewCustomer.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.AddressID.Visible Then ' AddressID %>
	<% If QAddress.SortUrl(QAddress.AddressID) = "" Then %>
		<td><%= QAddress.AddressID.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.AddressID) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.AddressID.FldCaption %></td><td style="width: 10px;"><% If QAddress.AddressID.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.AddressID.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.Shipping2ECustomerId.Visible Then ' Shipping.CustomerId %>
	<% If QAddress.SortUrl(QAddress.Shipping2ECustomerId) = "" Then %>
		<td><%= QAddress.Shipping2ECustomerId.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.Shipping2ECustomerId) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.Shipping2ECustomerId.FldCaption %></td><td style="width: 10px;"><% If QAddress.Shipping2ECustomerId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.Shipping2ECustomerId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_FirstName.Visible Then ' ship_FirstName %>
	<% If QAddress.SortUrl(QAddress.ship_FirstName) = "" Then %>
		<td><%= QAddress.ship_FirstName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_FirstName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_FirstName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_LastName.Visible Then ' ship_LastName %>
	<% If QAddress.SortUrl(QAddress.ship_LastName) = "" Then %>
		<td><%= QAddress.ship_LastName.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_LastName) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_LastName.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_Address.Visible Then ' ship_Address %>
	<% If QAddress.SortUrl(QAddress.ship_Address) = "" Then %>
		<td><%= QAddress.ship_Address.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_Address) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_Address.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_Address.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_Address.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_City.Visible Then ' ship_City %>
	<% If QAddress.SortUrl(QAddress.ship_City) = "" Then %>
		<td><%= QAddress.ship_City.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_City) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_City.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_City.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_City.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_Province.Visible Then ' ship_Province %>
	<% If QAddress.SortUrl(QAddress.ship_Province) = "" Then %>
		<td><%= QAddress.ship_Province.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_Province) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_Province.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_Province.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_Province.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_PostalCode.Visible Then ' ship_PostalCode %>
	<% If QAddress.SortUrl(QAddress.ship_PostalCode) = "" Then %>
		<td><%= QAddress.ship_PostalCode.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_PostalCode) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_PostalCode.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_PostalCode.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_PostalCode.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_Country.Visible Then ' ship_Country %>
	<% If QAddress.SortUrl(QAddress.ship_Country) = "" Then %>
		<td><%= QAddress.ship_Country.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_Country) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_Country.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_Country.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_Country.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
	<% If QAddress.SortUrl(QAddress.ship_EmailAddress) = "" Then %>
		<td><%= QAddress.ship_EmailAddress.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_EmailAddress) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_EmailAddress.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_EmailAddress.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_EmailAddress.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.HomePhone.Visible Then ' HomePhone %>
	<% If QAddress.SortUrl(QAddress.HomePhone) = "" Then %>
		<td><%= QAddress.HomePhone.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.HomePhone) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.HomePhone.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.HomePhone.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.HomePhone.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.WorkPhone.Visible Then ' WorkPhone %>
	<% If QAddress.SortUrl(QAddress.WorkPhone) = "" Then %>
		<td><%= QAddress.WorkPhone.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.WorkPhone) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.WorkPhone.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.WorkPhone.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.WorkPhone.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If QAddress.ship_Address2.Visible Then ' ship_Address2 %>
	<% If QAddress.SortUrl(QAddress.ship_Address2) = "" Then %>
		<td><%= QAddress.ship_Address2.FldCaption %></td>
	<% Else %>
		<td><div class="ewPointer" onmousedown="ew_Sort(event,'<%= QAddress.SortUrl(QAddress.ship_Address2) %>',1);">
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= QAddress.ship_Address2.FldCaption %><%= Language.Phrase("SrchLegend") %></td><td style="width: 10px;"><% If QAddress.ship_Address2.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf QAddress.ship_Address2.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
QAddress_list.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (QAddress.ExportAll And QAddress.Export <> "") Then
	QAddress_list.StopRec = QAddress_list.TotalRecs
Else

	' Set the last record to display
	If QAddress_list.TotalRecs > QAddress_list.StartRec + QAddress_list.DisplayRecs - 1 Then
		QAddress_list.StopRec = QAddress_list.StartRec + QAddress_list.DisplayRecs - 1
	Else
		QAddress_list.StopRec = QAddress_list.TotalRecs
	End If
End If

' Move to first record
QAddress_list.RecCnt = QAddress_list.StartRec - 1
If Not QAddress_list.Recordset.Eof Then
	QAddress_list.Recordset.MoveFirst
	If QAddress_list.StartRec > 1 Then QAddress_list.Recordset.Move QAddress_list.StartRec - 1
ElseIf Not QAddress.AllowAddDeleteRow And QAddress_list.StopRec = 0 Then
	QAddress_list.StopRec = QAddress.GridAddRowCount
End If

' Initialize Aggregate
QAddress.RowType = EW_ROWTYPE_AGGREGATEINIT
Call QAddress.ResetAttrs()
Call QAddress_list.RenderRow()
QAddress_list.RowCnt = 0

' Output date rows
Do While CLng(QAddress_list.RecCnt) < CLng(QAddress_list.StopRec)
	QAddress_list.RecCnt = QAddress_list.RecCnt + 1
	If CLng(QAddress_list.RecCnt) >= CLng(QAddress_list.StartRec) Then
		QAddress_list.RowCnt = QAddress_list.RowCnt + 1

	' Set up key count
	QAddress_list.KeyCount = QAddress_list.RowIndex
	Call QAddress.ResetAttrs()
	QAddress.CssClass = ""
	If QAddress.CurrentAction = "gridadd" Then
	Else
		Call QAddress_list.LoadRowValues(QAddress_list.Recordset) ' Load row values
	End If
	QAddress.RowType = EW_ROWTYPE_VIEW ' Render view
	QAddress.RowAttrs.AddAttributes Array(Array("onmouseover", "ew_MouseOver(event, this);"), Array("onmouseout", "ew_MouseOut(event, this);"), Array("onclick", "ew_Click(event, this);"))

	' Render row
	Call QAddress_list.RenderRow()

	' Render list options
	Call QAddress_list.RenderListOptions()
%>
	<tr<%= QAddress.RowAttributes %>>
<%

' Render list options (body, left)
QAddress_list.ListOptions.Render "body", "left"
%>
	<% If QAddress.Customers2ECustomerId.Visible Then ' Customers.CustomerId %>
		<td<%= QAddress.Customers2ECustomerId.CellAttributes %>>
<div<%= QAddress.Customers2ECustomerId.ViewAttributes %>><%= QAddress.Customers2ECustomerId.ListViewValue %></div>
<a name="<%= QAddress_list.PageObjName & "_row_" & QAddress_list.RowCnt %>" id="<%= QAddress_list.PageObjName & "_row_" & QAddress_list.RowCnt %>"></a></td>
	<% End If %>
	<% If QAddress.Inv_FirstName.Visible Then ' Inv_FirstName %>
		<td<%= QAddress.Inv_FirstName.CellAttributes %>>
<div<%= QAddress.Inv_FirstName.ViewAttributes %>><%= QAddress.Inv_FirstName.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.Inv_LastName.Visible Then ' Inv_LastName %>
		<td<%= QAddress.Inv_LastName.CellAttributes %>>
<div<%= QAddress.Inv_LastName.ViewAttributes %>><%= QAddress.Inv_LastName.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.Inv_Address.Visible Then ' Inv_Address %>
		<td<%= QAddress.Inv_Address.CellAttributes %>>
<div<%= QAddress.Inv_Address.ViewAttributes %>><%= QAddress.Inv_Address.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_City.Visible Then ' inv_City %>
		<td<%= QAddress.inv_City.CellAttributes %>>
<div<%= QAddress.inv_City.ViewAttributes %>><%= QAddress.inv_City.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_Province.Visible Then ' inv_Province %>
		<td<%= QAddress.inv_Province.CellAttributes %>>
<div<%= QAddress.inv_Province.ViewAttributes %>><%= QAddress.inv_Province.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_PostalCode.Visible Then ' inv_PostalCode %>
		<td<%= QAddress.inv_PostalCode.CellAttributes %>>
<div<%= QAddress.inv_PostalCode.ViewAttributes %>><%= QAddress.inv_PostalCode.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_Country.Visible Then ' inv_Country %>
		<td<%= QAddress.inv_Country.CellAttributes %>>
<div<%= QAddress.inv_Country.ViewAttributes %>><%= QAddress.inv_Country.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_PhoneNumber.Visible Then ' inv_PhoneNumber %>
		<td<%= QAddress.inv_PhoneNumber.CellAttributes %>>
<div<%= QAddress.inv_PhoneNumber.ViewAttributes %>><%= QAddress.inv_PhoneNumber.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_EmailAddress.Visible Then ' inv_EmailAddress %>
		<td<%= QAddress.inv_EmailAddress.CellAttributes %>>
<div<%= QAddress.inv_EmailAddress.ViewAttributes %>><%= QAddress.inv_EmailAddress.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.inv_Fax.Visible Then ' inv_Fax %>
		<td<%= QAddress.inv_Fax.CellAttributes %>>
<div<%= QAddress.inv_Fax.ViewAttributes %>><%= QAddress.inv_Fax.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.Inv_Address2.Visible Then ' Inv_Address2 %>
		<td<%= QAddress.Inv_Address2.CellAttributes %>>
<div<%= QAddress.Inv_Address2.ViewAttributes %>><%= QAddress.Inv_Address2.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.UserName.Visible Then ' UserName %>
		<td<%= QAddress.UserName.CellAttributes %>>
<div<%= QAddress.UserName.ViewAttributes %>><%= QAddress.UserName.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.passwrd.Visible Then ' passwrd %>
		<td<%= QAddress.passwrd.CellAttributes %>>
<div<%= QAddress.passwrd.ViewAttributes %>><%= QAddress.passwrd.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.NewCustomer.Visible Then ' NewCustomer %>
		<td<%= QAddress.NewCustomer.CellAttributes %>>
<% If ew_ConvertToBool(QAddress.NewCustomer.CurrentValue) Then %>
<input type="checkbox" value="<%= QAddress.NewCustomer.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= QAddress.NewCustomer.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
</td>
	<% End If %>
	<% If QAddress.AddressID.Visible Then ' AddressID %>
		<td<%= QAddress.AddressID.CellAttributes %>>
<div<%= QAddress.AddressID.ViewAttributes %>><%= QAddress.AddressID.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.Shipping2ECustomerId.Visible Then ' Shipping.CustomerId %>
		<td<%= QAddress.Shipping2ECustomerId.CellAttributes %>>
<div<%= QAddress.Shipping2ECustomerId.ViewAttributes %>><%= QAddress.Shipping2ECustomerId.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_FirstName.Visible Then ' ship_FirstName %>
		<td<%= QAddress.ship_FirstName.CellAttributes %>>
<div<%= QAddress.ship_FirstName.ViewAttributes %>><%= QAddress.ship_FirstName.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_LastName.Visible Then ' ship_LastName %>
		<td<%= QAddress.ship_LastName.CellAttributes %>>
<div<%= QAddress.ship_LastName.ViewAttributes %>><%= QAddress.ship_LastName.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_Address.Visible Then ' ship_Address %>
		<td<%= QAddress.ship_Address.CellAttributes %>>
<div<%= QAddress.ship_Address.ViewAttributes %>><%= QAddress.ship_Address.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_City.Visible Then ' ship_City %>
		<td<%= QAddress.ship_City.CellAttributes %>>
<div<%= QAddress.ship_City.ViewAttributes %>><%= QAddress.ship_City.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_Province.Visible Then ' ship_Province %>
		<td<%= QAddress.ship_Province.CellAttributes %>>
<div<%= QAddress.ship_Province.ViewAttributes %>><%= QAddress.ship_Province.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_PostalCode.Visible Then ' ship_PostalCode %>
		<td<%= QAddress.ship_PostalCode.CellAttributes %>>
<div<%= QAddress.ship_PostalCode.ViewAttributes %>><%= QAddress.ship_PostalCode.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_Country.Visible Then ' ship_Country %>
		<td<%= QAddress.ship_Country.CellAttributes %>>
<div<%= QAddress.ship_Country.ViewAttributes %>><%= QAddress.ship_Country.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_EmailAddress.Visible Then ' ship_EmailAddress %>
		<td<%= QAddress.ship_EmailAddress.CellAttributes %>>
<div<%= QAddress.ship_EmailAddress.ViewAttributes %>><%= QAddress.ship_EmailAddress.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.HomePhone.Visible Then ' HomePhone %>
		<td<%= QAddress.HomePhone.CellAttributes %>>
<div<%= QAddress.HomePhone.ViewAttributes %>><%= QAddress.HomePhone.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.WorkPhone.Visible Then ' WorkPhone %>
		<td<%= QAddress.WorkPhone.CellAttributes %>>
<div<%= QAddress.WorkPhone.ViewAttributes %>><%= QAddress.WorkPhone.ListViewValue %></div>
</td>
	<% End If %>
	<% If QAddress.ship_Address2.Visible Then ' ship_Address2 %>
		<td<%= QAddress.ship_Address2.CellAttributes %>>
<div<%= QAddress.ship_Address2.ViewAttributes %>><%= QAddress.ship_Address2.ListViewValue %></div>
</td>
	<% End If %>
<%

' Render list options (body, right)
QAddress_list.ListOptions.Render "body", "right"
%>
	</tr>
<%
	End If
	If QAddress.CurrentAction <> "gridadd" Then
		QAddress_list.Recordset.MoveNext()
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
QAddress_list.Recordset.Close
Set QAddress_list.Recordset = Nothing
%>
<% If QAddress.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If QAddress.CurrentAction <> "gridadd" And QAddress.CurrentAction <> "gridedit" Then %>
<form name="ewpagerform" id="ewpagerform" class="ewForm" action="<%= ew_CurrentPage %>">
<table border="0" cellspacing="0" cellpadding="0" class="ewPager">
	<tr>
		<td>
<span class="aspmaker">
<% If Not IsObject(QAddress_list.Pager) Then Set QAddress_list.Pager = ew_NewNumericPager(QAddress_list.StartRec, QAddress_list.DisplayRecs, QAddress_list.TotalRecs, QAddress_list.RecRange) %>
<% If QAddress_list.Pager.RecordCount > 0 Then %>
	<% If QAddress_list.Pager.FirstButton.Enabled Then %>
	<a href="<%= QAddress_list.PageUrl %>start=<%= QAddress_list.Pager.FirstButton.Start %>"><b><%= Language.Phrase("PagerFirst") %></b></a>&nbsp;
	<% End If %>
	<% If QAddress_list.Pager.PrevButton.Enabled Then %>
	<a href="<%= QAddress_list.PageUrl %>start=<%= QAddress_list.Pager.PrevButton.Start %>"><b><%= Language.Phrase("PagerPrevious") %></b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In QAddress_list.Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="<%= QAddress_list.PageUrl %>start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If QAddress_list.Pager.NextButton.Enabled Then %>
	<a href="<%= QAddress_list.PageUrl %>start=<%= QAddress_list.Pager.NextButton.Start %>"><b><%= Language.Phrase("PagerNext") %></b></a>&nbsp;
	<% End If %>
	<% If QAddress_list.Pager.LastButton.Enabled Then %>
	<a href="<%= QAddress_list.PageUrl %>start=<%= QAddress_list.Pager.LastButton.Start %>"><b><%= Language.Phrase("PagerLast") %></b></a>&nbsp;
	<% End If %>
	<% If QAddress_list.Pager.ButtonCount > 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%	End If %>
	<%= Language.Phrase("Record") %>&nbsp;<%= QAddress_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= QAddress_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= QAddress_list.Pager.RecordCount %>
<% Else %>
	<% If QAddress_list.SearchWhere = "0=101" Then %>
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
</td></tr></table>
<% If QAddress.Export = "" And QAddress.CurrentAction = "" Then %>
<% End If %>
<%
QAddress_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If QAddress.Export = "" Then %>
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
Set QAddress_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cQAddress_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "QAddress"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "QAddress_list"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If QAddress.UseTokenInUrl Then PageUrl = PageUrl & "t=" & QAddress.TableVar & "&" ' add page token
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
		If QAddress.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (QAddress.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (QAddress.TableVar = Request.QueryString("t"))
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
		If IsEmpty(QAddress) Then Set QAddress = New cQAddress
		Set Table = QAddress

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		AddUrl = "QAddressadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "QAddressdelete.asp"
		MultiUpdateUrl = "QAddressupdate.asp"

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "QAddress"

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
				QAddress.GridAddRowCount = gridaddcnt
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
		Set QAddress = Nothing
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
			If QAddress.Export <> "" Or QAddress.CurrentAction = "gridadd" Or QAddress.CurrentAction = "gridedit" Then
				ListOptions.HideAllOptions()
				ExportOptions.HideAllOptions()
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session
			Call RestoreSearchParms()

			' Call Recordset SearchValidated event
			Call QAddress.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If QAddress.RecordsPerPage <> "" Then
			DisplayRecs = QAddress.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 20 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call QAddress.Recordset_Searching(SearchWhere)

		' Save search criteria
		If SearchWhere <> "" Then
			If sSrchBasic = "" Then Call ResetBasicSearchParms()
			QAddress.SearchWhere = SearchWhere ' Save to Session
			If Not RestoreSearch Then
				StartRec = 1 ' Reset start record counter
				QAddress.StartRecordNumber = StartRec
			End If
		Else
			SearchWhere = QAddress.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		QAddress.SessionWhere = sFilter
		QAddress.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, QAddress.Inv_FirstName, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.Inv_LastName, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.Inv_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_PostalCode, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_PhoneNumber, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_EmailAddress, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.Notes, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.inv_Fax, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.Inv_Address2, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.UserName, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.passwrd, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_FirstName, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_LastName, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_Address, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_City, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_Province, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_PostalCode, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_Country, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_EmailAddress, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.HomePhone, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.WorkPhone, Keyword)
			Call BuildBasicSearchSQL(sWhere, QAddress.ship_Address2, Keyword)
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
		sSearchKeyword = QAddress.BasicSearchKeyword
		sSearchType = QAddress.BasicSearchType
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
			QAddress.SessionBasicSearchKeyword = sSearchKeyword
			QAddress.SessionBasicSearchType = sSearchType
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' -----------------------------------------------------------------
	' Clear all search parameters
	'
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		QAddress.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' -----------------------------------------------------------------
	' Clear all basic search parameters
	'
	Sub ResetBasicSearchParms()

		' Clear basic search parameters
		QAddress.SessionBasicSearchKeyword = ""
		QAddress.SessionBasicSearchType = ""
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()
		Dim bRestore
		bRestore = True
		If QAddress.BasicSearchKeyword & "" <> "" Then bRestore = False
		RestoreSearch = bRestore
		If bRestore Then

			' Restore basic search values
			QAddress.BasicSearchKeyword = QAddress.SessionBasicSearchKeyword
			QAddress.BasicSearchType = QAddress.SessionBasicSearchType
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
			QAddress.CurrentOrder = Request.QueryString("order")
			QAddress.CurrentOrderType = Request.QueryString("ordertype")

			' Field Customers.CustomerId
			Call QAddress.UpdateSort(QAddress.Customers2ECustomerId)

			' Field Inv_FirstName
			Call QAddress.UpdateSort(QAddress.Inv_FirstName)

			' Field Inv_LastName
			Call QAddress.UpdateSort(QAddress.Inv_LastName)

			' Field Inv_Address
			Call QAddress.UpdateSort(QAddress.Inv_Address)

			' Field inv_City
			Call QAddress.UpdateSort(QAddress.inv_City)

			' Field inv_Province
			Call QAddress.UpdateSort(QAddress.inv_Province)

			' Field inv_PostalCode
			Call QAddress.UpdateSort(QAddress.inv_PostalCode)

			' Field inv_Country
			Call QAddress.UpdateSort(QAddress.inv_Country)

			' Field inv_PhoneNumber
			Call QAddress.UpdateSort(QAddress.inv_PhoneNumber)

			' Field inv_EmailAddress
			Call QAddress.UpdateSort(QAddress.inv_EmailAddress)

			' Field inv_Fax
			Call QAddress.UpdateSort(QAddress.inv_Fax)

			' Field Inv_Address2
			Call QAddress.UpdateSort(QAddress.Inv_Address2)

			' Field UserName
			Call QAddress.UpdateSort(QAddress.UserName)

			' Field passwrd
			Call QAddress.UpdateSort(QAddress.passwrd)

			' Field NewCustomer
			Call QAddress.UpdateSort(QAddress.NewCustomer)

			' Field AddressID
			Call QAddress.UpdateSort(QAddress.AddressID)

			' Field Shipping.CustomerId
			Call QAddress.UpdateSort(QAddress.Shipping2ECustomerId)

			' Field ship_FirstName
			Call QAddress.UpdateSort(QAddress.ship_FirstName)

			' Field ship_LastName
			Call QAddress.UpdateSort(QAddress.ship_LastName)

			' Field ship_Address
			Call QAddress.UpdateSort(QAddress.ship_Address)

			' Field ship_City
			Call QAddress.UpdateSort(QAddress.ship_City)

			' Field ship_Province
			Call QAddress.UpdateSort(QAddress.ship_Province)

			' Field ship_PostalCode
			Call QAddress.UpdateSort(QAddress.ship_PostalCode)

			' Field ship_Country
			Call QAddress.UpdateSort(QAddress.ship_Country)

			' Field ship_EmailAddress
			Call QAddress.UpdateSort(QAddress.ship_EmailAddress)

			' Field HomePhone
			Call QAddress.UpdateSort(QAddress.HomePhone)

			' Field WorkPhone
			Call QAddress.UpdateSort(QAddress.WorkPhone)

			' Field ship_Address2
			Call QAddress.UpdateSort(QAddress.ship_Address2)
			QAddress.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = QAddress.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If QAddress.SqlOrderBy <> "" Then
				sOrderBy = QAddress.SqlOrderBy
				QAddress.SessionOrderBy = sOrderBy
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
				QAddress.SessionOrderBy = sOrderBy
				QAddress.Customers2ECustomerId.Sort = ""
				QAddress.Inv_FirstName.Sort = ""
				QAddress.Inv_LastName.Sort = ""
				QAddress.Inv_Address.Sort = ""
				QAddress.inv_City.Sort = ""
				QAddress.inv_Province.Sort = ""
				QAddress.inv_PostalCode.Sort = ""
				QAddress.inv_Country.Sort = ""
				QAddress.inv_PhoneNumber.Sort = ""
				QAddress.inv_EmailAddress.Sort = ""
				QAddress.inv_Fax.Sort = ""
				QAddress.Inv_Address2.Sort = ""
				QAddress.UserName.Sort = ""
				QAddress.passwrd.Sort = ""
				QAddress.NewCustomer.Sort = ""
				QAddress.AddressID.Sort = ""
				QAddress.Shipping2ECustomerId.Sort = ""
				QAddress.ship_FirstName.Sort = ""
				QAddress.ship_LastName.Sort = ""
				QAddress.ship_Address.Sort = ""
				QAddress.ship_City.Sort = ""
				QAddress.ship_Province.Sort = ""
				QAddress.ship_PostalCode.Sort = ""
				QAddress.ship_Country.Sort = ""
				QAddress.ship_EmailAddress.Sort = ""
				QAddress.HomePhone.Sort = ""
				QAddress.WorkPhone.Sort = ""
				QAddress.ship_Address2.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			QAddress.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item
		Call ListOptions_Load()
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
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
				QAddress.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					QAddress.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = QAddress.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			QAddress.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			QAddress.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			QAddress.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		QAddress.BasicSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
		QAddress.BasicSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = QAddress.CurrentFilter
		Call QAddress.Recordset_Selecting(sFilter)
		QAddress.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = QAddress.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call QAddress.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = QAddress.KeyFilter

		' Call Row Selecting event
		Call QAddress.Row_Selecting(sFilter)

		' Load sql based on filter
		QAddress.CurrentFilter = sFilter
		sSql = QAddress.SQL
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
		Call QAddress.Row_Selected(RsRow)
		QAddress.Customers2ECustomerId.DbValue = RsRow("Customers.CustomerId")
		QAddress.Inv_FirstName.DbValue = RsRow("Inv_FirstName")
		QAddress.Inv_LastName.DbValue = RsRow("Inv_LastName")
		QAddress.Inv_Address.DbValue = RsRow("Inv_Address")
		QAddress.inv_City.DbValue = RsRow("inv_City")
		QAddress.inv_Province.DbValue = RsRow("inv_Province")
		QAddress.inv_PostalCode.DbValue = RsRow("inv_PostalCode")
		QAddress.inv_Country.DbValue = RsRow("inv_Country")
		QAddress.inv_PhoneNumber.DbValue = RsRow("inv_PhoneNumber")
		QAddress.inv_EmailAddress.DbValue = RsRow("inv_EmailAddress")
		QAddress.Notes.DbValue = RsRow("Notes")
		QAddress.inv_Fax.DbValue = RsRow("inv_Fax")
		QAddress.Inv_Address2.DbValue = RsRow("Inv_Address2")
		QAddress.UserName.DbValue = RsRow("UserName")
		QAddress.passwrd.DbValue = RsRow("passwrd")
		QAddress.NewCustomer.DbValue = ew_IIf(RsRow("NewCustomer"), "1", "0")
		QAddress.AddressID.DbValue = RsRow("AddressID")
		QAddress.Shipping2ECustomerId.DbValue = RsRow("Shipping.CustomerId")
		QAddress.ship_FirstName.DbValue = RsRow("ship_FirstName")
		QAddress.ship_LastName.DbValue = RsRow("ship_LastName")
		QAddress.ship_Address.DbValue = RsRow("ship_Address")
		QAddress.ship_City.DbValue = RsRow("ship_City")
		QAddress.ship_Province.DbValue = RsRow("ship_Province")
		QAddress.ship_PostalCode.DbValue = RsRow("ship_PostalCode")
		QAddress.ship_Country.DbValue = RsRow("ship_Country")
		QAddress.ship_EmailAddress.DbValue = RsRow("ship_EmailAddress")
		QAddress.HomePhone.DbValue = RsRow("HomePhone")
		QAddress.WorkPhone.DbValue = RsRow("WorkPhone")
		QAddress.ship_Address2.DbValue = RsRow("ship_Address2")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True

		' Load old recordset
		If bValidKey Then
			QAddress.CurrentFilter = QAddress.KeyFilter
			Dim sSql
			sSql = QAddress.SQL
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
		ViewUrl = QAddress.ViewUrl
		EditUrl = QAddress.EditUrl("")
		InlineEditUrl = QAddress.InlineEditUrl
		CopyUrl = QAddress.CopyUrl("")
		InlineCopyUrl = QAddress.InlineCopyUrl
		DeleteUrl = QAddress.DeleteUrl

		' Call Row Rendering event
		Call QAddress.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Customers.CustomerId
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
		' AddressID
		' Shipping.CustomerId
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

		If QAddress.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Customers.CustomerId
			QAddress.Customers2ECustomerId.ViewValue = QAddress.Customers2ECustomerId.CurrentValue
			QAddress.Customers2ECustomerId.ViewCustomAttributes = ""

			' Inv_FirstName
			QAddress.Inv_FirstName.ViewValue = QAddress.Inv_FirstName.CurrentValue
			QAddress.Inv_FirstName.ViewCustomAttributes = ""

			' Inv_LastName
			QAddress.Inv_LastName.ViewValue = QAddress.Inv_LastName.CurrentValue
			QAddress.Inv_LastName.ViewCustomAttributes = ""

			' Inv_Address
			QAddress.Inv_Address.ViewValue = QAddress.Inv_Address.CurrentValue
			QAddress.Inv_Address.ViewCustomAttributes = ""

			' inv_City
			QAddress.inv_City.ViewValue = QAddress.inv_City.CurrentValue
			QAddress.inv_City.ViewCustomAttributes = ""

			' inv_Province
			QAddress.inv_Province.ViewValue = QAddress.inv_Province.CurrentValue
			QAddress.inv_Province.ViewCustomAttributes = ""

			' inv_PostalCode
			QAddress.inv_PostalCode.ViewValue = QAddress.inv_PostalCode.CurrentValue
			QAddress.inv_PostalCode.ViewCustomAttributes = ""

			' inv_Country
			QAddress.inv_Country.ViewValue = QAddress.inv_Country.CurrentValue
			QAddress.inv_Country.ViewCustomAttributes = ""

			' inv_PhoneNumber
			QAddress.inv_PhoneNumber.ViewValue = QAddress.inv_PhoneNumber.CurrentValue
			QAddress.inv_PhoneNumber.ViewCustomAttributes = ""

			' inv_EmailAddress
			QAddress.inv_EmailAddress.ViewValue = QAddress.inv_EmailAddress.CurrentValue
			QAddress.inv_EmailAddress.ViewCustomAttributes = ""

			' inv_Fax
			QAddress.inv_Fax.ViewValue = QAddress.inv_Fax.CurrentValue
			QAddress.inv_Fax.ViewCustomAttributes = ""

			' Inv_Address2
			QAddress.Inv_Address2.ViewValue = QAddress.Inv_Address2.CurrentValue
			QAddress.Inv_Address2.ViewCustomAttributes = ""

			' UserName
			QAddress.UserName.ViewValue = QAddress.UserName.CurrentValue
			QAddress.UserName.ViewCustomAttributes = ""

			' passwrd
			QAddress.passwrd.ViewValue = QAddress.passwrd.CurrentValue
			QAddress.passwrd.ViewCustomAttributes = ""

			' NewCustomer
			If ew_ConvertToBool(QAddress.NewCustomer.CurrentValue) Then
				QAddress.NewCustomer.ViewValue = ew_IIf(QAddress.NewCustomer.FldTagCaption(1) <> "", QAddress.NewCustomer.FldTagCaption(1), "Yes")
			Else
				QAddress.NewCustomer.ViewValue = ew_IIf(QAddress.NewCustomer.FldTagCaption(2) <> "", QAddress.NewCustomer.FldTagCaption(2), "No")
			End If
			QAddress.NewCustomer.ViewCustomAttributes = ""

			' AddressID
			QAddress.AddressID.ViewValue = QAddress.AddressID.CurrentValue
			QAddress.AddressID.ViewCustomAttributes = ""

			' Shipping.CustomerId
			QAddress.Shipping2ECustomerId.ViewValue = QAddress.Shipping2ECustomerId.CurrentValue
			QAddress.Shipping2ECustomerId.ViewCustomAttributes = ""

			' ship_FirstName
			QAddress.ship_FirstName.ViewValue = QAddress.ship_FirstName.CurrentValue
			QAddress.ship_FirstName.ViewCustomAttributes = ""

			' ship_LastName
			QAddress.ship_LastName.ViewValue = QAddress.ship_LastName.CurrentValue
			QAddress.ship_LastName.ViewCustomAttributes = ""

			' ship_Address
			QAddress.ship_Address.ViewValue = QAddress.ship_Address.CurrentValue
			QAddress.ship_Address.ViewCustomAttributes = ""

			' ship_City
			QAddress.ship_City.ViewValue = QAddress.ship_City.CurrentValue
			QAddress.ship_City.ViewCustomAttributes = ""

			' ship_Province
			QAddress.ship_Province.ViewValue = QAddress.ship_Province.CurrentValue
			QAddress.ship_Province.ViewCustomAttributes = ""

			' ship_PostalCode
			QAddress.ship_PostalCode.ViewValue = QAddress.ship_PostalCode.CurrentValue
			QAddress.ship_PostalCode.ViewCustomAttributes = ""

			' ship_Country
			QAddress.ship_Country.ViewValue = QAddress.ship_Country.CurrentValue
			QAddress.ship_Country.ViewCustomAttributes = ""

			' ship_EmailAddress
			QAddress.ship_EmailAddress.ViewValue = QAddress.ship_EmailAddress.CurrentValue
			QAddress.ship_EmailAddress.ViewCustomAttributes = ""

			' HomePhone
			QAddress.HomePhone.ViewValue = QAddress.HomePhone.CurrentValue
			QAddress.HomePhone.ViewCustomAttributes = ""

			' WorkPhone
			QAddress.WorkPhone.ViewValue = QAddress.WorkPhone.CurrentValue
			QAddress.WorkPhone.ViewCustomAttributes = ""

			' ship_Address2
			QAddress.ship_Address2.ViewValue = QAddress.ship_Address2.CurrentValue
			QAddress.ship_Address2.ViewCustomAttributes = ""

			' View refer script
			' Customers.CustomerId

			QAddress.Customers2ECustomerId.LinkCustomAttributes = ""
			QAddress.Customers2ECustomerId.HrefValue = ""
			QAddress.Customers2ECustomerId.TooltipValue = ""

			' Inv_FirstName
			QAddress.Inv_FirstName.LinkCustomAttributes = ""
			QAddress.Inv_FirstName.HrefValue = ""
			QAddress.Inv_FirstName.TooltipValue = ""

			' Inv_LastName
			QAddress.Inv_LastName.LinkCustomAttributes = ""
			QAddress.Inv_LastName.HrefValue = ""
			QAddress.Inv_LastName.TooltipValue = ""

			' Inv_Address
			QAddress.Inv_Address.LinkCustomAttributes = ""
			QAddress.Inv_Address.HrefValue = ""
			QAddress.Inv_Address.TooltipValue = ""

			' inv_City
			QAddress.inv_City.LinkCustomAttributes = ""
			QAddress.inv_City.HrefValue = ""
			QAddress.inv_City.TooltipValue = ""

			' inv_Province
			QAddress.inv_Province.LinkCustomAttributes = ""
			QAddress.inv_Province.HrefValue = ""
			QAddress.inv_Province.TooltipValue = ""

			' inv_PostalCode
			QAddress.inv_PostalCode.LinkCustomAttributes = ""
			QAddress.inv_PostalCode.HrefValue = ""
			QAddress.inv_PostalCode.TooltipValue = ""

			' inv_Country
			QAddress.inv_Country.LinkCustomAttributes = ""
			QAddress.inv_Country.HrefValue = ""
			QAddress.inv_Country.TooltipValue = ""

			' inv_PhoneNumber
			QAddress.inv_PhoneNumber.LinkCustomAttributes = ""
			QAddress.inv_PhoneNumber.HrefValue = ""
			QAddress.inv_PhoneNumber.TooltipValue = ""

			' inv_EmailAddress
			QAddress.inv_EmailAddress.LinkCustomAttributes = ""
			QAddress.inv_EmailAddress.HrefValue = ""
			QAddress.inv_EmailAddress.TooltipValue = ""

			' inv_Fax
			QAddress.inv_Fax.LinkCustomAttributes = ""
			QAddress.inv_Fax.HrefValue = ""
			QAddress.inv_Fax.TooltipValue = ""

			' Inv_Address2
			QAddress.Inv_Address2.LinkCustomAttributes = ""
			QAddress.Inv_Address2.HrefValue = ""
			QAddress.Inv_Address2.TooltipValue = ""

			' UserName
			QAddress.UserName.LinkCustomAttributes = ""
			QAddress.UserName.HrefValue = ""
			QAddress.UserName.TooltipValue = ""

			' passwrd
			QAddress.passwrd.LinkCustomAttributes = ""
			QAddress.passwrd.HrefValue = ""
			QAddress.passwrd.TooltipValue = ""

			' NewCustomer
			QAddress.NewCustomer.LinkCustomAttributes = ""
			QAddress.NewCustomer.HrefValue = ""
			QAddress.NewCustomer.TooltipValue = ""

			' AddressID
			QAddress.AddressID.LinkCustomAttributes = ""
			QAddress.AddressID.HrefValue = ""
			QAddress.AddressID.TooltipValue = ""

			' Shipping.CustomerId
			QAddress.Shipping2ECustomerId.LinkCustomAttributes = ""
			QAddress.Shipping2ECustomerId.HrefValue = ""
			QAddress.Shipping2ECustomerId.TooltipValue = ""

			' ship_FirstName
			QAddress.ship_FirstName.LinkCustomAttributes = ""
			QAddress.ship_FirstName.HrefValue = ""
			QAddress.ship_FirstName.TooltipValue = ""

			' ship_LastName
			QAddress.ship_LastName.LinkCustomAttributes = ""
			QAddress.ship_LastName.HrefValue = ""
			QAddress.ship_LastName.TooltipValue = ""

			' ship_Address
			QAddress.ship_Address.LinkCustomAttributes = ""
			QAddress.ship_Address.HrefValue = ""
			QAddress.ship_Address.TooltipValue = ""

			' ship_City
			QAddress.ship_City.LinkCustomAttributes = ""
			QAddress.ship_City.HrefValue = ""
			QAddress.ship_City.TooltipValue = ""

			' ship_Province
			QAddress.ship_Province.LinkCustomAttributes = ""
			QAddress.ship_Province.HrefValue = ""
			QAddress.ship_Province.TooltipValue = ""

			' ship_PostalCode
			QAddress.ship_PostalCode.LinkCustomAttributes = ""
			QAddress.ship_PostalCode.HrefValue = ""
			QAddress.ship_PostalCode.TooltipValue = ""

			' ship_Country
			QAddress.ship_Country.LinkCustomAttributes = ""
			QAddress.ship_Country.HrefValue = ""
			QAddress.ship_Country.TooltipValue = ""

			' ship_EmailAddress
			QAddress.ship_EmailAddress.LinkCustomAttributes = ""
			QAddress.ship_EmailAddress.HrefValue = ""
			QAddress.ship_EmailAddress.TooltipValue = ""

			' HomePhone
			QAddress.HomePhone.LinkCustomAttributes = ""
			QAddress.HomePhone.HrefValue = ""
			QAddress.HomePhone.TooltipValue = ""

			' WorkPhone
			QAddress.WorkPhone.LinkCustomAttributes = ""
			QAddress.WorkPhone.HrefValue = ""
			QAddress.WorkPhone.TooltipValue = ""

			' ship_Address2
			QAddress.ship_Address2.LinkCustomAttributes = ""
			QAddress.ship_Address2.HrefValue = ""
			QAddress.ship_Address2.TooltipValue = ""
		End If

		' Call Row Rendered event
		If QAddress.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call QAddress.Row_Rendered()
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
