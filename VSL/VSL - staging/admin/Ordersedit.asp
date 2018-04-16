<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Ordersinfo.asp"-->
<!--#include file="OrderDetailsinfo.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="OrderDetailsgridcls.asp" -->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Orders_edit
Set Orders_edit = New cOrders_edit
Set Page = Orders_edit

' Page init processing
Call Orders_edit.Page_Init()

' Page main processing
Call Orders_edit.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Orders_edit = new ew_Page("Orders_edit");
// page properties
Orders_edit.PageID = "edit"; // page ID
Orders_edit.FormID = "fOrdersedit"; // form ID
var EW_PAGE_ID = Orders_edit.PageID; // for backward compatibility
// extend page with ValidateForm function
Orders_edit.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_InvoiceId"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.InvoiceId.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Amount"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.Amount.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_payment_gross"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.payment_gross.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_payment_fee"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.payment_fee.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Tax"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.Tax.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Shipping"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Orders.Shipping.FldErrMsg) %>");
		// Set up row object
		var row = {};
		row["index"] = infix;
		for (var j = 0; j < fobj.elements.length; j++) {
			var el = fobj.elements[j];
			var len = infix.length + 2;
			if (el.name.substr(0, len) == "x" + infix + "_") {
				var elname = "x_" + el.name.substr(len);
				if (ewLang.isObject(row[elname])) { // already exists
					if (ewLang.isArray(row[elname])) {
						row[elname][row[elname].length] = el; // add to array
					} else {
						row[elname] = [row[elname], el]; // convert to array
					}
				} else {
					row[elname] = el;
				}
			}
		}
		fobj.row = row;
		// Call Form Custom Validate event
		if (!this.Form_CustomValidate(fobj)) return false;
	}
	// Process detail page
	var detailpage = (fobj.detailpage) ? fobj.detailpage.value : "";
	if (detailpage != "") {
		return eval(detailpage+".ValidateForm(fobj)");
	}
	return true;
}
// extend page with Form_CustomValidate function
Orders_edit.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Orders_edit.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Orders_edit.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Orders_edit.ValidateRequired = false; // no JavaScript validation
<% End If %>
// multi page properties
Orders_edit.MultiPage = new ew_MultiPage();
Orders_edit.MultiPage.AddElement("x_OrderId", 1);
Orders_edit.MultiPage.AddElement("x_CustomerId", 1);
Orders_edit.MultiPage.AddElement("x_InvoiceId", 1);
Orders_edit.MultiPage.AddElement("x_Amount", 1);
Orders_edit.MultiPage.AddElement("x_Ship_FirstName", 1);
Orders_edit.MultiPage.AddElement("x_Ship_LastName", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Address", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Address2", 1);
Orders_edit.MultiPage.AddElement("x_Ship_City", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Province", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Postal", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Country", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Phone", 1);
Orders_edit.MultiPage.AddElement("x_Ship_Email", 1);
Orders_edit.MultiPage.AddElement("x_payment_status", 1);
Orders_edit.MultiPage.AddElement("x_Ordered_Date", 1);
Orders_edit.MultiPage.AddElement("x_payment_date", 1);
Orders_edit.MultiPage.AddElement("x_pfirst_name", 2);
Orders_edit.MultiPage.AddElement("x_plast_name", 2);
Orders_edit.MultiPage.AddElement("x_payer_email", 2);
Orders_edit.MultiPage.AddElement("x_txn_id", 1);
Orders_edit.MultiPage.AddElement("x_payment_gross", 2);
Orders_edit.MultiPage.AddElement("x_payment_fee", 2);
Orders_edit.MultiPage.AddElement("x_payment_type", 2);
Orders_edit.MultiPage.AddElement("x_txn_type", 2);
Orders_edit.MultiPage.AddElement("x_receiver_email", 2);
Orders_edit.MultiPage.AddElement("x_pShip_Name", 2);
Orders_edit.MultiPage.AddElement("x_pShip_Address", 2);
Orders_edit.MultiPage.AddElement("x_pShip_City", 2);
Orders_edit.MultiPage.AddElement("x_pShip_Province", 2);
Orders_edit.MultiPage.AddElement("x_pShip_Postal", 2);
Orders_edit.MultiPage.AddElement("x_pShip_Country", 2);
Orders_edit.MultiPage.AddElement("x_Tax", 1);
Orders_edit.MultiPage.AddElement("x_Shipping", 1);
Orders_edit.MultiPage.AddElement("x_EmailSent", 1);
Orders_edit.MultiPage.AddElement("x_EmailDate", 1);
Orders_edit.MultiPage.AddElement("x_PromoCodeUsed", 1);
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
<% Orders_edit.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Edit") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Orders.TableCaption %></p>
<p class="aspmaker"><a href="<%= Orders.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Orders_edit.ShowMessage %>
<form name="fOrdersedit" id="fOrdersedit" action="<%= ew_CurrentPage %>" method="post" onsubmit="return Orders_edit.ValidateForm(this);">
<p>
<input type="hidden" name="a_table" id="a_table" value="Orders">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table cellspacing="0" cellpadding="0"><tr><td>
<div id="Orders_edit" class="yui-navset">
	<ul class="yui-nav">
		<li class="selected"><a href="#tab_Orders_1"><em><span class="aspmaker"><%= Orders.PageCaption(1) %></span></em></a></li>
		<li><a href="#tab_Orders_2"><em><span class="aspmaker"><%= Orders.PageCaption(2) %></span></em></a></li>
	</ul>            
	<div class="yui-content">
		<div id="tab_Orders_1">
<table cellspacing="0" class="ewGrid" style="width: 100%"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Orders.OrderId.Visible Then ' OrderId %>
	<tr id="r_OrderId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.OrderId.FldCaption %></td>
		<td<%= Orders.OrderId.CellAttributes %>><span id="el_OrderId">
<div<%= Orders.OrderId.ViewAttributes %>><%= Orders.OrderId.EditValue %></div>
<input type="hidden" name="x_OrderId" id="x_OrderId" value="<%= Server.HTMLEncode(Orders.OrderId.CurrentValue&"") %>">
</span><%= Orders.OrderId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.CustomerId.Visible Then ' CustomerId %>
	<tr id="r_CustomerId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.CustomerId.FldCaption %></td>
		<td<%= Orders.CustomerId.CellAttributes %>><span id="el_CustomerId">
<% If Orders.CustomerId.SessionValue <> "" Then %>
<div<%= Orders.CustomerId.ViewAttributes %>>
<% If Orders.CustomerId.LinkAttributes <> "" Then %>
<a<%= Orders.CustomerId.LinkAttributes %>><%= Orders.CustomerId.ViewValue %></a>
<% Else %>
<%= Orders.CustomerId.ViewValue %>
<% End If %>
</div>
<input type="hidden" id="x_CustomerId" name="x_CustomerId" value="<%= ew_HtmlEncode(Orders.CustomerId.CurrentValue) %>">
<% Else %>
<select id="x_CustomerId" name="x_CustomerId"<%= Orders.CustomerId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Orders.CustomerId.EditValue) Then
	arwrk = Orders.CustomerId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Orders.CustomerId.CurrentValue&"" Then
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
</span><%= Orders.CustomerId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.InvoiceId.Visible Then ' InvoiceId %>
	<tr id="r_InvoiceId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.InvoiceId.FldCaption %></td>
		<td<%= Orders.InvoiceId.CellAttributes %>><span id="el_InvoiceId">
<input type="text" name="x_InvoiceId" id="x_InvoiceId" size="30" value="<%= Orders.InvoiceId.EditValue %>"<%= Orders.InvoiceId.EditAttributes %>>
</span><%= Orders.InvoiceId.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Amount.Visible Then ' Amount %>
	<tr id="r_Amount"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Amount.FldCaption %></td>
		<td<%= Orders.Amount.CellAttributes %>><span id="el_Amount">
<input type="text" name="x_Amount" id="x_Amount" size="30" value="<%= Orders.Amount.EditValue %>"<%= Orders.Amount.EditAttributes %>>
</span><%= Orders.Amount.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_FirstName.Visible Then ' Ship_FirstName %>
	<tr id="r_Ship_FirstName"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_FirstName.FldCaption %></td>
		<td<%= Orders.Ship_FirstName.CellAttributes %>><span id="el_Ship_FirstName">
<input type="text" name="x_Ship_FirstName" id="x_Ship_FirstName" size="30" maxlength="255" value="<%= Orders.Ship_FirstName.EditValue %>"<%= Orders.Ship_FirstName.EditAttributes %>>
</span><%= Orders.Ship_FirstName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_LastName.Visible Then ' Ship_LastName %>
	<tr id="r_Ship_LastName"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_LastName.FldCaption %></td>
		<td<%= Orders.Ship_LastName.CellAttributes %>><span id="el_Ship_LastName">
<input type="text" name="x_Ship_LastName" id="x_Ship_LastName" size="30" maxlength="255" value="<%= Orders.Ship_LastName.EditValue %>"<%= Orders.Ship_LastName.EditAttributes %>>
</span><%= Orders.Ship_LastName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Address.Visible Then ' Ship_Address %>
	<tr id="r_Ship_Address"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Address.FldCaption %></td>
		<td<%= Orders.Ship_Address.CellAttributes %>><span id="el_Ship_Address">
<input type="text" name="x_Ship_Address" id="x_Ship_Address" size="30" maxlength="50" value="<%= Orders.Ship_Address.EditValue %>"<%= Orders.Ship_Address.EditAttributes %>>
</span><%= Orders.Ship_Address.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Address2.Visible Then ' Ship_Address2 %>
	<tr id="r_Ship_Address2"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Address2.FldCaption %></td>
		<td<%= Orders.Ship_Address2.CellAttributes %>><span id="el_Ship_Address2">
<input type="text" name="x_Ship_Address2" id="x_Ship_Address2" size="30" maxlength="255" value="<%= Orders.Ship_Address2.EditValue %>"<%= Orders.Ship_Address2.EditAttributes %>>
</span><%= Orders.Ship_Address2.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_City.Visible Then ' Ship_City %>
	<tr id="r_Ship_City"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_City.FldCaption %></td>
		<td<%= Orders.Ship_City.CellAttributes %>><span id="el_Ship_City">
<input type="text" name="x_Ship_City" id="x_Ship_City" size="30" maxlength="255" value="<%= Orders.Ship_City.EditValue %>"<%= Orders.Ship_City.EditAttributes %>>
</span><%= Orders.Ship_City.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Province.Visible Then ' Ship_Province %>
	<tr id="r_Ship_Province"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Province.FldCaption %></td>
		<td<%= Orders.Ship_Province.CellAttributes %>><span id="el_Ship_Province">
<input type="text" name="x_Ship_Province" id="x_Ship_Province" size="30" maxlength="255" value="<%= Orders.Ship_Province.EditValue %>"<%= Orders.Ship_Province.EditAttributes %>>
</span><%= Orders.Ship_Province.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Postal.Visible Then ' Ship_Postal %>
	<tr id="r_Ship_Postal"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Postal.FldCaption %></td>
		<td<%= Orders.Ship_Postal.CellAttributes %>><span id="el_Ship_Postal">
<input type="text" name="x_Ship_Postal" id="x_Ship_Postal" size="30" maxlength="255" value="<%= Orders.Ship_Postal.EditValue %>"<%= Orders.Ship_Postal.EditAttributes %>>
</span><%= Orders.Ship_Postal.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Country.Visible Then ' Ship_Country %>
	<tr id="r_Ship_Country"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Country.FldCaption %></td>
		<td<%= Orders.Ship_Country.CellAttributes %>><span id="el_Ship_Country">
<input type="text" name="x_Ship_Country" id="x_Ship_Country" size="30" maxlength="255" value="<%= Orders.Ship_Country.EditValue %>"<%= Orders.Ship_Country.EditAttributes %>>
</span><%= Orders.Ship_Country.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Phone.Visible Then ' Ship_Phone %>
	<tr id="r_Ship_Phone"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Phone.FldCaption %></td>
		<td<%= Orders.Ship_Phone.CellAttributes %>><span id="el_Ship_Phone">
<input type="text" name="x_Ship_Phone" id="x_Ship_Phone" size="30" maxlength="255" value="<%= Orders.Ship_Phone.EditValue %>"<%= Orders.Ship_Phone.EditAttributes %>>
</span><%= Orders.Ship_Phone.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ship_Email.Visible Then ' Ship_Email %>
	<tr id="r_Ship_Email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Email.FldCaption %></td>
		<td<%= Orders.Ship_Email.CellAttributes %>><span id="el_Ship_Email">
<input type="text" name="x_Ship_Email" id="x_Ship_Email" size="30" maxlength="255" value="<%= Orders.Ship_Email.EditValue %>"<%= Orders.Ship_Email.EditAttributes %>>
</span><%= Orders.Ship_Email.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payment_status.Visible Then ' payment_status %>
	<tr id="r_payment_status"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_status.FldCaption %></td>
		<td<%= Orders.payment_status.CellAttributes %>><span id="el_payment_status">
<select id="x_payment_status" name="x_payment_status"<%= Orders.payment_status.EditAttributes %>>
<%
emptywrk = True
If IsArray(Orders.payment_status.EditValue) Then
	arwrk = Orders.payment_status.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Orders.payment_status.CurrentValue&"" Then
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
</span><%= Orders.payment_status.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Ordered_Date.Visible Then ' Ordered_Date %>
	<tr id="r_Ordered_Date"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ordered_Date.FldCaption %></td>
		<td<%= Orders.Ordered_Date.CellAttributes %>><span id="el_Ordered_Date">
<input type="text" name="x_Ordered_Date" id="x_Ordered_Date" value="<%= Orders.Ordered_Date.EditValue %>"<%= Orders.Ordered_Date.EditAttributes %>>
</span><%= Orders.Ordered_Date.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payment_date.Visible Then ' payment_date %>
	<tr id="r_payment_date"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_date.FldCaption %></td>
		<td<%= Orders.payment_date.CellAttributes %>><span id="el_payment_date">
<input type="text" name="x_payment_date" id="x_payment_date" value="<%= Orders.payment_date.EditValue %>"<%= Orders.payment_date.EditAttributes %>>
</span><%= Orders.payment_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.txn_id.Visible Then ' txn_id %>
	<tr id="r_txn_id"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.txn_id.FldCaption %></td>
		<td<%= Orders.txn_id.CellAttributes %>><span id="el_txn_id">
<input type="text" name="x_txn_id" id="x_txn_id" size="30" maxlength="255" value="<%= Orders.txn_id.EditValue %>"<%= Orders.txn_id.EditAttributes %>>
</span><%= Orders.txn_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Tax.Visible Then ' Tax %>
	<tr id="r_Tax"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Tax.FldCaption %></td>
		<td<%= Orders.Tax.CellAttributes %>><span id="el_Tax">
<input type="text" name="x_Tax" id="x_Tax" size="30" value="<%= Orders.Tax.EditValue %>"<%= Orders.Tax.EditAttributes %>>
</span><%= Orders.Tax.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.Shipping.Visible Then ' Shipping %>
	<tr id="r_Shipping"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Shipping.FldCaption %></td>
		<td<%= Orders.Shipping.CellAttributes %>><span id="el_Shipping">
<input type="text" name="x_Shipping" id="x_Shipping" size="30" value="<%= Orders.Shipping.EditValue %>"<%= Orders.Shipping.EditAttributes %>>
</span><%= Orders.Shipping.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.EmailSent.Visible Then ' EmailSent %>
	<tr id="r_EmailSent"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.EmailSent.FldCaption %></td>
		<td<%= Orders.EmailSent.CellAttributes %>><span id="el_EmailSent">
<select id="x_EmailSent" name="x_EmailSent"<%= Orders.EmailSent.EditAttributes %>>
<%
emptywrk = True
If IsArray(Orders.EmailSent.EditValue) Then
	arwrk = Orders.EmailSent.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Orders.EmailSent.CurrentValue&"" Then
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
</span><%= Orders.EmailSent.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.EmailDate.Visible Then ' EmailDate %>
	<tr id="r_EmailDate"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.EmailDate.FldCaption %></td>
		<td<%= Orders.EmailDate.CellAttributes %>><span id="el_EmailDate">
<input type="text" name="x_EmailDate" id="x_EmailDate" value="<%= Orders.EmailDate.EditValue %>"<%= Orders.EmailDate.EditAttributes %>>
</span><%= Orders.EmailDate.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.PromoCodeUsed.Visible Then ' PromoCodeUsed %>
	<tr id="r_PromoCodeUsed"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.PromoCodeUsed.FldCaption %></td>
		<td<%= Orders.PromoCodeUsed.CellAttributes %>><span id="el_PromoCodeUsed">
<input type="text" name="x_PromoCodeUsed" id="x_PromoCodeUsed" size="30" maxlength="6" value="<%= Orders.PromoCodeUsed.EditValue %>"<%= Orders.PromoCodeUsed.EditAttributes %>>
</span><%= Orders.PromoCodeUsed.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
		</div>
		<div id="tab_Orders_2">
<table cellspacing="0" class="ewGrid" style="width: 100%"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Orders.pfirst_name.Visible Then ' pfirst_name %>
	<tr id="r_pfirst_name"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pfirst_name.FldCaption %></td>
		<td<%= Orders.pfirst_name.CellAttributes %>><span id="el_pfirst_name">
<input type="text" name="x_pfirst_name" id="x_pfirst_name" size="30" maxlength="255" value="<%= Orders.pfirst_name.EditValue %>"<%= Orders.pfirst_name.EditAttributes %>>
</span><%= Orders.pfirst_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.plast_name.Visible Then ' plast_name %>
	<tr id="r_plast_name"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.plast_name.FldCaption %></td>
		<td<%= Orders.plast_name.CellAttributes %>><span id="el_plast_name">
<input type="text" name="x_plast_name" id="x_plast_name" size="30" maxlength="255" value="<%= Orders.plast_name.EditValue %>"<%= Orders.plast_name.EditAttributes %>>
</span><%= Orders.plast_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payer_email.Visible Then ' payer_email %>
	<tr id="r_payer_email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payer_email.FldCaption %></td>
		<td<%= Orders.payer_email.CellAttributes %>><span id="el_payer_email">
<input type="text" name="x_payer_email" id="x_payer_email" size="30" maxlength="255" value="<%= Orders.payer_email.EditValue %>"<%= Orders.payer_email.EditAttributes %>>
</span><%= Orders.payer_email.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payment_gross.Visible Then ' payment_gross %>
	<tr id="r_payment_gross"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_gross.FldCaption %></td>
		<td<%= Orders.payment_gross.CellAttributes %>><span id="el_payment_gross">
<input type="text" name="x_payment_gross" id="x_payment_gross" size="30" value="<%= Orders.payment_gross.EditValue %>"<%= Orders.payment_gross.EditAttributes %>>
</span><%= Orders.payment_gross.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payment_fee.Visible Then ' payment_fee %>
	<tr id="r_payment_fee"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_fee.FldCaption %></td>
		<td<%= Orders.payment_fee.CellAttributes %>><span id="el_payment_fee">
<input type="text" name="x_payment_fee" id="x_payment_fee" size="30" value="<%= Orders.payment_fee.EditValue %>"<%= Orders.payment_fee.EditAttributes %>>
</span><%= Orders.payment_fee.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.payment_type.Visible Then ' payment_type %>
	<tr id="r_payment_type"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_type.FldCaption %></td>
		<td<%= Orders.payment_type.CellAttributes %>><span id="el_payment_type">
<input type="text" name="x_payment_type" id="x_payment_type" size="30" maxlength="255" value="<%= Orders.payment_type.EditValue %>"<%= Orders.payment_type.EditAttributes %>>
</span><%= Orders.payment_type.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.txn_type.Visible Then ' txn_type %>
	<tr id="r_txn_type"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.txn_type.FldCaption %></td>
		<td<%= Orders.txn_type.CellAttributes %>><span id="el_txn_type">
<input type="text" name="x_txn_type" id="x_txn_type" size="30" maxlength="255" value="<%= Orders.txn_type.EditValue %>"<%= Orders.txn_type.EditAttributes %>>
</span><%= Orders.txn_type.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.receiver_email.Visible Then ' receiver_email %>
	<tr id="r_receiver_email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.receiver_email.FldCaption %></td>
		<td<%= Orders.receiver_email.CellAttributes %>><span id="el_receiver_email">
<input type="text" name="x_receiver_email" id="x_receiver_email" size="30" maxlength="255" value="<%= Orders.receiver_email.EditValue %>"<%= Orders.receiver_email.EditAttributes %>>
</span><%= Orders.receiver_email.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_Name.Visible Then ' pShip_Name %>
	<tr id="r_pShip_Name"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Name.FldCaption %></td>
		<td<%= Orders.pShip_Name.CellAttributes %>><span id="el_pShip_Name">
<input type="text" name="x_pShip_Name" id="x_pShip_Name" size="30" maxlength="255" value="<%= Orders.pShip_Name.EditValue %>"<%= Orders.pShip_Name.EditAttributes %>>
</span><%= Orders.pShip_Name.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_Address.Visible Then ' pShip_Address %>
	<tr id="r_pShip_Address"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Address.FldCaption %></td>
		<td<%= Orders.pShip_Address.CellAttributes %>><span id="el_pShip_Address">
<input type="text" name="x_pShip_Address" id="x_pShip_Address" size="30" maxlength="255" value="<%= Orders.pShip_Address.EditValue %>"<%= Orders.pShip_Address.EditAttributes %>>
</span><%= Orders.pShip_Address.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_City.Visible Then ' pShip_City %>
	<tr id="r_pShip_City"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_City.FldCaption %></td>
		<td<%= Orders.pShip_City.CellAttributes %>><span id="el_pShip_City">
<input type="text" name="x_pShip_City" id="x_pShip_City" size="30" maxlength="255" value="<%= Orders.pShip_City.EditValue %>"<%= Orders.pShip_City.EditAttributes %>>
</span><%= Orders.pShip_City.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_Province.Visible Then ' pShip_Province %>
	<tr id="r_pShip_Province"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Province.FldCaption %></td>
		<td<%= Orders.pShip_Province.CellAttributes %>><span id="el_pShip_Province">
<input type="text" name="x_pShip_Province" id="x_pShip_Province" size="30" maxlength="255" value="<%= Orders.pShip_Province.EditValue %>"<%= Orders.pShip_Province.EditAttributes %>>
</span><%= Orders.pShip_Province.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_Postal.Visible Then ' pShip_Postal %>
	<tr id="r_pShip_Postal"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Postal.FldCaption %></td>
		<td<%= Orders.pShip_Postal.CellAttributes %>><span id="el_pShip_Postal">
<input type="text" name="x_pShip_Postal" id="x_pShip_Postal" size="30" maxlength="255" value="<%= Orders.pShip_Postal.EditValue %>"<%= Orders.pShip_Postal.EditAttributes %>>
</span><%= Orders.pShip_Postal.CustomMsg %></td>
	</tr>
<% End If %>
<% If Orders.pShip_Country.Visible Then ' pShip_Country %>
	<tr id="r_pShip_Country"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Country.FldCaption %></td>
		<td<%= Orders.pShip_Country.CellAttributes %>><span id="el_pShip_Country">
<input type="text" name="x_pShip_Country" id="x_pShip_Country" size="30" maxlength="255" value="<%= Orders.pShip_Country.EditValue %>"<%= Orders.pShip_Country.EditAttributes %>>
</span><%= Orders.pShip_Country.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
		</div>
	</div>
</div>
</td></tr></table>
<script type="text/javascript">
<!--
ew_TabView(Orders_edit);
//-->
</script>	
<p>
<% If Orders.CurrentDetailTable = "OrderDetails" And OrderDetails.DetailEdit Then %>
<br>
<!--#include file="OrderDetailsgrid.asp" -->
<br>
<% End If %>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("EditBtn")) %>">
</form>
<%
Orders_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Orders_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrders_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Orders"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Orders_edit"
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
		' Initialize other table object

		If IsEmpty(OrderDetails) Then Set OrderDetails = New cOrderDetails

		' Initialize other table object
		If IsEmpty(Customers) Then Set Customers = New cCustomers

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Orders"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
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

	' Create form object
	Set ObjForm = New cFormObj

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Load key from QueryString
		If Request.QueryString("OrderId").Count > 0 Then
			Orders.OrderId.QueryStringValue = Request.QueryString("OrderId")
		End If

		' Set up master detail parameters
		SetUpMasterParms()
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			Orders.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values

			' Set up detail parameters
			SetUpDetailParms()

			' Validate Form
			If Not ValidateForm() Then
				Orders.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				Orders.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		Else
			Orders.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If Orders.OrderId.CurrentValue = "" Then Call Page_Terminate("Orderslist.asp") ' Invalid key, return to list

		' Set up detail parameters
		SetUpDetailParms()
		Select Case Orders.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Orderslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				Orders.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					Dim sReturnUrl
					If Orders.CurrentDetailTable <> "" Then ' Master/detail edit
						sReturnUrl = Orders.DetailUrl
					Else
						sReturnUrl = Orders.ReturnUrl
					End If
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					Orders.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		Orders.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call Orders.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
		Dim index, confirmPage
		index = ObjForm.Index ' Save form index
		ObjForm.Index = 0
		confirmPage = (ObjForm.GetValue("a_confirm") & "" <> "")
		ObjForm.Index = index ' Restore form index
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not Orders.OrderId.FldIsDetailKey Then Orders.OrderId.FormValue = ObjForm.GetValue("x_OrderId")
		If Not Orders.CustomerId.FldIsDetailKey Then Orders.CustomerId.FormValue = ObjForm.GetValue("x_CustomerId")
		If Not Orders.InvoiceId.FldIsDetailKey Then Orders.InvoiceId.FormValue = ObjForm.GetValue("x_InvoiceId")
		If Not Orders.Amount.FldIsDetailKey Then Orders.Amount.FormValue = ObjForm.GetValue("x_Amount")
		If Not Orders.Ship_FirstName.FldIsDetailKey Then Orders.Ship_FirstName.FormValue = ObjForm.GetValue("x_Ship_FirstName")
		If Not Orders.Ship_LastName.FldIsDetailKey Then Orders.Ship_LastName.FormValue = ObjForm.GetValue("x_Ship_LastName")
		If Not Orders.Ship_Address.FldIsDetailKey Then Orders.Ship_Address.FormValue = ObjForm.GetValue("x_Ship_Address")
		If Not Orders.Ship_Address2.FldIsDetailKey Then Orders.Ship_Address2.FormValue = ObjForm.GetValue("x_Ship_Address2")
		If Not Orders.Ship_City.FldIsDetailKey Then Orders.Ship_City.FormValue = ObjForm.GetValue("x_Ship_City")
		If Not Orders.Ship_Province.FldIsDetailKey Then Orders.Ship_Province.FormValue = ObjForm.GetValue("x_Ship_Province")
		If Not Orders.Ship_Postal.FldIsDetailKey Then Orders.Ship_Postal.FormValue = ObjForm.GetValue("x_Ship_Postal")
		If Not Orders.Ship_Country.FldIsDetailKey Then Orders.Ship_Country.FormValue = ObjForm.GetValue("x_Ship_Country")
		If Not Orders.Ship_Phone.FldIsDetailKey Then Orders.Ship_Phone.FormValue = ObjForm.GetValue("x_Ship_Phone")
		If Not Orders.Ship_Email.FldIsDetailKey Then Orders.Ship_Email.FormValue = ObjForm.GetValue("x_Ship_Email")
		If Not Orders.payment_status.FldIsDetailKey Then Orders.payment_status.FormValue = ObjForm.GetValue("x_payment_status")
		If Not Orders.Ordered_Date.FldIsDetailKey Then Orders.Ordered_Date.FormValue = ObjForm.GetValue("x_Ordered_Date")
		If Not Orders.Ordered_Date.FldIsDetailKey Then Orders.Ordered_Date.CurrentValue = ew_UnFormatDateTime(Orders.Ordered_Date.CurrentValue, 8)
		If Not Orders.payment_date.FldIsDetailKey Then Orders.payment_date.FormValue = ObjForm.GetValue("x_payment_date")
		If Not Orders.payment_date.FldIsDetailKey Then Orders.payment_date.CurrentValue = ew_UnFormatDateTime(Orders.payment_date.CurrentValue, 8)
		If Not Orders.pfirst_name.FldIsDetailKey Then Orders.pfirst_name.FormValue = ObjForm.GetValue("x_pfirst_name")
		If Not Orders.plast_name.FldIsDetailKey Then Orders.plast_name.FormValue = ObjForm.GetValue("x_plast_name")
		If Not Orders.payer_email.FldIsDetailKey Then Orders.payer_email.FormValue = ObjForm.GetValue("x_payer_email")
		If Not Orders.txn_id.FldIsDetailKey Then Orders.txn_id.FormValue = ObjForm.GetValue("x_txn_id")
		If Not Orders.payment_gross.FldIsDetailKey Then Orders.payment_gross.FormValue = ObjForm.GetValue("x_payment_gross")
		If Not Orders.payment_fee.FldIsDetailKey Then Orders.payment_fee.FormValue = ObjForm.GetValue("x_payment_fee")
		If Not Orders.payment_type.FldIsDetailKey Then Orders.payment_type.FormValue = ObjForm.GetValue("x_payment_type")
		If Not Orders.txn_type.FldIsDetailKey Then Orders.txn_type.FormValue = ObjForm.GetValue("x_txn_type")
		If Not Orders.receiver_email.FldIsDetailKey Then Orders.receiver_email.FormValue = ObjForm.GetValue("x_receiver_email")
		If Not Orders.pShip_Name.FldIsDetailKey Then Orders.pShip_Name.FormValue = ObjForm.GetValue("x_pShip_Name")
		If Not Orders.pShip_Address.FldIsDetailKey Then Orders.pShip_Address.FormValue = ObjForm.GetValue("x_pShip_Address")
		If Not Orders.pShip_City.FldIsDetailKey Then Orders.pShip_City.FormValue = ObjForm.GetValue("x_pShip_City")
		If Not Orders.pShip_Province.FldIsDetailKey Then Orders.pShip_Province.FormValue = ObjForm.GetValue("x_pShip_Province")
		If Not Orders.pShip_Postal.FldIsDetailKey Then Orders.pShip_Postal.FormValue = ObjForm.GetValue("x_pShip_Postal")
		If Not Orders.pShip_Country.FldIsDetailKey Then Orders.pShip_Country.FormValue = ObjForm.GetValue("x_pShip_Country")
		If Not Orders.Tax.FldIsDetailKey Then Orders.Tax.FormValue = ObjForm.GetValue("x_Tax")
		If Not Orders.Shipping.FldIsDetailKey Then Orders.Shipping.FormValue = ObjForm.GetValue("x_Shipping")
		If Not Orders.EmailSent.FldIsDetailKey Then Orders.EmailSent.FormValue = ObjForm.GetValue("x_EmailSent")
		If Not Orders.EmailDate.FldIsDetailKey Then Orders.EmailDate.FormValue = ObjForm.GetValue("x_EmailDate")
		If Not Orders.EmailDate.FldIsDetailKey Then Orders.EmailDate.CurrentValue = ew_UnFormatDateTime(Orders.EmailDate.CurrentValue, 8)
		If Not Orders.PromoCodeUsed.FldIsDetailKey Then Orders.PromoCodeUsed.FormValue = ObjForm.GetValue("x_PromoCodeUsed")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		Orders.OrderId.CurrentValue = Orders.OrderId.FormValue
		Orders.CustomerId.CurrentValue = Orders.CustomerId.FormValue
		Orders.InvoiceId.CurrentValue = Orders.InvoiceId.FormValue
		Orders.Amount.CurrentValue = Orders.Amount.FormValue
		Orders.Ship_FirstName.CurrentValue = Orders.Ship_FirstName.FormValue
		Orders.Ship_LastName.CurrentValue = Orders.Ship_LastName.FormValue
		Orders.Ship_Address.CurrentValue = Orders.Ship_Address.FormValue
		Orders.Ship_Address2.CurrentValue = Orders.Ship_Address2.FormValue
		Orders.Ship_City.CurrentValue = Orders.Ship_City.FormValue
		Orders.Ship_Province.CurrentValue = Orders.Ship_Province.FormValue
		Orders.Ship_Postal.CurrentValue = Orders.Ship_Postal.FormValue
		Orders.Ship_Country.CurrentValue = Orders.Ship_Country.FormValue
		Orders.Ship_Phone.CurrentValue = Orders.Ship_Phone.FormValue
		Orders.Ship_Email.CurrentValue = Orders.Ship_Email.FormValue
		Orders.payment_status.CurrentValue = Orders.payment_status.FormValue
		Orders.Ordered_Date.CurrentValue = Orders.Ordered_Date.FormValue
		Orders.Ordered_Date.CurrentValue = ew_UnFormatDateTime(Orders.Ordered_Date.CurrentValue, 8)
		Orders.payment_date.CurrentValue = Orders.payment_date.FormValue
		Orders.payment_date.CurrentValue = ew_UnFormatDateTime(Orders.payment_date.CurrentValue, 8)
		Orders.pfirst_name.CurrentValue = Orders.pfirst_name.FormValue
		Orders.plast_name.CurrentValue = Orders.plast_name.FormValue
		Orders.payer_email.CurrentValue = Orders.payer_email.FormValue
		Orders.txn_id.CurrentValue = Orders.txn_id.FormValue
		Orders.payment_gross.CurrentValue = Orders.payment_gross.FormValue
		Orders.payment_fee.CurrentValue = Orders.payment_fee.FormValue
		Orders.payment_type.CurrentValue = Orders.payment_type.FormValue
		Orders.txn_type.CurrentValue = Orders.txn_type.FormValue
		Orders.receiver_email.CurrentValue = Orders.receiver_email.FormValue
		Orders.pShip_Name.CurrentValue = Orders.pShip_Name.FormValue
		Orders.pShip_Address.CurrentValue = Orders.pShip_Address.FormValue
		Orders.pShip_City.CurrentValue = Orders.pShip_City.FormValue
		Orders.pShip_Province.CurrentValue = Orders.pShip_Province.FormValue
		Orders.pShip_Postal.CurrentValue = Orders.pShip_Postal.FormValue
		Orders.pShip_Country.CurrentValue = Orders.pShip_Country.FormValue
		Orders.Tax.CurrentValue = Orders.Tax.FormValue
		Orders.Shipping.CurrentValue = Orders.Shipping.FormValue
		Orders.EmailSent.CurrentValue = Orders.EmailSent.FormValue
		Orders.EmailDate.CurrentValue = Orders.EmailDate.FormValue
		Orders.EmailDate.CurrentValue = ew_UnFormatDateTime(Orders.EmailDate.CurrentValue, 8)
		Orders.PromoCodeUsed.CurrentValue = Orders.PromoCodeUsed.FormValue
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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
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

			' InvoiceId
			Orders.InvoiceId.LinkCustomAttributes = ""
			Orders.InvoiceId.HrefValue = ""
			Orders.InvoiceId.TooltipValue = ""

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

			' Ship_Address
			Orders.Ship_Address.LinkCustomAttributes = ""
			Orders.Ship_Address.HrefValue = ""
			Orders.Ship_Address.TooltipValue = ""

			' Ship_Address2
			Orders.Ship_Address2.LinkCustomAttributes = ""
			Orders.Ship_Address2.HrefValue = ""
			Orders.Ship_Address2.TooltipValue = ""

			' Ship_City
			Orders.Ship_City.LinkCustomAttributes = ""
			Orders.Ship_City.HrefValue = ""
			Orders.Ship_City.TooltipValue = ""

			' Ship_Province
			Orders.Ship_Province.LinkCustomAttributes = ""
			Orders.Ship_Province.HrefValue = ""
			Orders.Ship_Province.TooltipValue = ""

			' Ship_Postal
			Orders.Ship_Postal.LinkCustomAttributes = ""
			Orders.Ship_Postal.HrefValue = ""
			Orders.Ship_Postal.TooltipValue = ""

			' Ship_Country
			Orders.Ship_Country.LinkCustomAttributes = ""
			Orders.Ship_Country.HrefValue = ""
			Orders.Ship_Country.TooltipValue = ""

			' Ship_Phone
			Orders.Ship_Phone.LinkCustomAttributes = ""
			Orders.Ship_Phone.HrefValue = ""
			Orders.Ship_Phone.TooltipValue = ""

			' Ship_Email
			Orders.Ship_Email.LinkCustomAttributes = ""
			Orders.Ship_Email.HrefValue = ""
			Orders.Ship_Email.TooltipValue = ""

			' payment_status
			Orders.payment_status.LinkCustomAttributes = ""
			Orders.payment_status.HrefValue = ""
			Orders.payment_status.TooltipValue = ""

			' Ordered_Date
			Orders.Ordered_Date.LinkCustomAttributes = ""
			Orders.Ordered_Date.HrefValue = ""
			Orders.Ordered_Date.TooltipValue = ""

			' payment_date
			Orders.payment_date.LinkCustomAttributes = ""
			Orders.payment_date.HrefValue = ""
			Orders.payment_date.TooltipValue = ""

			' pfirst_name
			Orders.pfirst_name.LinkCustomAttributes = ""
			Orders.pfirst_name.HrefValue = ""
			Orders.pfirst_name.TooltipValue = ""

			' plast_name
			Orders.plast_name.LinkCustomAttributes = ""
			Orders.plast_name.HrefValue = ""
			Orders.plast_name.TooltipValue = ""

			' payer_email
			Orders.payer_email.LinkCustomAttributes = ""
			Orders.payer_email.HrefValue = ""
			Orders.payer_email.TooltipValue = ""

			' txn_id
			Orders.txn_id.LinkCustomAttributes = ""
			Orders.txn_id.HrefValue = ""
			Orders.txn_id.TooltipValue = ""

			' payment_gross
			Orders.payment_gross.LinkCustomAttributes = ""
			Orders.payment_gross.HrefValue = ""
			Orders.payment_gross.TooltipValue = ""

			' payment_fee
			Orders.payment_fee.LinkCustomAttributes = ""
			Orders.payment_fee.HrefValue = ""
			Orders.payment_fee.TooltipValue = ""

			' payment_type
			Orders.payment_type.LinkCustomAttributes = ""
			Orders.payment_type.HrefValue = ""
			Orders.payment_type.TooltipValue = ""

			' txn_type
			Orders.txn_type.LinkCustomAttributes = ""
			Orders.txn_type.HrefValue = ""
			Orders.txn_type.TooltipValue = ""

			' receiver_email
			Orders.receiver_email.LinkCustomAttributes = ""
			Orders.receiver_email.HrefValue = ""
			Orders.receiver_email.TooltipValue = ""

			' pShip_Name
			Orders.pShip_Name.LinkCustomAttributes = ""
			Orders.pShip_Name.HrefValue = ""
			Orders.pShip_Name.TooltipValue = ""

			' pShip_Address
			Orders.pShip_Address.LinkCustomAttributes = ""
			Orders.pShip_Address.HrefValue = ""
			Orders.pShip_Address.TooltipValue = ""

			' pShip_City
			Orders.pShip_City.LinkCustomAttributes = ""
			Orders.pShip_City.HrefValue = ""
			Orders.pShip_City.TooltipValue = ""

			' pShip_Province
			Orders.pShip_Province.LinkCustomAttributes = ""
			Orders.pShip_Province.HrefValue = ""
			Orders.pShip_Province.TooltipValue = ""

			' pShip_Postal
			Orders.pShip_Postal.LinkCustomAttributes = ""
			Orders.pShip_Postal.HrefValue = ""
			Orders.pShip_Postal.TooltipValue = ""

			' pShip_Country
			Orders.pShip_Country.LinkCustomAttributes = ""
			Orders.pShip_Country.HrefValue = ""
			Orders.pShip_Country.TooltipValue = ""

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

		' ----------
		'  Edit Row
		' ----------

		ElseIf Orders.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' OrderId
			Orders.OrderId.EditCustomAttributes = ""
			Orders.OrderId.EditValue = Orders.OrderId.CurrentValue
			Orders.OrderId.ViewCustomAttributes = ""

			' CustomerId
			Orders.CustomerId.EditCustomAttributes = ""
			If Orders.CustomerId.SessionValue <> "" Then
				Orders.CustomerId.CurrentValue = Orders.CustomerId.SessionValue
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

			' InvoiceId
			Orders.InvoiceId.EditCustomAttributes = ""
			Orders.InvoiceId.EditValue = ew_HtmlEncode(Orders.InvoiceId.CurrentValue)

			' Amount
			Orders.Amount.EditCustomAttributes = ""
			Orders.Amount.EditValue = ew_HtmlEncode(Orders.Amount.CurrentValue)

			' Ship_FirstName
			Orders.Ship_FirstName.EditCustomAttributes = ""
			Orders.Ship_FirstName.EditValue = ew_HtmlEncode(Orders.Ship_FirstName.CurrentValue)

			' Ship_LastName
			Orders.Ship_LastName.EditCustomAttributes = ""
			Orders.Ship_LastName.EditValue = ew_HtmlEncode(Orders.Ship_LastName.CurrentValue)

			' Ship_Address
			Orders.Ship_Address.EditCustomAttributes = ""
			Orders.Ship_Address.EditValue = ew_HtmlEncode(Orders.Ship_Address.CurrentValue)

			' Ship_Address2
			Orders.Ship_Address2.EditCustomAttributes = ""
			Orders.Ship_Address2.EditValue = ew_HtmlEncode(Orders.Ship_Address2.CurrentValue)

			' Ship_City
			Orders.Ship_City.EditCustomAttributes = ""
			Orders.Ship_City.EditValue = ew_HtmlEncode(Orders.Ship_City.CurrentValue)

			' Ship_Province
			Orders.Ship_Province.EditCustomAttributes = ""
			Orders.Ship_Province.EditValue = ew_HtmlEncode(Orders.Ship_Province.CurrentValue)

			' Ship_Postal
			Orders.Ship_Postal.EditCustomAttributes = ""
			Orders.Ship_Postal.EditValue = ew_HtmlEncode(Orders.Ship_Postal.CurrentValue)

			' Ship_Country
			Orders.Ship_Country.EditCustomAttributes = ""
			Orders.Ship_Country.EditValue = ew_HtmlEncode(Orders.Ship_Country.CurrentValue)

			' Ship_Phone
			Orders.Ship_Phone.EditCustomAttributes = ""
			Orders.Ship_Phone.EditValue = ew_HtmlEncode(Orders.Ship_Phone.CurrentValue)

			' Ship_Email
			Orders.Ship_Email.EditCustomAttributes = ""
			Orders.Ship_Email.EditValue = ew_HtmlEncode(Orders.Ship_Email.CurrentValue)

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
			Orders.Ordered_Date.EditValue = Orders.Ordered_Date.CurrentValue

			' payment_date
			Orders.payment_date.EditCustomAttributes = ""
			Orders.payment_date.EditValue = Orders.payment_date.CurrentValue

			' pfirst_name
			Orders.pfirst_name.EditCustomAttributes = ""
			Orders.pfirst_name.EditValue = ew_HtmlEncode(Orders.pfirst_name.CurrentValue)

			' plast_name
			Orders.plast_name.EditCustomAttributes = ""
			Orders.plast_name.EditValue = ew_HtmlEncode(Orders.plast_name.CurrentValue)

			' payer_email
			Orders.payer_email.EditCustomAttributes = ""
			Orders.payer_email.EditValue = ew_HtmlEncode(Orders.payer_email.CurrentValue)

			' txn_id
			Orders.txn_id.EditCustomAttributes = ""
			Orders.txn_id.EditValue = ew_HtmlEncode(Orders.txn_id.CurrentValue)

			' payment_gross
			Orders.payment_gross.EditCustomAttributes = ""
			Orders.payment_gross.EditValue = ew_HtmlEncode(Orders.payment_gross.CurrentValue)

			' payment_fee
			Orders.payment_fee.EditCustomAttributes = ""
			Orders.payment_fee.EditValue = ew_HtmlEncode(Orders.payment_fee.CurrentValue)

			' payment_type
			Orders.payment_type.EditCustomAttributes = ""
			Orders.payment_type.EditValue = ew_HtmlEncode(Orders.payment_type.CurrentValue)

			' txn_type
			Orders.txn_type.EditCustomAttributes = ""
			Orders.txn_type.EditValue = ew_HtmlEncode(Orders.txn_type.CurrentValue)

			' receiver_email
			Orders.receiver_email.EditCustomAttributes = ""
			Orders.receiver_email.EditValue = ew_HtmlEncode(Orders.receiver_email.CurrentValue)

			' pShip_Name
			Orders.pShip_Name.EditCustomAttributes = ""
			Orders.pShip_Name.EditValue = ew_HtmlEncode(Orders.pShip_Name.CurrentValue)

			' pShip_Address
			Orders.pShip_Address.EditCustomAttributes = ""
			Orders.pShip_Address.EditValue = ew_HtmlEncode(Orders.pShip_Address.CurrentValue)

			' pShip_City
			Orders.pShip_City.EditCustomAttributes = ""
			Orders.pShip_City.EditValue = ew_HtmlEncode(Orders.pShip_City.CurrentValue)

			' pShip_Province
			Orders.pShip_Province.EditCustomAttributes = ""
			Orders.pShip_Province.EditValue = ew_HtmlEncode(Orders.pShip_Province.CurrentValue)

			' pShip_Postal
			Orders.pShip_Postal.EditCustomAttributes = ""
			Orders.pShip_Postal.EditValue = ew_HtmlEncode(Orders.pShip_Postal.CurrentValue)

			' pShip_Country
			Orders.pShip_Country.EditCustomAttributes = ""
			Orders.pShip_Country.EditValue = ew_HtmlEncode(Orders.pShip_Country.CurrentValue)

			' Tax
			Orders.Tax.EditCustomAttributes = ""
			Orders.Tax.EditValue = ew_HtmlEncode(Orders.Tax.CurrentValue)

			' Shipping
			Orders.Shipping.EditCustomAttributes = ""
			Orders.Shipping.EditValue = ew_HtmlEncode(Orders.Shipping.CurrentValue)

			' EmailSent
			Orders.EmailSent.EditCustomAttributes = ""
			Redim arwrk(1, 0)
			arwrk(0, 0) = "confirm"
			arwrk(1, 0) = ew_IIf(Orders.EmailSent.FldTagCaption(1) <> "", Orders.EmailSent.FldTagCaption(1), "Confirm")
			arwrk = ew_AddItemToArray(arwrk, 0, Array("", Language.Phrase("PleaseSelect")))
			Orders.EmailSent.EditValue = arwrk

			' EmailDate
			Orders.EmailDate.EditCustomAttributes = ""
			Orders.EmailDate.EditValue = Orders.EmailDate.CurrentValue

			' PromoCodeUsed
			Orders.PromoCodeUsed.EditCustomAttributes = ""
			Orders.PromoCodeUsed.EditValue = ew_HtmlEncode(Orders.PromoCodeUsed.CurrentValue)

			' Edit refer script
			' OrderId

			Orders.OrderId.HrefValue = ""

			' CustomerId
			If Not ew_Empty(Orders.CustomerId.CurrentValue) Then
				Orders.CustomerId.HrefValue = "Customersedit.asp?CustomerID=" & Orders.CustomerId.CurrentValue
				Orders.CustomerId.LinkAttrs.AddAttribute "target", "", True ' Add target
				If Orders.Export <> "" Then Orders.CustomerId.HrefValue = ew_ConvertFullUrl(Orders.CustomerId.HrefValue)
			Else
				Orders.CustomerId.HrefValue = ""
			End If

			' InvoiceId
			Orders.InvoiceId.HrefValue = ""

			' Amount
			Orders.Amount.HrefValue = ""

			' Ship_FirstName
			Orders.Ship_FirstName.HrefValue = ""

			' Ship_LastName
			Orders.Ship_LastName.HrefValue = ""

			' Ship_Address
			Orders.Ship_Address.HrefValue = ""

			' Ship_Address2
			Orders.Ship_Address2.HrefValue = ""

			' Ship_City
			Orders.Ship_City.HrefValue = ""

			' Ship_Province
			Orders.Ship_Province.HrefValue = ""

			' Ship_Postal
			Orders.Ship_Postal.HrefValue = ""

			' Ship_Country
			Orders.Ship_Country.HrefValue = ""

			' Ship_Phone
			Orders.Ship_Phone.HrefValue = ""

			' Ship_Email
			Orders.Ship_Email.HrefValue = ""

			' payment_status
			Orders.payment_status.HrefValue = ""

			' Ordered_Date
			Orders.Ordered_Date.HrefValue = ""

			' payment_date
			Orders.payment_date.HrefValue = ""

			' pfirst_name
			Orders.pfirst_name.HrefValue = ""

			' plast_name
			Orders.plast_name.HrefValue = ""

			' payer_email
			Orders.payer_email.HrefValue = ""

			' txn_id
			Orders.txn_id.HrefValue = ""

			' payment_gross
			Orders.payment_gross.HrefValue = ""

			' payment_fee
			Orders.payment_fee.HrefValue = ""

			' payment_type
			Orders.payment_type.HrefValue = ""

			' txn_type
			Orders.txn_type.HrefValue = ""

			' receiver_email
			Orders.receiver_email.HrefValue = ""

			' pShip_Name
			Orders.pShip_Name.HrefValue = ""

			' pShip_Address
			Orders.pShip_Address.HrefValue = ""

			' pShip_City
			Orders.pShip_City.HrefValue = ""

			' pShip_Province
			Orders.pShip_Province.HrefValue = ""

			' pShip_Postal
			Orders.pShip_Postal.HrefValue = ""

			' pShip_Country
			Orders.pShip_Country.HrefValue = ""

			' Tax
			Orders.Tax.HrefValue = ""

			' Shipping
			Orders.Shipping.HrefValue = ""

			' EmailSent
			Orders.EmailSent.HrefValue = ""

			' EmailDate
			Orders.EmailDate.HrefValue = ""

			' PromoCodeUsed
			Orders.PromoCodeUsed.HrefValue = ""
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
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(Orders.InvoiceId.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.InvoiceId.FldErrMsg)
		End If
		If Not ew_CheckNumber(Orders.Amount.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.Amount.FldErrMsg)
		End If
		If Not ew_CheckNumber(Orders.payment_gross.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.payment_gross.FldErrMsg)
		End If
		If Not ew_CheckNumber(Orders.payment_fee.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.payment_fee.FldErrMsg)
		End If
		If Not ew_CheckNumber(Orders.Tax.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.Tax.FldErrMsg)
		End If
		If Not ew_CheckNumber(Orders.Shipping.FormValue) Then
			Call ew_AddMessage(gsFormError, Orders.Shipping.FldErrMsg)
		End If

		' Validate detail grid
		If Orders.CurrentDetailTable = "OrderDetails" And OrderDetails.DetailEdit Then
			Dim OrderDetails_grid
			Set OrderDetails_grid = new cOrderDetails_grid ' get detail page object
			Call OrderDetails_grid.ValidateGridForm()
			Set OrderDetails_grid = Nothing
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = Orders.KeyFilter
		Orders.CurrentFilter  = sFilter
		sSql = Orders.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Begin transaction
			If Orders.CurrentDetailTable <> "" Then Conn.BeginTrans

			' Field CustomerId
			Call Orders.CustomerId.SetDbValue(Rs, Orders.CustomerId.CurrentValue, Null, Orders.CustomerId.ReadOnly)

			' Field InvoiceId
			Call Orders.InvoiceId.SetDbValue(Rs, Orders.InvoiceId.CurrentValue, Null, Orders.InvoiceId.ReadOnly)

			' Field Amount
			Call Orders.Amount.SetDbValue(Rs, Orders.Amount.CurrentValue, Null, Orders.Amount.ReadOnly)

			' Field Ship_FirstName
			Call Orders.Ship_FirstName.SetDbValue(Rs, Orders.Ship_FirstName.CurrentValue, Null, Orders.Ship_FirstName.ReadOnly)

			' Field Ship_LastName
			Call Orders.Ship_LastName.SetDbValue(Rs, Orders.Ship_LastName.CurrentValue, Null, Orders.Ship_LastName.ReadOnly)

			' Field Ship_Address
			Call Orders.Ship_Address.SetDbValue(Rs, Orders.Ship_Address.CurrentValue, Null, Orders.Ship_Address.ReadOnly)

			' Field Ship_Address2
			Call Orders.Ship_Address2.SetDbValue(Rs, Orders.Ship_Address2.CurrentValue, Null, Orders.Ship_Address2.ReadOnly)

			' Field Ship_City
			Call Orders.Ship_City.SetDbValue(Rs, Orders.Ship_City.CurrentValue, Null, Orders.Ship_City.ReadOnly)

			' Field Ship_Province
			Call Orders.Ship_Province.SetDbValue(Rs, Orders.Ship_Province.CurrentValue, Null, Orders.Ship_Province.ReadOnly)

			' Field Ship_Postal
			Call Orders.Ship_Postal.SetDbValue(Rs, Orders.Ship_Postal.CurrentValue, Null, Orders.Ship_Postal.ReadOnly)

			' Field Ship_Country
			Call Orders.Ship_Country.SetDbValue(Rs, Orders.Ship_Country.CurrentValue, Null, Orders.Ship_Country.ReadOnly)

			' Field Ship_Phone
			Call Orders.Ship_Phone.SetDbValue(Rs, Orders.Ship_Phone.CurrentValue, Null, Orders.Ship_Phone.ReadOnly)

			' Field Ship_Email
			Call Orders.Ship_Email.SetDbValue(Rs, Orders.Ship_Email.CurrentValue, Null, Orders.Ship_Email.ReadOnly)

			' Field payment_status
			Call Orders.payment_status.SetDbValue(Rs, Orders.payment_status.CurrentValue, Null, Orders.payment_status.ReadOnly)

			' Field Ordered_Date
			Call Orders.Ordered_Date.SetDbValue(Rs, Orders.Ordered_Date.CurrentValue, Null, Orders.Ordered_Date.ReadOnly)

			' Field payment_date
			Call Orders.payment_date.SetDbValue(Rs, Orders.payment_date.CurrentValue, Null, Orders.payment_date.ReadOnly)

			' Field pfirst_name
			Call Orders.pfirst_name.SetDbValue(Rs, Orders.pfirst_name.CurrentValue, Null, Orders.pfirst_name.ReadOnly)

			' Field plast_name
			Call Orders.plast_name.SetDbValue(Rs, Orders.plast_name.CurrentValue, Null, Orders.plast_name.ReadOnly)

			' Field payer_email
			Call Orders.payer_email.SetDbValue(Rs, Orders.payer_email.CurrentValue, Null, Orders.payer_email.ReadOnly)

			' Field txn_id
			Call Orders.txn_id.SetDbValue(Rs, Orders.txn_id.CurrentValue, Null, Orders.txn_id.ReadOnly)

			' Field payment_gross
			Call Orders.payment_gross.SetDbValue(Rs, Orders.payment_gross.CurrentValue, Null, Orders.payment_gross.ReadOnly)

			' Field payment_fee
			Call Orders.payment_fee.SetDbValue(Rs, Orders.payment_fee.CurrentValue, Null, Orders.payment_fee.ReadOnly)

			' Field payment_type
			Call Orders.payment_type.SetDbValue(Rs, Orders.payment_type.CurrentValue, Null, Orders.payment_type.ReadOnly)

			' Field txn_type
			Call Orders.txn_type.SetDbValue(Rs, Orders.txn_type.CurrentValue, Null, Orders.txn_type.ReadOnly)

			' Field receiver_email
			Call Orders.receiver_email.SetDbValue(Rs, Orders.receiver_email.CurrentValue, Null, Orders.receiver_email.ReadOnly)

			' Field pShip_Name
			Call Orders.pShip_Name.SetDbValue(Rs, Orders.pShip_Name.CurrentValue, Null, Orders.pShip_Name.ReadOnly)

			' Field pShip_Address
			Call Orders.pShip_Address.SetDbValue(Rs, Orders.pShip_Address.CurrentValue, Null, Orders.pShip_Address.ReadOnly)

			' Field pShip_City
			Call Orders.pShip_City.SetDbValue(Rs, Orders.pShip_City.CurrentValue, Null, Orders.pShip_City.ReadOnly)

			' Field pShip_Province
			Call Orders.pShip_Province.SetDbValue(Rs, Orders.pShip_Province.CurrentValue, Null, Orders.pShip_Province.ReadOnly)

			' Field pShip_Postal
			Call Orders.pShip_Postal.SetDbValue(Rs, Orders.pShip_Postal.CurrentValue, Null, Orders.pShip_Postal.ReadOnly)

			' Field pShip_Country
			Call Orders.pShip_Country.SetDbValue(Rs, Orders.pShip_Country.CurrentValue, Null, Orders.pShip_Country.ReadOnly)

			' Field Tax
			Call Orders.Tax.SetDbValue(Rs, Orders.Tax.CurrentValue, Null, Orders.Tax.ReadOnly)

			' Field Shipping
			Call Orders.Shipping.SetDbValue(Rs, Orders.Shipping.CurrentValue, Null, Orders.Shipping.ReadOnly)

			' Field EmailSent
			Call Orders.EmailSent.SetDbValue(Rs, Orders.EmailSent.CurrentValue, Null, Orders.EmailSent.ReadOnly)

			' Field EmailDate
			Call Orders.EmailDate.SetDbValue(Rs, Orders.EmailDate.CurrentValue, Null, Orders.EmailDate.ReadOnly)

			' Field PromoCodeUsed
			Call Orders.PromoCodeUsed.SetDbValue(Rs, Orders.PromoCodeUsed.CurrentValue, Null, Orders.PromoCodeUsed.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = Orders.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				Rs.Update
				If Err.Number <> 0 Then
					FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If

				' Update detail records
				If EditRow Then
					If Orders.CurrentDetailTable = "OrderDetails" And OrderDetails.DetailEdit Then
						Dim OrderDetails_grid
						Set OrderDetails_grid = New cOrderDetails_grid ' get detail page object
						EditRow = OrderDetails_grid.GridUpdate
						Set OrderDetails_grid = Nothing
					End If
				End If

				' Commit/Rollback transaction
				If Orders.CurrentDetailTable <> "" Then
					If EditRow Then
						Conn.CommitTrans ' Commit transaction
					Else
						Conn.RollbackTrans ' Rollback transaction
					End If
				End If
			Else
				Rs.CancelUpdate
				If Orders.CancelMessage <> "" Then
					FailureMessage = Orders.CancelMessage
					Orders.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call Orders.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
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

	' Set up detail parms based on QueryString
	Sub SetUpDetailParms()
		Dim sDetailTblVar, bValidDetail
		bValidDetail = False

		' Get the keys for master table
		If Request.QueryString(EW_TABLE_SHOW_DETAIL).Count > 0 Then
			sDetailTblVar = Request.QueryString(EW_TABLE_SHOW_DETAIL)
			Orders.CurrentDetailTable = sDetailTblVar
		Else
			sDetailTblVar = Orders.CurrentDetailTable
		End If
		If sDetailTblVar <> "" Then
			If sDetailTblVar = "OrderDetails" Then
				If IsEmpty(OrderDetails) Then
					Set OrderDetails = New cOrderDetails
				End If
				If OrderDetails.DetailEdit Then
					OrderDetails.CurrentMode = "edit"
					OrderDetails.CurrentAction = "gridedit"

					' Save current master table to detail table
					OrderDetails.CurrentMasterTable = Orders.TableVar
					OrderDetails.StartRecordNumber = 1
					OrderDetails.OrderId.FldIsDetailKey = True
					OrderDetails.OrderId.CurrentValue = Orders.OrderId.CurrentValue
					OrderDetails.OrderId.SessionValue = OrderDetails.OrderId.CurrentValue
				End If
			End If
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
End Class
%>