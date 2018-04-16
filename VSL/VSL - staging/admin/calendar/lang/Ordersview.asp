<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Ordersinfo.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Orders_view
Set Orders_view = New cOrders_view
Set Page = Orders_view

' Page init processing
Call Orders_view.Page_Init()

' Page main processing
Call Orders_view.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Orders.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Orders_view = new ew_Page("Orders_view");
// page properties
Orders_view.PageID = "view"; // page ID
Orders_view.FormID = "fOrdersview"; // form ID
var EW_PAGE_ID = Orders_view.PageID; // for backward compatibility
// extend page with Form_CustomValidate function
Orders_view.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Orders_view.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Orders_view.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Orders_view.ValidateRequired = false; // no JavaScript validation
<% End If %>
// multi page properties
Orders_view.MultiPage = new ew_MultiPage();
Orders_view.MultiPage.AddElement("x_OrderId", 1);
Orders_view.MultiPage.AddElement("x_CustomerId", 1);
Orders_view.MultiPage.AddElement("x_InvoiceId", 1);
Orders_view.MultiPage.AddElement("x_Amount", 1);
Orders_view.MultiPage.AddElement("x_Ship_FirstName", 1);
Orders_view.MultiPage.AddElement("x_Ship_LastName", 1);
Orders_view.MultiPage.AddElement("x_Ship_Address", 1);
Orders_view.MultiPage.AddElement("x_Ship_Address2", 1);
Orders_view.MultiPage.AddElement("x_Ship_City", 1);
Orders_view.MultiPage.AddElement("x_Ship_Province", 1);
Orders_view.MultiPage.AddElement("x_Ship_Postal", 1);
Orders_view.MultiPage.AddElement("x_Ship_Country", 1);
Orders_view.MultiPage.AddElement("x_Ship_Phone", 1);
Orders_view.MultiPage.AddElement("x_Ship_Email", 1);
Orders_view.MultiPage.AddElement("x_payment_status", 1);
Orders_view.MultiPage.AddElement("x_Ordered_Date", 1);
Orders_view.MultiPage.AddElement("x_payment_date", 1);
Orders_view.MultiPage.AddElement("x_pfirst_name", 2);
Orders_view.MultiPage.AddElement("x_plast_name", 2);
Orders_view.MultiPage.AddElement("x_payer_email", 2);
Orders_view.MultiPage.AddElement("x_txn_id", 1);
Orders_view.MultiPage.AddElement("x_payment_gross", 2);
Orders_view.MultiPage.AddElement("x_payment_fee", 2);
Orders_view.MultiPage.AddElement("x_payment_type", 2);
Orders_view.MultiPage.AddElement("x_txn_type", 2);
Orders_view.MultiPage.AddElement("x_receiver_email", 2);
Orders_view.MultiPage.AddElement("x_pShip_Name", 2);
Orders_view.MultiPage.AddElement("x_pShip_Address", 2);
Orders_view.MultiPage.AddElement("x_pShip_City", 2);
Orders_view.MultiPage.AddElement("x_pShip_Province", 2);
Orders_view.MultiPage.AddElement("x_pShip_Postal", 2);
Orders_view.MultiPage.AddElement("x_pShip_Country", 2);
Orders_view.MultiPage.AddElement("x_Tax", 1);
Orders_view.MultiPage.AddElement("x_Shipping", 1);
Orders_view.MultiPage.AddElement("x_EmailSent", 1);
Orders_view.MultiPage.AddElement("x_EmailDate", 1);
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
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% Orders_view.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("View") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Orders.TableCaption %>
&nbsp;&nbsp;<% Orders_view.ExportOptions.Render "body", "" %>
</p>
<% If Orders.Export = "" Then %>
<p class="aspmaker">
<a href="<%= Orders_view.ListUrl %>"><%= Language.Phrase("BackToList") %></a>&nbsp;
<% If Security.IsLoggedIn() Then %>
<a href="<%= Orders_view.AddUrl %>"><%= Language.Phrase("ViewPageAddLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Orders_view.EditUrl %>"><%= Language.Phrase("ViewPageEditLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<a href="<%= Orders_view.DeleteUrl %>"><%= Language.Phrase("ViewPageDeleteLink") %></a>&nbsp;
<% End If %>
<% If Security.IsLoggedIn() Then %>
<%
sSqlWrk = "[OrderId]=" & ew_AdjustSql(Orders.OrderId.CurrentValue) & ""
sSqlWrk = ew_Encode(TEAencrypt(sSqlWrk, EW_RANDOM_KEY))
sSqlWrk = Replace(sSqlWrk, "'", "\'")
%>
<a href="OrderDetailslist.asp?<%= EW_TABLE_SHOW_MASTER %>=Orders&OrderId=<%= Server.URLEncode(Orders.OrderId.CurrentValue&"") %>" name="ew_Orders_OrderDetails_DetailLink" id="ew_Orders_OrderDetails_DetailLink" onmouseover="ew_AjaxShowDetails(this, 'OrderDetailspreview.asp?f=<%= sSqlWrk %>')" onmouseout="ew_AjaxHideDetails(this);"><%= Language.TablePhrase("OrderDetails", "TblCaption") %>
</a>
&nbsp;
<% End If %>
<% End If %>
</p>
<% Orders_view.ShowMessage %>
<p>
<% If Orders.Export = "" Then %>
<table cellspacing="0" cellpadding="0"><tr><td>
<div id="Orders_view" class="yui-navset">
	<ul class="yui-nav">
		<li class="selected"><a href="#tab_Orders_1"><em><%= Orders.PageCaption(1) %></em></a></li>
		<li><a href="#tab_Orders_2"><em><%= Orders.PageCaption(2) %></em></a></li>
	</ul>            
	<div class="yui-content">
<% End If %>
		<div id="tab_Orders_1">
<table cellspacing="0" class="ewGrid" style="width: 100%"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Orders.OrderId.Visible Then ' OrderId %>
	<tr id="r_OrderId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.OrderId.FldCaption %></td>
		<td<%= Orders.OrderId.CellAttributes %>>
<div<%= Orders.OrderId.ViewAttributes %>><%= Orders.OrderId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.CustomerId.Visible Then ' CustomerId %>
	<tr id="r_CustomerId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.CustomerId.FldCaption %></td>
		<td<%= Orders.CustomerId.CellAttributes %>>
<div<%= Orders.CustomerId.ViewAttributes %>><%= Orders.CustomerId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.InvoiceId.Visible Then ' InvoiceId %>
	<tr id="r_InvoiceId"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.InvoiceId.FldCaption %></td>
		<td<%= Orders.InvoiceId.CellAttributes %>>
<div<%= Orders.InvoiceId.ViewAttributes %>><%= Orders.InvoiceId.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Amount.Visible Then ' Amount %>
	<tr id="r_Amount"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Amount.FldCaption %></td>
		<td<%= Orders.Amount.CellAttributes %>>
<div<%= Orders.Amount.ViewAttributes %>><%= Orders.Amount.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_FirstName.Visible Then ' Ship_FirstName %>
	<tr id="r_Ship_FirstName"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_FirstName.FldCaption %></td>
		<td<%= Orders.Ship_FirstName.CellAttributes %>>
<div<%= Orders.Ship_FirstName.ViewAttributes %>><%= Orders.Ship_FirstName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_LastName.Visible Then ' Ship_LastName %>
	<tr id="r_Ship_LastName"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_LastName.FldCaption %></td>
		<td<%= Orders.Ship_LastName.CellAttributes %>>
<div<%= Orders.Ship_LastName.ViewAttributes %>><%= Orders.Ship_LastName.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Address.Visible Then ' Ship_Address %>
	<tr id="r_Ship_Address"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Address.FldCaption %></td>
		<td<%= Orders.Ship_Address.CellAttributes %>>
<div<%= Orders.Ship_Address.ViewAttributes %>><%= Orders.Ship_Address.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Address2.Visible Then ' Ship_Address2 %>
	<tr id="r_Ship_Address2"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Address2.FldCaption %></td>
		<td<%= Orders.Ship_Address2.CellAttributes %>>
<div<%= Orders.Ship_Address2.ViewAttributes %>><%= Orders.Ship_Address2.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_City.Visible Then ' Ship_City %>
	<tr id="r_Ship_City"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_City.FldCaption %></td>
		<td<%= Orders.Ship_City.CellAttributes %>>
<div<%= Orders.Ship_City.ViewAttributes %>><%= Orders.Ship_City.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Province.Visible Then ' Ship_Province %>
	<tr id="r_Ship_Province"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Province.FldCaption %></td>
		<td<%= Orders.Ship_Province.CellAttributes %>>
<div<%= Orders.Ship_Province.ViewAttributes %>><%= Orders.Ship_Province.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Postal.Visible Then ' Ship_Postal %>
	<tr id="r_Ship_Postal"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Postal.FldCaption %></td>
		<td<%= Orders.Ship_Postal.CellAttributes %>>
<div<%= Orders.Ship_Postal.ViewAttributes %>><%= Orders.Ship_Postal.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Country.Visible Then ' Ship_Country %>
	<tr id="r_Ship_Country"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Country.FldCaption %></td>
		<td<%= Orders.Ship_Country.CellAttributes %>>
<div<%= Orders.Ship_Country.ViewAttributes %>><%= Orders.Ship_Country.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Phone.Visible Then ' Ship_Phone %>
	<tr id="r_Ship_Phone"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Phone.FldCaption %></td>
		<td<%= Orders.Ship_Phone.CellAttributes %>>
<div<%= Orders.Ship_Phone.ViewAttributes %>><%= Orders.Ship_Phone.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ship_Email.Visible Then ' Ship_Email %>
	<tr id="r_Ship_Email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ship_Email.FldCaption %></td>
		<td<%= Orders.Ship_Email.CellAttributes %>>
<div<%= Orders.Ship_Email.ViewAttributes %>><%= Orders.Ship_Email.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payment_status.Visible Then ' payment_status %>
	<tr id="r_payment_status"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_status.FldCaption %></td>
		<td<%= Orders.payment_status.CellAttributes %>>
<div<%= Orders.payment_status.ViewAttributes %>><%= Orders.payment_status.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Ordered_Date.Visible Then ' Ordered_Date %>
	<tr id="r_Ordered_Date"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Ordered_Date.FldCaption %></td>
		<td<%= Orders.Ordered_Date.CellAttributes %>>
<div<%= Orders.Ordered_Date.ViewAttributes %>><%= Orders.Ordered_Date.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payment_date.Visible Then ' payment_date %>
	<tr id="r_payment_date"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_date.FldCaption %></td>
		<td<%= Orders.payment_date.CellAttributes %>>
<div<%= Orders.payment_date.ViewAttributes %>><%= Orders.payment_date.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.txn_id.Visible Then ' txn_id %>
	<tr id="r_txn_id"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.txn_id.FldCaption %></td>
		<td<%= Orders.txn_id.CellAttributes %>>
<div<%= Orders.txn_id.ViewAttributes %>><%= Orders.txn_id.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Tax.Visible Then ' Tax %>
	<tr id="r_Tax"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Tax.FldCaption %></td>
		<td<%= Orders.Tax.CellAttributes %>>
<div<%= Orders.Tax.ViewAttributes %>><%= Orders.Tax.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.Shipping.Visible Then ' Shipping %>
	<tr id="r_Shipping"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.Shipping.FldCaption %></td>
		<td<%= Orders.Shipping.CellAttributes %>>
<div<%= Orders.Shipping.ViewAttributes %>><%= Orders.Shipping.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.EmailSent.Visible Then ' EmailSent %>
	<tr id="r_EmailSent"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.EmailSent.FldCaption %></td>
		<td<%= Orders.EmailSent.CellAttributes %>>
<div<%= Orders.EmailSent.ViewAttributes %>><%= Orders.EmailSent.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.EmailDate.Visible Then ' EmailDate %>
	<tr id="r_EmailDate"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.EmailDate.FldCaption %></td>
		<td<%= Orders.EmailDate.CellAttributes %>>
<div<%= Orders.EmailDate.ViewAttributes %>><%= Orders.EmailDate.ViewValue %></div>
</td>
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
		<td<%= Orders.pfirst_name.CellAttributes %>>
<div<%= Orders.pfirst_name.ViewAttributes %>><%= Orders.pfirst_name.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.plast_name.Visible Then ' plast_name %>
	<tr id="r_plast_name"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.plast_name.FldCaption %></td>
		<td<%= Orders.plast_name.CellAttributes %>>
<div<%= Orders.plast_name.ViewAttributes %>><%= Orders.plast_name.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payer_email.Visible Then ' payer_email %>
	<tr id="r_payer_email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payer_email.FldCaption %></td>
		<td<%= Orders.payer_email.CellAttributes %>>
<div<%= Orders.payer_email.ViewAttributes %>><%= Orders.payer_email.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payment_gross.Visible Then ' payment_gross %>
	<tr id="r_payment_gross"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_gross.FldCaption %></td>
		<td<%= Orders.payment_gross.CellAttributes %>>
<div<%= Orders.payment_gross.ViewAttributes %>><%= Orders.payment_gross.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payment_fee.Visible Then ' payment_fee %>
	<tr id="r_payment_fee"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_fee.FldCaption %></td>
		<td<%= Orders.payment_fee.CellAttributes %>>
<div<%= Orders.payment_fee.ViewAttributes %>><%= Orders.payment_fee.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.payment_type.Visible Then ' payment_type %>
	<tr id="r_payment_type"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.payment_type.FldCaption %></td>
		<td<%= Orders.payment_type.CellAttributes %>>
<div<%= Orders.payment_type.ViewAttributes %>><%= Orders.payment_type.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.txn_type.Visible Then ' txn_type %>
	<tr id="r_txn_type"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.txn_type.FldCaption %></td>
		<td<%= Orders.txn_type.CellAttributes %>>
<div<%= Orders.txn_type.ViewAttributes %>><%= Orders.txn_type.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.receiver_email.Visible Then ' receiver_email %>
	<tr id="r_receiver_email"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.receiver_email.FldCaption %></td>
		<td<%= Orders.receiver_email.CellAttributes %>>
<div<%= Orders.receiver_email.ViewAttributes %>><%= Orders.receiver_email.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_Name.Visible Then ' pShip_Name %>
	<tr id="r_pShip_Name"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Name.FldCaption %></td>
		<td<%= Orders.pShip_Name.CellAttributes %>>
<div<%= Orders.pShip_Name.ViewAttributes %>><%= Orders.pShip_Name.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_Address.Visible Then ' pShip_Address %>
	<tr id="r_pShip_Address"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Address.FldCaption %></td>
		<td<%= Orders.pShip_Address.CellAttributes %>>
<div<%= Orders.pShip_Address.ViewAttributes %>><%= Orders.pShip_Address.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_City.Visible Then ' pShip_City %>
	<tr id="r_pShip_City"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_City.FldCaption %></td>
		<td<%= Orders.pShip_City.CellAttributes %>>
<div<%= Orders.pShip_City.ViewAttributes %>><%= Orders.pShip_City.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_Province.Visible Then ' pShip_Province %>
	<tr id="r_pShip_Province"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Province.FldCaption %></td>
		<td<%= Orders.pShip_Province.CellAttributes %>>
<div<%= Orders.pShip_Province.ViewAttributes %>><%= Orders.pShip_Province.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_Postal.Visible Then ' pShip_Postal %>
	<tr id="r_pShip_Postal"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Postal.FldCaption %></td>
		<td<%= Orders.pShip_Postal.CellAttributes %>>
<div<%= Orders.pShip_Postal.ViewAttributes %>><%= Orders.pShip_Postal.ViewValue %></div>
</td>
	</tr>
<% End If %>
<% If Orders.pShip_Country.Visible Then ' pShip_Country %>
	<tr id="r_pShip_Country"<%= Orders.RowAttributes %>>
		<td class="ewTableHeader"><%= Orders.pShip_Country.FldCaption %></td>
		<td<%= Orders.pShip_Country.CellAttributes %>>
<div<%= Orders.pShip_Country.ViewAttributes %>><%= Orders.pShip_Country.ViewValue %></div>
</td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
		</div>
<% If Orders.Export = "" Then %>
	</div>
</div>
</td></tr></table>
<% End If %>
<% If Orders.Export = "" Then %>
<script type="text/javascript">
ew_TabView(Orders_view);
//-->
</script>
<% End If %>
<p>
<%
Orders_view.ShowPageFooter()
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
Set Orders_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cOrders_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Orders"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Orders_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("OrderId").Count > 0 Then
			ew_AddKey RecKey, "OrderId", Request.QueryString("OrderId")
			KeyUrl = KeyUrl & "&OrderId=" & Server.URLEncode(Request.QueryString("OrderId"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Orders"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

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

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False
		If IsPageRequest Then ' Validate request
			If Request.QueryString("OrderId").Count > 0 Then
				Orders.OrderId.QueryStringValue = Request.QueryString("OrderId")
			Else
				sReturnUrl = "Orderslist.asp" ' Return to list
			End If

			' Get action
			Orders.CurrentAction = "I" ' Display form
			Select Case Orders.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "Orderslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "Orderslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		Orders.RowType = EW_ROWTYPE_VIEW
		Call Orders.ResetAttrs()
		Call RenderRow()
	End Sub
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
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = Orders.AddUrl
		EditUrl = Orders.EditUrl("")
		CopyUrl = Orders.CopyUrl("")
		DeleteUrl = Orders.DeleteUrl
		ListUrl = Orders.ListUrl

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
			Orders.EmailSent.ViewValue = Orders.EmailSent.CurrentValue
			Orders.EmailSent.ViewCustomAttributes = ""

			' EmailDate
			Orders.EmailDate.ViewValue = Orders.EmailDate.CurrentValue
			Orders.EmailDate.ViewCustomAttributes = ""

			' View refer script
			' OrderId

			Orders.OrderId.LinkCustomAttributes = ""
			Orders.OrderId.HrefValue = ""
			Orders.OrderId.TooltipValue = ""

			' CustomerId
			Orders.CustomerId.LinkCustomAttributes = ""
			Orders.CustomerId.HrefValue = ""
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
		End If

		' Call Row Rendered event
		If Orders.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Orders.Row_Rendered()
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
End Class
%>
