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
Dim Products_add
Set Products_add = New cProducts_add
Set Page = Products_add

' Page init processing
Call Products_add.Page_Init()

' Page main processing
Call Products_add.Page_Main()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
// Create page object
var Products_add = new ew_Page("Products_add");
// page properties
Products_add.PageID = "add"; // page ID
Products_add.FormID = "fProductsadd"; // form ID
var EW_PAGE_ID = Products_add.PageID; // for backward compatibility
// extend page with ValidateForm function
Products_add.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = 1;
	for (i=0; i<rowcnt; i++) {
		infix = "";
		elm = fobj.elements["x" + infix + "_Image"];
		if (elm && !ew_CheckFileType(elm.value))
			return ew_OnError(this, elm, ewLanguage.Phrase("WrongFileType"));
		elm = fobj.elements["x" + infix + "_Image_Thumb"];
		if (elm && !ew_CheckFileType(elm.value))
			return ew_OnError(this, elm, ewLanguage.Phrase("WrongFileType"));
		elm = fobj.elements["x" + infix + "_fImage"];
		if (elm && !ew_CheckFileType(elm.value))
			return ew_OnError(this, elm, ewLanguage.Phrase("WrongFileType"));
		elm = fobj.elements["x" + infix + "_fImage_Thumb"];
		if (elm && !ew_CheckFileType(elm.value))
			return ew_OnError(this, elm, ewLanguage.Phrase("WrongFileType"));
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
Products_add.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Products_add.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Products_add.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Products_add.ValidateRequired = false; // no JavaScript validation
<% End If %>
// multi page properties
Products_add.MultiPage = new ew_MultiPage();
Products_add.MultiPage.AddElement("x_Description", 1);
Products_add.MultiPage.AddElement("x_Price", 1);
Products_add.MultiPage.AddElement("x_Active", 1);
Products_add.MultiPage.AddElement("x_Image", 1);
Products_add.MultiPage.AddElement("x_Sizes", 1);
Products_add.MultiPage.AddElement("x_Image_Thumb", 1);
Products_add.MultiPage.AddElement("x_ProductName", 1);
Products_add.MultiPage.AddElement("x_ItemNo", 1);
Products_add.MultiPage.AddElement("x_UPC", 1);
Products_add.MultiPage.AddElement("x_fDescription", 2);
Products_add.MultiPage.AddElement("x_fImage", 2);
Products_add.MultiPage.AddElement("x_fSizes", 2);
Products_add.MultiPage.AddElement("x_fImage_Thumb", 2);
Products_add.MultiPage.AddElement("x_fProductName", 2);
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
<% Products_add.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("Add") %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Products.TableCaption %></p>
<p class="aspmaker"><a href="<%= Products.ReturnUrl %>"><%= Language.Phrase("GoBack") %></a></p>
<% Products_add.ShowMessage %>
<form name="fProductsadd" id="fProductsadd" action="<%= ew_CurrentPage %>" method="post" enctype="multipart/form-data" onsubmit="return Products_add.ValidateForm(this);">
<p>
<input type="hidden" name="t" id="t" value="Products">
<input type="hidden" name="a_add" id="a_add" value="A">
<table cellspacing="0" cellpadding="0"><tr><td>
<div id="Products_add" class="yui-navset">
	<ul class="yui-nav">
		<li class="selected"><a href="#tab_Products_1"><em><span class="aspmaker"><%= Products.PageCaption(1) %></span></em></a></li>
		<li><a href="#tab_Products_2"><em><span class="aspmaker"><%= Products.PageCaption(2) %></span></em></a></li>
	</ul>
	<div class="yui-content">
		<div id="tab_Products_1">
<table cellspacing="0" class="ewGrid" style="width: 100%"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Products.Description.Visible Then ' Description %>
	<tr id="r_Description"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Description.FldCaption %></td>
		<td<%= Products.Description.CellAttributes %>><span id="el_Description">
<input type="text" name="x_Description" id="x_Description" size="30" maxlength="100" value="<%= Products.Description.EditValue %>"<%= Products.Description.EditAttributes %>>
</span><%= Products.Description.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.Price.Visible Then ' Price %>
	<tr id="r_Price"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Price.FldCaption %></td>
		<td<%= Products.Price.CellAttributes %>><span id="el_Price">
<input type="text" name="x_Price" id="x_Price" size="30" maxlength="50" value="<%= Products.Price.EditValue %>"<%= Products.Price.EditAttributes %>>
</span><%= Products.Price.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.Active.Visible Then ' Active %>
	<tr id="r_Active"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Active.FldCaption %></td>
		<td<%= Products.Active.CellAttributes %>><span id="el_Active">
<% selwrk = ew_IIf(ew_ConvertToBool(Products.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x_Active" id="x_Active" value="1"<%= selwrk %><%= Products.Active.EditAttributes %>>
</span><%= Products.Active.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.Image.Visible Then ' Image %>
	<tr id="r_Image"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Image.FldCaption %></td>
		<td<%= Products.Image.CellAttributes %>><span id="el_Image">
<div id="old_x_Image">
<% If Products.Image.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image.UploadPath) & Products.Image.Upload.DbValue %>" border=0<%= Products.Image.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image.UploadPath) & Products.Image.Upload.DbValue %>" border=0<%= Products.Image.ViewAttributes %>>
<% End If %>
<% End If %>
</div>
<div id="new_x_Image">
<% If Not ew_Empty(Products.Image.Upload.DbValue) Then %>
<label><input type="radio" name="a_Image" id="a_Image" value="1" checked><%= Language.Phrase("Keep") %></label>&nbsp;
<label><input type="radio" name="a_Image" id="a_Image" value="2"><%= Language.Phrase("Remove") %></label>&nbsp;
<label><input type="radio" name="a_Image" id="a_Image" value="3"><%= Language.Phrase("Replace") %></label><br>
<% Products.Image.EditAttrs.AddAttribute "onchange", "if (this.form.a_Image[2]) this.form.a_Image[2].checked=true;", True %>
<% Else %>
<input type="hidden" name="a_Image" id="a_Image" value="3">
<% End If %>
<input type="file" name="x_Image" id="x_Image" size="30"<%= Products.Image.EditAttributes %>>
</div>
</span><%= Products.Image.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.Sizes.Visible Then ' Sizes %>
	<tr id="r_Sizes"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Sizes.FldCaption %></td>
		<td<%= Products.Sizes.CellAttributes %>><span id="el_Sizes">
<input type="text" name="x_Sizes" id="x_Sizes" size="30" maxlength="10" value="<%= Products.Sizes.EditValue %>"<%= Products.Sizes.EditAttributes %>>
</span><%= Products.Sizes.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.Image_Thumb.Visible Then ' Image_Thumb %>
	<tr id="r_Image_Thumb"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.Image_Thumb.FldCaption %></td>
		<td<%= Products.Image_Thumb.CellAttributes %>><span id="el_Image_Thumb">
<div id="old_x_Image_Thumb">
<% If Products.Image_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue %>" border=0<%= Products.Image_Thumb.ViewAttributes %>>
<% End If %>
<% End If %>
</div>
<div id="new_x_Image_Thumb">
<% If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then %>
<label><input type="radio" name="a_Image_Thumb" id="a_Image_Thumb" value="1" checked><%= Language.Phrase("Keep") %></label>&nbsp;
<label><input type="radio" name="a_Image_Thumb" id="a_Image_Thumb" value="2"><%= Language.Phrase("Remove") %></label>&nbsp;
<label><input type="radio" name="a_Image_Thumb" id="a_Image_Thumb" value="3"><%= Language.Phrase("Replace") %></label><br>
<% Products.Image_Thumb.EditAttrs.AddAttribute "onchange", "if (this.form.a_Image_Thumb[2]) this.form.a_Image_Thumb[2].checked=true;", True %>
<% Else %>
<input type="hidden" name="a_Image_Thumb" id="a_Image_Thumb" value="3">
<% End If %>
<input type="file" name="x_Image_Thumb" id="x_Image_Thumb" size="30"<%= Products.Image_Thumb.EditAttributes %>>
</div>
</span><%= Products.Image_Thumb.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.ProductName.Visible Then ' ProductName %>
	<tr id="r_ProductName"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.ProductName.FldCaption %></td>
		<td<%= Products.ProductName.CellAttributes %>><span id="el_ProductName">
<input type="text" name="x_ProductName" id="x_ProductName" size="30" maxlength="150" value="<%= Products.ProductName.EditValue %>"<%= Products.ProductName.EditAttributes %>>
</span><%= Products.ProductName.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.ItemNo.Visible Then ' ItemNo %>
	<tr id="r_ItemNo"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.ItemNo.FldCaption %></td>
		<td<%= Products.ItemNo.CellAttributes %>><span id="el_ItemNo">
<input type="text" name="x_ItemNo" id="x_ItemNo" size="30" maxlength="50" value="<%= Products.ItemNo.EditValue %>"<%= Products.ItemNo.EditAttributes %>>
</span><%= Products.ItemNo.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.UPC.Visible Then ' UPC %>
	<tr id="r_UPC"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.UPC.FldCaption %></td>
		<td<%= Products.UPC.CellAttributes %>><span id="el_UPC">
<input type="text" name="x_UPC" id="x_UPC" size="30" maxlength="50" value="<%= Products.UPC.EditValue %>"<%= Products.UPC.EditAttributes %>>
</span><%= Products.UPC.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</div>
</td></tr></table>
		</div>
		<div id="tab_Products_2">
<table cellspacing="0" class="ewGrid" style="width: 100%"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable">
<% If Products.fDescription.Visible Then ' fDescription %>
	<tr id="r_fDescription"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fDescription.FldCaption %></td>
		<td<%= Products.fDescription.CellAttributes %>><span id="el_fDescription">
<input type="text" name="x_fDescription" id="x_fDescription" size="30" maxlength="100" value="<%= Products.fDescription.EditValue %>"<%= Products.fDescription.EditAttributes %>>
</span><%= Products.fDescription.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.fImage.Visible Then ' fImage %>
	<tr id="r_fImage"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fImage.FldCaption %></td>
		<td<%= Products.fImage.CellAttributes %>><span id="el_fImage">
<div id="old_x_fImage">
<% If Products.fImage.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.fImage.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage.UploadPath) & Products.fImage.Upload.DbValue %>" border=0<%= Products.fImage.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.fImage.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage.UploadPath) & Products.fImage.Upload.DbValue %>" border=0<%= Products.fImage.ViewAttributes %>>
<% End If %>
<% End If %>
</div>
<div id="new_x_fImage">
<% If Not ew_Empty(Products.fImage.Upload.DbValue) Then %>
<label><input type="radio" name="a_fImage" id="a_fImage" value="1" checked><%= Language.Phrase("Keep") %></label>&nbsp;
<label><input type="radio" name="a_fImage" id="a_fImage" value="2"><%= Language.Phrase("Remove") %></label>&nbsp;
<label><input type="radio" name="a_fImage" id="a_fImage" value="3"><%= Language.Phrase("Replace") %></label><br>
<% Products.fImage.EditAttrs.AddAttribute "onchange", "if (this.form.a_fImage[2]) this.form.a_fImage[2].checked=true;", True %>
<% Else %>
<input type="hidden" name="a_fImage" id="a_fImage" value="3">
<% End If %>
<input type="file" name="x_fImage" id="x_fImage" size="30"<%= Products.fImage.EditAttributes %>>
</div>
</span><%= Products.fImage.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.fSizes.Visible Then ' fSizes %>
	<tr id="r_fSizes"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fSizes.FldCaption %></td>
		<td<%= Products.fSizes.CellAttributes %>><span id="el_fSizes">
<input type="text" name="x_fSizes" id="x_fSizes" size="30" maxlength="10" value="<%= Products.fSizes.EditValue %>"<%= Products.fSizes.EditAttributes %>>
</span><%= Products.fSizes.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.fImage_Thumb.Visible Then ' fImage_Thumb %>
	<tr id="r_fImage_Thumb"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fImage_Thumb.FldCaption %></td>
		<td<%= Products.fImage_Thumb.CellAttributes %>><span id="el_fImage_Thumb">
<div id="old_x_fImage_Thumb">
<% If Products.fImage_Thumb.LinkAttributes <> "" Then %>
<% If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage_Thumb.UploadPath) & Products.fImage_Thumb.Upload.DbValue %>" border=0<%= Products.fImage_Thumb.ViewAttributes %>>
<% End If %>
<% Else %>
<% If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then %>
<img src="<%= ew_UploadPathEx(False, Products.fImage_Thumb.UploadPath) & Products.fImage_Thumb.Upload.DbValue %>" border=0<%= Products.fImage_Thumb.ViewAttributes %>>
<% End If %>
<% End If %>
</div>
<div id="new_x_fImage_Thumb">
<% If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then %>
<label><input type="radio" name="a_fImage_Thumb" id="a_fImage_Thumb" value="1" checked><%= Language.Phrase("Keep") %></label>&nbsp;
<label><input type="radio" name="a_fImage_Thumb" id="a_fImage_Thumb" value="2"><%= Language.Phrase("Remove") %></label>&nbsp;
<label><input type="radio" name="a_fImage_Thumb" id="a_fImage_Thumb" value="3"><%= Language.Phrase("Replace") %></label><br>
<% Products.fImage_Thumb.EditAttrs.AddAttribute "onchange", "if (this.form.a_fImage_Thumb[2]) this.form.a_fImage_Thumb[2].checked=true;", True %>
<% Else %>
<input type="hidden" name="a_fImage_Thumb" id="a_fImage_Thumb" value="3">
<% End If %>
<input type="file" name="x_fImage_Thumb" id="x_fImage_Thumb" size="30"<%= Products.fImage_Thumb.EditAttributes %>>
</div>
</span><%= Products.fImage_Thumb.CustomMsg %></td>
	</tr>
<% End If %>
<% If Products.fProductName.Visible Then ' fProductName %>
	<tr id="r_fProductName"<%= Products.RowAttributes %>>
		<td class="ewTableHeader"><%= Products.fProductName.FldCaption %></td>
		<td<%= Products.fProductName.CellAttributes %>><span id="el_fProductName">
<input type="text" name="x_fProductName" id="x_fProductName" size="30" maxlength="150" value="<%= Products.fProductName.EditValue %>"<%= Products.fProductName.EditAttributes %>>
</span><%= Products.fProductName.CustomMsg %></td>
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
ew_TabView(Products_add);
//-->
</script>	
<p>
<input type="submit" name="btnAction" id="btnAction" value="<%= ew_BtnCaption(Language.Phrase("AddBtn")) %>">
</form>
<%
Products_add.ShowPageFooter()
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
Set Products_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cProducts_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Products"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Products_add"
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
		' Initialize other table object

		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Products"

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
	If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then
		Set ObjForm = New cFormObj
	Else
		Set ObjForm = ew_GetUploadObj()
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
		Set Products = Nothing
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
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			Products.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

			' Validate Form
			If Not ValidateForm() Then
				Products.CurrentAction = "I" ' Form error, reset action
				Products.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("ItemId").Count > 0 Then
				Products.ItemId.QueryStringValue = Request.QueryString("ItemId")
				Call Products.SetKey("ItemId", Products.ItemId.CurrentValue) ' Set up key
			Else
				Call Products.SetKey("ItemId", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				Products.CurrentAction = "C" ' Copy Record
			Else
				Products.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Perform action based on action code
		Select Case Products.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("Productslist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				Products.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = Products.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "Productsview.asp" Then sReturnUrl = Products.ViewUrl ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					Products.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		Products.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call Products.ResetAttrs()
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
		Products.Image.Upload.Index = ObjForm.Index
		Call Products.Image.Upload.RestoreDbFromSession()
		If confirmPage Then ' Post from confirm page
			Call Products.Image.Upload.RestoreFromSession()
		Else
			If Products.Image.Upload.UploadFile() Then

				' No action required
			Else
				Response.Write Products.Image.Upload.Message
				Page_Terminate("")
				Response.End
			End If
			Call Products.Image.Upload.SaveToSession()
			Products.Image.CurrentValue = Products.Image.Upload.FileName
		End If
		Products.Image_Thumb.Upload.Index = ObjForm.Index
		Call Products.Image_Thumb.Upload.RestoreDbFromSession()
		If confirmPage Then ' Post from confirm page
			Call Products.Image_Thumb.Upload.RestoreFromSession()
		Else
			If Products.Image_Thumb.Upload.UploadFile() Then

				' No action required
			Else
				Response.Write Products.Image_Thumb.Upload.Message
				Page_Terminate("")
				Response.End
			End If
			Call Products.Image_Thumb.Upload.SaveToSession()
			Products.Image_Thumb.CurrentValue = Products.Image_Thumb.Upload.FileName
		End If
		Products.fImage.Upload.Index = ObjForm.Index
		Call Products.fImage.Upload.RestoreDbFromSession()
		If confirmPage Then ' Post from confirm page
			Call Products.fImage.Upload.RestoreFromSession()
		Else
			If Products.fImage.Upload.UploadFile() Then

				' No action required
			Else
				Response.Write Products.fImage.Upload.Message
				Page_Terminate("")
				Response.End
			End If
			Call Products.fImage.Upload.SaveToSession()
			Products.fImage.CurrentValue = Products.fImage.Upload.FileName
		End If
		Products.fImage_Thumb.Upload.Index = ObjForm.Index
		Call Products.fImage_Thumb.Upload.RestoreDbFromSession()
		If confirmPage Then ' Post from confirm page
			Call Products.fImage_Thumb.Upload.RestoreFromSession()
		Else
			If Products.fImage_Thumb.Upload.UploadFile() Then

				' No action required
			Else
				Response.Write Products.fImage_Thumb.Upload.Message
				Page_Terminate("")
				Response.End
			End If
			Call Products.fImage_Thumb.Upload.SaveToSession()
			Products.fImage_Thumb.CurrentValue = Products.fImage_Thumb.Upload.FileName
		End If
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		Products.Description.CurrentValue = Null
		Products.Description.OldValue = Products.Description.CurrentValue
		Products.Price.CurrentValue = Null
		Products.Price.OldValue = Products.Price.CurrentValue
		Products.Active.CurrentValue = Null
		Products.Active.OldValue = Products.Active.CurrentValue
		Products.Image.Upload.DbValue = Null
		Products.Image.OldValue = Products.Image.Upload.DbValue
		Products.Image.CurrentValue = Null ' Clear file related field
		Products.Sizes.CurrentValue = Null
		Products.Sizes.OldValue = Products.Sizes.CurrentValue
		Products.Image_Thumb.Upload.DbValue = Null
		Products.Image_Thumb.OldValue = Products.Image_Thumb.Upload.DbValue
		Products.Image_Thumb.CurrentValue = Null ' Clear file related field
		Products.ProductName.CurrentValue = Null
		Products.ProductName.OldValue = Products.ProductName.CurrentValue
		Products.ItemNo.CurrentValue = Null
		Products.ItemNo.OldValue = Products.ItemNo.CurrentValue
		Products.UPC.CurrentValue = Null
		Products.UPC.OldValue = Products.UPC.CurrentValue
		Products.fDescription.CurrentValue = Null
		Products.fDescription.OldValue = Products.fDescription.CurrentValue
		Products.fImage.Upload.DbValue = Null
		Products.fImage.OldValue = Products.fImage.Upload.DbValue
		Products.fImage.CurrentValue = Null ' Clear file related field
		Products.fSizes.CurrentValue = Null
		Products.fSizes.OldValue = Products.fSizes.CurrentValue
		Products.fImage_Thumb.Upload.DbValue = Null
		Products.fImage_Thumb.OldValue = Products.fImage_Thumb.Upload.DbValue
		Products.fImage_Thumb.CurrentValue = Null ' Clear file related field
		Products.fProductName.CurrentValue = Null
		Products.fProductName.OldValue = Products.fProductName.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		Call GetUploadFiles() ' Get upload files
		If Not Products.Description.FldIsDetailKey Then Products.Description.FormValue = ObjForm.GetValue("x_Description")
		If Not Products.Price.FldIsDetailKey Then Products.Price.FormValue = ObjForm.GetValue("x_Price")
		If Not Products.Active.FldIsDetailKey Then Products.Active.FormValue = ObjForm.GetValue("x_Active")
		If Not Products.Sizes.FldIsDetailKey Then Products.Sizes.FormValue = ObjForm.GetValue("x_Sizes")
		If Not Products.ProductName.FldIsDetailKey Then Products.ProductName.FormValue = ObjForm.GetValue("x_ProductName")
		If Not Products.ItemNo.FldIsDetailKey Then Products.ItemNo.FormValue = ObjForm.GetValue("x_ItemNo")
		If Not Products.UPC.FldIsDetailKey Then Products.UPC.FormValue = ObjForm.GetValue("x_UPC")
		If Not Products.fDescription.FldIsDetailKey Then Products.fDescription.FormValue = ObjForm.GetValue("x_fDescription")
		If Not Products.fSizes.FldIsDetailKey Then Products.fSizes.FormValue = ObjForm.GetValue("x_fSizes")
		If Not Products.fProductName.FldIsDetailKey Then Products.fProductName.FormValue = ObjForm.GetValue("x_fProductName")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		Products.Description.CurrentValue = Products.Description.FormValue
		Products.Price.CurrentValue = Products.Price.FormValue
		Products.Active.CurrentValue = Products.Active.FormValue
		Products.Sizes.CurrentValue = Products.Sizes.FormValue
		Products.ProductName.CurrentValue = Products.ProductName.FormValue
		Products.ItemNo.CurrentValue = Products.ItemNo.FormValue
		Products.UPC.CurrentValue = Products.UPC.FormValue
		Products.fDescription.CurrentValue = Products.fDescription.FormValue
		Products.fSizes.CurrentValue = Products.fSizes.FormValue
		Products.fProductName.CurrentValue = Products.fProductName.FormValue
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

			' Image
			Products.Image.LinkCustomAttributes = ""
			Products.Image.HrefValue = ""
			Products.Image.TooltipValue = ""

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

			' fDescription
			Products.fDescription.LinkCustomAttributes = ""
			Products.fDescription.HrefValue = ""
			Products.fDescription.TooltipValue = ""

			' fImage
			Products.fImage.LinkCustomAttributes = ""
			Products.fImage.HrefValue = ""
			Products.fImage.TooltipValue = ""

			' fSizes
			Products.fSizes.LinkCustomAttributes = ""
			Products.fSizes.HrefValue = ""
			Products.fSizes.TooltipValue = ""

			' fImage_Thumb
			Products.fImage_Thumb.LinkCustomAttributes = ""
			Products.fImage_Thumb.HrefValue = ""
			Products.fImage_Thumb.TooltipValue = ""

			' fProductName
			Products.fProductName.LinkCustomAttributes = ""
			Products.fProductName.HrefValue = ""
			Products.fProductName.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf Products.RowType = EW_ROWTYPE_ADD Then ' Add row

			' Description
			Products.Description.EditCustomAttributes = ""
			Products.Description.EditValue = ew_HtmlEncode(Products.Description.CurrentValue)

			' Price
			Products.Price.EditCustomAttributes = ""
			Products.Price.EditValue = ew_HtmlEncode(Products.Price.CurrentValue)

			' Active
			Products.Active.EditCustomAttributes = ""
			Redim arwrk(1, 1)
			arwrk(0, 0) = "1"
			arwrk(1, 0) = ew_IIf(Products.Active.FldTagCaption(1) <> "", Products.Active.FldTagCaption(1), "Yes")
			arwrk(0, 1) = "0"
			arwrk(1, 1) = ew_IIf(Products.Active.FldTagCaption(2) <> "", Products.Active.FldTagCaption(2), "No")
			Products.Active.EditValue = arwrk

			' Image
			Products.Image.EditCustomAttributes = ""
			If Not ew_Empty(Products.Image.Upload.DbValue) Then
				Products.Image.EditValue = Products.Image.Upload.DbValue
				Products.Image.ImageAlt = Products.Image.FldAlt
			Else
				Products.Image.EditValue = ""
			End If

			' Sizes
			Products.Sizes.EditCustomAttributes = ""
			Products.Sizes.EditValue = ew_HtmlEncode(Products.Sizes.CurrentValue)

			' Image_Thumb
			Products.Image_Thumb.EditCustomAttributes = ""
			If Not ew_Empty(Products.Image_Thumb.Upload.DbValue) Then
				Products.Image_Thumb.EditValue = Products.Image_Thumb.Upload.DbValue
				Products.Image_Thumb.ImageAlt = Products.Image_Thumb.FldAlt
			Else
				Products.Image_Thumb.EditValue = ""
			End If

			' ProductName
			Products.ProductName.EditCustomAttributes = ""
			Products.ProductName.EditValue = ew_HtmlEncode(Products.ProductName.CurrentValue)

			' ItemNo
			Products.ItemNo.EditCustomAttributes = ""
			Products.ItemNo.EditValue = ew_HtmlEncode(Products.ItemNo.CurrentValue)

			' UPC
			Products.UPC.EditCustomAttributes = ""
			Products.UPC.EditValue = ew_HtmlEncode(Products.UPC.CurrentValue)

			' fDescription
			Products.fDescription.EditCustomAttributes = ""
			Products.fDescription.EditValue = ew_HtmlEncode(Products.fDescription.CurrentValue)

			' fImage
			Products.fImage.EditCustomAttributes = ""
			If Not ew_Empty(Products.fImage.Upload.DbValue) Then
				Products.fImage.EditValue = Products.fImage.Upload.DbValue
				Products.fImage.ImageAlt = Products.fImage.FldAlt
			Else
				Products.fImage.EditValue = ""
			End If

			' fSizes
			Products.fSizes.EditCustomAttributes = ""
			Products.fSizes.EditValue = ew_HtmlEncode(Products.fSizes.CurrentValue)

			' fImage_Thumb
			Products.fImage_Thumb.EditCustomAttributes = ""
			If Not ew_Empty(Products.fImage_Thumb.Upload.DbValue) Then
				Products.fImage_Thumb.EditValue = Products.fImage_Thumb.Upload.DbValue
				Products.fImage_Thumb.ImageAlt = Products.fImage_Thumb.FldAlt
			Else
				Products.fImage_Thumb.EditValue = ""
			End If

			' fProductName
			Products.fProductName.EditCustomAttributes = ""
			Products.fProductName.EditValue = ew_HtmlEncode(Products.fProductName.CurrentValue)

			' Edit refer script
			' Description

			Products.Description.HrefValue = ""

			' Price
			Products.Price.HrefValue = ""

			' Active
			Products.Active.HrefValue = ""

			' Image
			Products.Image.HrefValue = ""

			' Sizes
			Products.Sizes.HrefValue = ""

			' Image_Thumb
			Products.Image_Thumb.HrefValue = ""

			' ProductName
			Products.ProductName.HrefValue = ""

			' ItemNo
			Products.ItemNo.HrefValue = ""

			' UPC
			Products.UPC.HrefValue = ""

			' fDescription
			Products.fDescription.HrefValue = ""

			' fImage
			Products.fImage.HrefValue = ""

			' fSizes
			Products.fSizes.HrefValue = ""

			' fImage_Thumb
			Products.fImage_Thumb.HrefValue = ""

			' fProductName
			Products.fProductName.HrefValue = ""
		End If
		If Products.RowType = EW_ROWTYPE_ADD Or Products.RowType = EW_ROWTYPE_EDIT Or Products.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call Products.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If Products.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Products.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""
		If Not ew_CheckFileType(Products.Image.Upload.FileName) Then
			Call ew_AddMessage(gsFormError, Language.Phrase("WrongFileType"))
		End If
		If Products.Image.Upload.FileSize > 0 And CLng(EW_MAX_FILE_SIZE) > 0 Then
			If Products.Image.Upload.FileSize > CLng(EW_MAX_FILE_SIZE) Then
				Call ew_AddMessage(gsFormError, Replace(Language.Phrase("MaxFileSize"), "%s", EW_MAX_FILE_SIZE))
			End If
		End If
		If Not ew_CheckFileType(Products.Image_Thumb.Upload.FileName) Then
			Call ew_AddMessage(gsFormError, Language.Phrase("WrongFileType"))
		End If
		If Products.Image_Thumb.Upload.FileSize > 0 And CLng(EW_MAX_FILE_SIZE) > 0 Then
			If Products.Image_Thumb.Upload.FileSize > CLng(EW_MAX_FILE_SIZE) Then
				Call ew_AddMessage(gsFormError, Replace(Language.Phrase("MaxFileSize"), "%s", EW_MAX_FILE_SIZE))
			End If
		End If
		If Not ew_CheckFileType(Products.fImage.Upload.FileName) Then
			Call ew_AddMessage(gsFormError, Language.Phrase("WrongFileType"))
		End If
		If Products.fImage.Upload.FileSize > 0 And CLng(EW_MAX_FILE_SIZE) > 0 Then
			If Products.fImage.Upload.FileSize > CLng(EW_MAX_FILE_SIZE) Then
				Call ew_AddMessage(gsFormError, Replace(Language.Phrase("MaxFileSize"), "%s", EW_MAX_FILE_SIZE))
			End If
		End If
		If Not ew_CheckFileType(Products.fImage_Thumb.Upload.FileName) Then
			Call ew_AddMessage(gsFormError, Language.Phrase("WrongFileType"))
		End If
		If Products.fImage_Thumb.Upload.FileSize > 0 And CLng(EW_MAX_FILE_SIZE) > 0 Then
			If Products.fImage_Thumb.Upload.FileSize > CLng(EW_MAX_FILE_SIZE) Then
				Call ew_AddMessage(gsFormError, Replace(Language.Phrase("MaxFileSize"), "%s", EW_MAX_FILE_SIZE))
			End If
		End If

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
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
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear

		' Add new record
		sFilter = "(0 = 1)"
		Products.CurrentFilter = sFilter
		sSql = Products.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field Description
		Call Products.Description.SetDbValue(Rs, Products.Description.CurrentValue, Null, False)

		' Field Price
		Call Products.Price.SetDbValue(Rs, Products.Price.CurrentValue, Null, False)

		' Field Active
		boolwrk = Products.Active.CurrentValue
		If boolwrk&"" <> "1" And boolwrk&"" <> "0" Then boolwrk = ew_IIf(boolwrk&"" <> "", "1", "0")
		Call Products.Active.SetDbValue(Rs, boolwrk, Null, False)

		' Field Image
		If Products.Image.Upload.Action = "1" Then ' Keep
			If Not IsNull(RsOld) And Not IsEmpty(RsOld) Then
				Rs("Image") = RsOld("Image")
			End If
		ElseIf Products.Image.Upload.Action = "2" Or Products.Image.Upload.Action = "3" Then ' Update/Remove
		Products.Image.Upload.DbValue = Rs("Image") ' Get original value
		If IsNull(Products.Image.Upload.Value) Then
			Rs("Image") = Null
		Else
			Rs("Image") = ew_UploadFileNameEx(ew_UploadPathEx(True, Products.Image.UploadPath), Products.Image.Upload.FileName)
		End If
		End If

		' Field Sizes
		Call Products.Sizes.SetDbValue(Rs, Products.Sizes.CurrentValue, Null, False)

		' Field Image_Thumb
		If Products.Image_Thumb.Upload.Action = "1" Then ' Keep
			If Not IsNull(RsOld) And Not IsEmpty(RsOld) Then
				Rs("Image_Thumb") = RsOld("Image_Thumb")
			End If
		ElseIf Products.Image_Thumb.Upload.Action = "2" Or Products.Image_Thumb.Upload.Action = "3" Then ' Update/Remove
		Products.Image_Thumb.Upload.DbValue = Rs("Image_Thumb") ' Get original value
		If IsNull(Products.Image_Thumb.Upload.Value) Then
			Rs("Image_Thumb") = Null
		Else
			Rs("Image_Thumb") = ew_UploadFileNameEx(ew_UploadPathEx(True, Products.Image_Thumb.UploadPath), Products.Image_Thumb.Upload.FileName)
		End If
		End If

		' Field ProductName
		Call Products.ProductName.SetDbValue(Rs, Products.ProductName.CurrentValue, Null, False)

		' Field ItemNo
		Call Products.ItemNo.SetDbValue(Rs, Products.ItemNo.CurrentValue, Null, False)

		' Field UPC
		Call Products.UPC.SetDbValue(Rs, Products.UPC.CurrentValue, Null, False)

		' Field fDescription
		Call Products.fDescription.SetDbValue(Rs, Products.fDescription.CurrentValue, Null, False)

		' Field fImage
		If Products.fImage.Upload.Action = "1" Then ' Keep
			If Not IsNull(RsOld) And Not IsEmpty(RsOld) Then
				Rs("fImage") = RsOld("fImage")
			End If
		ElseIf Products.fImage.Upload.Action = "2" Or Products.fImage.Upload.Action = "3" Then ' Update/Remove
		Products.fImage.Upload.DbValue = Rs("fImage") ' Get original value
		If IsNull(Products.fImage.Upload.Value) Then
			Rs("fImage") = Null
		Else
			Rs("fImage") = ew_UploadFileNameEx(ew_UploadPathEx(True, Products.fImage.UploadPath), Products.fImage.Upload.FileName)
		End If
		End If

		' Field fSizes
		Call Products.fSizes.SetDbValue(Rs, Products.fSizes.CurrentValue, Null, False)

		' Field fImage_Thumb
		If Products.fImage_Thumb.Upload.Action = "1" Then ' Keep
			If Not IsNull(RsOld) And Not IsEmpty(RsOld) Then
				Rs("fImage_Thumb") = RsOld("fImage_Thumb")
			End If
		ElseIf Products.fImage_Thumb.Upload.Action = "2" Or Products.fImage_Thumb.Upload.Action = "3" Then ' Update/Remove
		Products.fImage_Thumb.Upload.DbValue = Rs("fImage_Thumb") ' Get original value
		If IsNull(Products.fImage_Thumb.Upload.Value) Then
			Rs("fImage_Thumb") = Null
		Else
			Rs("fImage_Thumb") = ew_UploadFileNameEx(ew_UploadPathEx(True, Products.fImage_Thumb.UploadPath), Products.fImage_Thumb.Upload.FileName)
		End If
		End If

		' Field fProductName
		Call Products.fProductName.SetDbValue(Rs, Products.fProductName.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = Products.Row_Inserting(RsOld, Rs)
		If bInsertRow Then

			' Field Image
			If Not ew_Empty(Products.Image.Upload.Value) Then
				If Products.Image.Upload.FileName = Products.Image.Upload.DbValue Then ' Overwrite if same file name
					Products.Image.Upload.SaveToFile Products.Image.UploadPath, Rs("Image"), True
					Products.Image.Upload.DbValue = "" ' No need to delete any more
				Else
					Products.Image.Upload.SaveToFile Products.Image.UploadPath, Rs("Image"), False
				End If
			End If
			If Products.Image.Upload.Action = "2" Or Products.Image.Upload.Action = "3" Then ' Update/Remove
				If Products.Image.Upload.DbValue <> "" Then ew_DeleteFile ew_UploadPathEx(True, Products.Image.UploadPath) & Products.Image.Upload.DbValue
			End If

			' Field Image_Thumb
			If Not ew_Empty(Products.Image_Thumb.Upload.Value) Then
				If Products.Image_Thumb.Upload.FileName = Products.Image_Thumb.Upload.DbValue Then ' Overwrite if same file name
					Products.Image_Thumb.Upload.SaveToFile Products.Image_Thumb.UploadPath, Rs("Image_Thumb"), True
					Products.Image_Thumb.Upload.DbValue = "" ' No need to delete any more
				Else
					Products.Image_Thumb.Upload.SaveToFile Products.Image_Thumb.UploadPath, Rs("Image_Thumb"), False
				End If
			End If
			If Products.Image_Thumb.Upload.Action = "2" Or Products.Image_Thumb.Upload.Action = "3" Then ' Update/Remove
				If Products.Image_Thumb.Upload.DbValue <> "" Then ew_DeleteFile ew_UploadPathEx(True, Products.Image_Thumb.UploadPath) & Products.Image_Thumb.Upload.DbValue
			End If

			' Field fImage
			If Not ew_Empty(Products.fImage.Upload.Value) Then
				If Products.fImage.Upload.FileName = Products.fImage.Upload.DbValue Then ' Overwrite if same file name
					Products.fImage.Upload.SaveToFile Products.fImage.UploadPath, Rs("fImage"), True
					Products.fImage.Upload.DbValue = "" ' No need to delete any more
				Else
					Products.fImage.Upload.SaveToFile Products.fImage.UploadPath, Rs("fImage"), False
				End If
			End If
			If Products.fImage.Upload.Action = "2" Or Products.fImage.Upload.Action = "3" Then ' Update/Remove
				If Products.fImage.Upload.DbValue <> "" Then ew_DeleteFile ew_UploadPathEx(True, Products.fImage.UploadPath) & Products.fImage.Upload.DbValue
			End If

			' Field fImage_Thumb
			If Not ew_Empty(Products.fImage_Thumb.Upload.Value) Then
				If Products.fImage_Thumb.Upload.FileName = Products.fImage_Thumb.Upload.DbValue Then ' Overwrite if same file name
					Products.fImage_Thumb.Upload.SaveToFile Products.fImage_Thumb.UploadPath, Rs("fImage_Thumb"), True
					Products.fImage_Thumb.Upload.DbValue = "" ' No need to delete any more
				Else
					Products.fImage_Thumb.Upload.SaveToFile Products.fImage_Thumb.UploadPath, Rs("fImage_Thumb"), False
				End If
			End If
			If Products.fImage_Thumb.Upload.Action = "2" Or Products.fImage_Thumb.Upload.Action = "3" Then ' Update/Remove
				If Products.fImage_Thumb.Upload.DbValue <> "" Then ew_DeleteFile ew_UploadPathEx(True, Products.fImage_Thumb.UploadPath) & Products.fImage_Thumb.Upload.DbValue
			End If

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
		Else
			Rs.CancelUpdate
			If Products.CancelMessage <> "" Then
				FailureMessage = Products.CancelMessage
				Products.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			Products.ItemId.DbValue = RsNew("ItemId")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call Products.Row_Inserted(RsOld, RsNew)
		End If

		' Field Image
		Call Products.Image.Upload.RemoveFromSession() ' Remove file value from Session

		' Field Image_Thumb
		Call Products.Image_Thumb.Upload.RemoveFromSession() ' Remove file value from Session

		' Field fImage
		Call Products.fImage.Upload.RemoveFromSession() ' Remove file value from Session

		' Field fImage_Thumb
		Call Products.fImage_Thumb.Upload.RemoveFromSession() ' Remove file value from Session
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

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
