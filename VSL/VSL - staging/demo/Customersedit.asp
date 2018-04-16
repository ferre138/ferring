<%
Const EW_PAGE_ID = "edit"
Const EW_TABLE_NAME = "Customers"
%>
<!--#include file="ewcfg60.asp"--> 
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%

' Open connection to the database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
%>
<%
Dim Security
Set Security = New cAdvancedSecurity
%>
<%
If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
%>
<%

' Common page loading event (in userfn60.asp)
Call Page_Loading()
%>
<%

' Page load event, used in current page
Call Page_Load()
%>
<%
Response.Buffer = True

' Load key from QueryString
'If Request.QueryString("CustomerID").Count > 0 Then
'	Customers.CustomerID.QueryStringValue = Request.QueryString("CustomerID")
'End If
Customers.CustomerID.QueryStringValue = Security.CurrentUserID
' Create form object
Dim objForm
Set objForm = New cFormObj
If objForm.GetValue("a_edit")&"" <> "" Then
	Customers.CurrentAction = objForm.GetValue("a_edit") ' Get action code
	Call LoadFormValues() ' Get form values
Else
	Customers.CurrentAction = "I" ' Default action is display
End If

' Close form object
Set objForm = Nothing

' Check if valid key
If Customers.CustomerID.CurrentValue = "" Then Call Page_Terminate(Customers.ReturnUrl) ' Invalid key, exit
Select Case Customers.CurrentAction
	Case "I" ' Get a record to display
		If Not LoadRow() Then ' Load Record based on key
			Session(EW_SESSION_MESSAGE) = "No records found" ' No record found
			Call Page_Terminate(Customers.ReturnUrl) ' Return to caller
		End If
	Case "U" ' Update
		Customers.SendEmail = True ' Send email on update success
		If EditRow() Then ' Update Record based on key
			Session(EW_SESSION_MESSAGE) = "Update successful" ' Update success

			Call Page_Terminate(Customers.ReturnUrl) ' Return to caller
		Else
			Call RestoreFormValues() ' Restore form values if update failed
		End If
End Select

' Render the record
Customers.RowType = EW_ROWTYPE_EDIT ' Render as edit
Call RenderRow()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
var EW_PAGE_ID = "edit"; // Page id
//-->
</script>
<script type="text/javascript">
<!--

function ew_ValidateForm(fobj) {
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = (fobj.key_count) ? Number(fobj.key_count.value) : 1;
	for (i=0; i<rowcnt; i++) {
		infix = (fobj.key_count) ? String(i+1) : "";
		elm = fobj.elements["x" + infix + "_Inv_FirstName"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - First Name"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_Inv_LastName"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Last Name"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_Inv_Address"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Billing Address"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_City"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - City"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_Province"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Province"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_PostalCode"];
		if (elm && !isValidPostal(elm.value)) {
			if (!ew_OnError(elm, "Please enter a valid Postal Code"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_PhoneNumber"];
		if (elm && !checkInternationalPhone(elm.value)) {
			if (!ew_OnError(elm, "Incorrect phone number - Phone Number"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_EmailAddress"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Email Address"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_EmailAddress"];
		if (elm && !ew_CheckEmail(elm.value)) {
			if (!ew_OnError(elm, "Please Enter a Valid Email"))
				return false;
		}
	
	}
	return true;
}

function isValidPostal(postalCode){
postalCode=postalCode.replace(" ","");
var pc = /[^a-zA-Z0-9]/g;
if(postalCode.match(pc)!=null){
return false;
}
else if(postalCode.length!=6){
return false;
}
else{
return true;
}
}

// Declaring required variables
var digits = "0123456789";
// non-digit characters which are allowed in phone numbers
var phoneNumberDelimiters = "()- ";
// characters which are allowed in international phone numbers
// (a leading + is OK)
var validWorldPhoneChars = phoneNumberDelimiters + "+";
// Minimum no of digits in an international phone no.
var minDigitsInIPhoneNumber = 10;

function isInteger(s)
{   var i;
    for (i = 0; i < s.length; i++)
    {   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag)
{   var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++)
    {   
        // Check that current character isn't whitespace.
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function checkInternationalPhone(strPhone){
s=stripCharsInBag(strPhone,validWorldPhoneChars);
return (isInteger(s) && s.length >= minDigitsInIPhoneNumber);
}

//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script type="text/javascript">
<!--
var ew_MultiPagePage = "Page"; // multi-page Page Text
var ew_MultiPageOf = "of"; // multi-page Of Text
var ew_MultiPagePrev = "Prev"; // multi-page Prev Text
var ew_MultiPageNext = "Next"; // multi-page Next Text
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
// To include another .js script, use:
// ew_ClientScriptInclude("my_javascript.js"); 
//-->
</script><table width="820"  border="0" cellpadding="0" cellspacing="0" id="Table_01">
            <tr>
            <td width="680" rowspan="2"><img src="images/title_account.png" width="410" height="75"></td>
              <!--<td align="right" valign="top"><img src="images/fontsize.png" border="0" alt=""></td>
              <td width="24" valign="top"> <a href="#"
				onmouseover="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseout="changeImages('login_13', 'images/font1.png'); return true;"
				onmousedown="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseup="changeImages('login_13', 'images/login_13-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('default'); 
return false;"> <img name="login_13" src="images/font1.png" width="24" height="27" border="0" alt=""></a></td>
              <td width="24"  valign="top"> <a href="#"
				onmouseover="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseout="changeImages('login_14', 'images/font2.png'); return true;"
				onmousedown="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseup="changeImages('login_14', 'images/login_14-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Medium'); 
return false;"> <img name="login_14" src="images/font2.png" width="24" height="27" border="0" alt=""></a></td>
              <td width="24"  valign="top"> <a href="#"
				onmouseover="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseout="changeImages('login_15', 'images/font3.png'); return true;"
				onmousedown="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseup="changeImages('login_15', 'images/login_15-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Large'); 
return false;"><img name="login_15" src="images/font3.png" width="24" height="27" border="0" alt=""></a></td>
            </tr>
            <tr>
              <td colspan="4" align="right" valign="top"><div align="right">
                <p><a href="french/Customersedit.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
              </div></td>-->
              </tr>
        </table>
<%
If Session(EW_SESSION_MESSAGE) <> "" Then
%>
<p><span class="ewmsg"><%= Session(EW_SESSION_MESSAGE) %></span></p>
<%
	Session(EW_SESSION_MESSAGE) = "" ' Clear message
End If
%>
<form name="fCustomersedit" id="fCustomersedit" action="Customersedit.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<p>
<input type="hidden" name="a_edit" id="a_edit" value="U">
<div class="t">
  <div align="right" class="vslcss"><a href="VSLOrderForm.asp">Back to Products</a> : <a href="vslCart.asp">View Cart</a> :
		    <%

	if (Not Security.IsLoggedIn()) then%>
	
            <a href="login.asp">login</a>
            <%else%>
			 <a href="changepwd.asp">Change Password</a> :
            <a href="logout.asp">logout</a>
            <%end if
	set Security =nothing %> 
		    </div>
  <table width="800" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center" valign="top"><table width="400" class="ewTable">
	<tr class="ewTableAltRow">
	  <td colspan="2" class="ewTableHeader">Address</td>
	  </tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">First Name<span class='ewmsg'>&nbsp;*<span id="cb_x_CustomerID">
		  <input type="hidden" name="x_CustomerID" id="x_CustomerID" value="<%= Server.HTMLEncode(Customers.CustomerID.CurrentValue&"") %>">
		</span></span></td>
		<td<%= Customers.Inv_FirstName.CellAttributes %>><span id="cb_x_Inv_FirstName">
<input type="text" name="x_Inv_FirstName" id="x_Inv_FirstName" title="" size="30" maxlength="30" value="<%= Customers.Inv_FirstName.EditValue %>"<%= Customers.Inv_FirstName.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableRow">
		<td class="ewTableHeader">Last Name<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.Inv_LastName.CellAttributes %>><span id="cb_x_Inv_LastName">
<input type="text" name="x_Inv_LastName" id="x_Inv_LastName" title="" size="30" maxlength="50" value="<%= Customers.Inv_LastName.EditValue %>"<%= Customers.Inv_LastName.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">Billing Address<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.Inv_Address.CellAttributes %>><span id="cb_x_Inv_Address">
<input type="text" name="x_Inv_Address" id="x_Inv_Address" title="" size="30" maxlength="255" value="<%= Customers.Inv_Address.EditValue %>"<%= Customers.Inv_Address.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableRow">
		<td class="ewTableHeader">Address 2</td>
		<td<%= Customers.Inv_Address2.CellAttributes %>><span id="cb_x_Inv_Address2">
<input type="text" name="x_Inv_Address2" id="x_Inv_Address2" title="" size="30" maxlength="255" value="<%= Customers.Inv_Address2.EditValue %>"<%= Customers.Inv_Address2.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">City<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.inv_City.CellAttributes %>><span id="cb_x_inv_City">
<input type="text" name="x_inv_City" id="x_inv_City" title="" size="30" maxlength="50" value="<%= Customers.inv_City.EditValue %>"<%= Customers.inv_City.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableRow">
		<td class="ewTableHeader">Province<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.inv_Province.CellAttributes %>><span id="cb_x_inv_Province">
<select id='x_inv_Province' name='x_inv_Province'<%= Customers.inv_Province.EditAttributes %>>
<!--option value=''>Please Select</option-->
<%
If IsArray(Customers.inv_Province.EditValue) Then
	arwrk = Customers.inv_Province.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Customers.inv_Province.CurrentValue&"" Then
			selwrk = " selected"
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
</span></td>
	</tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">Postal Code<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.inv_PostalCode.CellAttributes %>><span id="cb_x_inv_PostalCode">
<input type="text" name="x_inv_PostalCode" id="x_inv_PostalCode" title="" size="30" maxlength="20" value="<%= Customers.inv_PostalCode.EditValue %>"<%= Customers.inv_PostalCode.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableRow">
		<td class="ewTableHeader">Country</td>
		<td<%= Customers.inv_Country.CellAttributes %>><span id="cb_x_inv_Country">
<input type="text"  name="x_inv_Country" id="x_inv_Country" title="" size="30" maxlength="50" value="<%= Customers.inv_Country.EditValue %>"<%= Customers.inv_Country.EditAttributes %>>
</span></td>
	</tr></table></td>
      <td align="center" valign="top"><table width="400" class="ewTable">
	<tr class="ewTableAltRow">
	  <td colspan="2" class="ewTableHeader">Contact info </td>
	  </tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">Phone Number</td>
		<td<%= Customers.inv_PhoneNumber.CellAttributes %>><span id="cb_x_inv_PhoneNumber">
<input type="text" name="x_inv_PhoneNumber" id="x_inv_PhoneNumber" title="" size="30" maxlength="30" value="<%= Customers.inv_PhoneNumber.EditValue %>"<%= Customers.inv_PhoneNumber.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableRow">
		<td class="ewTableHeader">Email Address<span class='ewmsg'>&nbsp;*</span></td>
		<td<%= Customers.inv_EmailAddress.CellAttributes %>><span id="cb_x_inv_EmailAddress">
<input type="text" name="x_inv_EmailAddress" id="x_inv_EmailAddress" title="" size="30" maxlength="50" value="<%= Customers.inv_EmailAddress.EditValue %>"<%= Customers.inv_EmailAddress.EditAttributes %>>
</span></td>
	</tr>
	<tr class="ewTableAltRow">
		<td class="ewTableHeader">Fax</td>
		<td<%= Customers.inv_Fax.CellAttributes %>><span id="cb_x_inv_Fax">
<input type="text" name="x_inv_Fax" id="x_inv_Fax" title="" size="30" maxlength="30" value="<%= Customers.inv_Fax.EditValue %>"<%= Customers.inv_Fax.EditAttributes %>>
</span></td>
	</tr>
</table>
        <p>&nbsp;</p>
        <p>&nbsp;          </p>
        <p>
          <input type="submit" name="btnAction" id="btnAction" value="   Save Changes   ">
        </p></td>
    </tr>
  </table>
</div>
<p>

</form>
<!--#include file="footer.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<%

' If control is passed here, simply terminate the page without redirect
Call Page_Terminate("")

' -----------------------------------------------------------------
'  Subroutine Page_Terminate
'  - called when exit page
'  - clean up ADO connection and objects
'  - if url specified, redirect to url, otherwise end response
'
Sub Page_Terminate(url)

	' Page unload event, used in current page
	Call Page_Unload()

	' Global page unloaded event (in userfn60.asp)
	Call Page_Unloaded()
	conn.Close ' Close Connection
	Set conn = Nothing
	Set Security = Nothing
	Set Customers = Nothing

	' Go to url if specified
	If url <> "" Then
		Response.Clear
		Response.Redirect url
	End If

	' Terminate response
	Response.End
End Sub

'
'  Subroutine Page_Terminate (End)
' ----------------------------------------

%>
<%

' Load form values
Function LoadFormValues()

	' Load from form
	Customers.CustomerID.FormValue = objForm.GetValue("x_CustomerID")
	Customers.Inv_FirstName.FormValue = objForm.GetValue("x_Inv_FirstName")
	Customers.Inv_LastName.FormValue = objForm.GetValue("x_Inv_LastName")
	Customers.Inv_Address.FormValue = objForm.GetValue("x_Inv_Address")
	Customers.Inv_Address2.FormValue = objForm.GetValue("x_Inv_Address2")
	Customers.inv_City.FormValue = objForm.GetValue("x_inv_City")
	Customers.inv_Province.FormValue = objForm.GetValue("x_inv_Province")
	Customers.inv_PostalCode.FormValue = objForm.GetValue("x_inv_PostalCode")
	Customers.inv_Country.FormValue = objForm.GetValue("x_inv_Country")
	Customers.inv_PhoneNumber.FormValue = objForm.GetValue("x_inv_PhoneNumber")
	Customers.inv_EmailAddress.FormValue = objForm.GetValue("x_inv_EmailAddress")
	Customers.inv_Fax.FormValue = objForm.GetValue("x_inv_Fax")
End Function

' Restore form values
Function RestoreFormValues()
	Customers.CustomerID.CurrentValue = Customers.CustomerID.FormValue
	Customers.Inv_FirstName.CurrentValue = Customers.Inv_FirstName.FormValue
	Customers.Inv_LastName.CurrentValue = Customers.Inv_LastName.FormValue
	Customers.Inv_Address.CurrentValue = Customers.Inv_Address.FormValue
	Customers.Inv_Address2.CurrentValue = Customers.Inv_Address2.FormValue
	Customers.inv_City.CurrentValue = Customers.inv_City.FormValue
	Customers.inv_Province.CurrentValue = Customers.inv_Province.FormValue
	Customers.inv_PostalCode.CurrentValue = Customers.inv_PostalCode.FormValue
	Customers.inv_Country.CurrentValue = Customers.inv_Country.FormValue
	Customers.inv_PhoneNumber.CurrentValue = Customers.inv_PhoneNumber.FormValue
	Customers.inv_EmailAddress.CurrentValue = Customers.inv_EmailAddress.FormValue
	Customers.inv_Fax.CurrentValue = Customers.inv_Fax.FormValue
End Function
%>
<%

' Load row based on key values
Function LoadRow()
	Dim rs, sSql, sFilter
	sFilter = Customers.SqlKeyFilter
	If Not IsNumeric(Customers.CustomerID.CurrentValue) Then
		LoadRow = False ' Invalid key, exit
		Exit Function
	End If
	sFilter = Replace(sFilter, "@CustomerID@", ew_AdjustSql(Customers.CustomerID.CurrentValue)) ' Replace key value

	' Call Row Selecting event
	Call Customers.Row_Selecting(sFilter)

	' Load sql based on filter
	Customers.CurrentFilter = sFilter
	sSql = Customers.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadRow = False
	Else
		LoadRow = True
		rs.MoveFirst
		Call LoadRowValues(rs) ' Load row values

		' Call Row Selected event
		Call Customers.Row_Selected(rs)
	End If
	rs.Close
	Set rs = Nothing
End Function

' Load row values from recordset
Sub LoadRowValues(rs)
	Customers.CustomerID.DbValue = rs("CustomerID")
	Customers.Inv_FirstName.DbValue = rs("Inv_FirstName")
	Customers.Inv_LastName.DbValue = rs("Inv_LastName")
	Customers.Inv_Address.DbValue = rs("Inv_Address")
	Customers.Inv_Address2.DbValue = rs("Inv_Address2")
	Customers.inv_City.DbValue = rs("inv_City")
	Customers.inv_Province.DbValue = rs("inv_Province")
	Customers.inv_PostalCode.DbValue = rs("inv_PostalCode")
	Customers.inv_Country.DbValue = rs("inv_Country")
	Customers.inv_PhoneNumber.DbValue = rs("inv_PhoneNumber")
	Customers.inv_EmailAddress.DbValue = rs("inv_EmailAddress")
	Customers.inv_Fax.DbValue = rs("inv_Fax")
	Customers.Notes.DbValue = rs("Notes")
	Customers.UserName.DbValue = rs("UserName")
	Customers.passwrd.DbValue = rs("passwrd")
End Sub
%>
<%

' Render row values based on field settings
Sub RenderRow()

	' Call Row Rendering event
	Call Customers.Row_Rendering()

	' Common render codes for all row types
	' CustomerID

	Customers.CustomerID.CellCssStyle = ""
	Customers.CustomerID.CellCssClass = ""

	' Inv_FirstName
	Customers.Inv_FirstName.CellCssStyle = ""
	Customers.Inv_FirstName.CellCssClass = ""

	' Inv_LastName
	Customers.Inv_LastName.CellCssStyle = ""
	Customers.Inv_LastName.CellCssClass = ""

	' Inv_Address
	Customers.Inv_Address.CellCssStyle = ""
	Customers.Inv_Address.CellCssClass = ""

	' Inv_Address2
	Customers.Inv_Address2.CellCssStyle = ""
	Customers.Inv_Address2.CellCssClass = ""

	' inv_City
	Customers.inv_City.CellCssStyle = ""
	Customers.inv_City.CellCssClass = ""

	' inv_Province
	Customers.inv_Province.CellCssStyle = ""
	Customers.inv_Province.CellCssClass = ""

	' inv_PostalCode
	Customers.inv_PostalCode.CellCssStyle = ""
	Customers.inv_PostalCode.CellCssClass = ""

	' inv_Country
	Customers.inv_Country.CellCssStyle = ""
	Customers.inv_Country.CellCssClass = ""

	' inv_PhoneNumber
	Customers.inv_PhoneNumber.CellCssStyle = ""
	Customers.inv_PhoneNumber.CellCssClass = ""

	' inv_EmailAddress
	Customers.inv_EmailAddress.CellCssStyle = ""
	Customers.inv_EmailAddress.CellCssClass = ""

	' inv_Fax
	Customers.inv_Fax.CellCssStyle = ""
	Customers.inv_Fax.CellCssClass = ""
	If Customers.RowType = EW_ROWTYPE_VIEW Then ' View row
	ElseIf Customers.RowType = EW_ROWTYPE_ADD Then ' Add row
	ElseIf Customers.RowType = EW_ROWTYPE_EDIT Then ' Edit row

		' CustomerID
		Customers.CustomerID.EditCustomAttributes = ""
		Customers.CustomerID.EditValue = Customers.CustomerID.CurrentValue
		Customers.CustomerID.CssStyle = ""
		Customers.CustomerID.CssClass = ""
		Customers.CustomerID.ViewCustomAttributes = ""

		' Inv_FirstName
		Customers.Inv_FirstName.EditCustomAttributes = ""
		Customers.Inv_FirstName.EditValue = ew_HtmlEncode(Customers.Inv_FirstName.CurrentValue)

		' Inv_LastName
		Customers.Inv_LastName.EditCustomAttributes = ""
		Customers.Inv_LastName.EditValue = ew_HtmlEncode(Customers.Inv_LastName.CurrentValue)

		' Inv_Address
		Customers.Inv_Address.EditCustomAttributes = ""
		Customers.Inv_Address.EditValue = ew_HtmlEncode(Customers.Inv_Address.CurrentValue)

		' Inv_Address2
		Customers.Inv_Address2.EditCustomAttributes = ""
		Customers.Inv_Address2.EditValue = ew_HtmlEncode(Customers.Inv_Address2.CurrentValue)

		' inv_City
		Customers.inv_City.EditCustomAttributes = ""
		Customers.inv_City.EditValue = ew_HtmlEncode(Customers.inv_City.CurrentValue)

		' inv_Province
		Customers.inv_Province.EditCustomAttributes = ""
		sSqlWrk = "SELECT [Prov], [Province] FROM [Province]"
		sSqlWrk = sSqlWrk & " ORDER BY [Province] Asc"
		Set rswrk = Server.CreateObject("ADODB.Recordset")
		rswrk.Open sSqlWrk, conn
		If Not rswrk.Eof Then
			arwrk = rswrk.GetRows
		Else
			arwrk = ""
		End If
		rswrk.Close
		Set rswrk = Nothing
		arwrk = ew_AddItemToArray(arwrk, 0, Array("", "Please Select"))
		Customers.inv_Province.EditValue = arwrk

		' inv_PostalCode
		Customers.inv_PostalCode.EditCustomAttributes = ""
		Customers.inv_PostalCode.EditValue = ew_HtmlEncode(Customers.inv_PostalCode.CurrentValue)

		' inv_Country
		Customers.inv_Country.EditCustomAttributes = ""
		Customers.inv_Country.EditValue = ew_HtmlEncode(Customers.inv_Country.CurrentValue)

		' inv_PhoneNumber
		Customers.inv_PhoneNumber.EditCustomAttributes = ""
		Customers.inv_PhoneNumber.EditValue = ew_HtmlEncode(Customers.inv_PhoneNumber.CurrentValue)

		' inv_EmailAddress
		Customers.inv_EmailAddress.EditCustomAttributes = ""
		Customers.inv_EmailAddress.EditValue = ew_HtmlEncode(Customers.inv_EmailAddress.CurrentValue)

		' inv_Fax
		Customers.inv_Fax.EditCustomAttributes = ""
		Customers.inv_Fax.EditValue = ew_HtmlEncode(Customers.inv_Fax.CurrentValue)
	ElseIf Customers.RowType = EW_ROWTYPE_SEARCH Then ' Search row
	End If

	' Call Row Rendered event
	Call Customers.Row_Rendered()
End Sub
%>
<%

' Update record based on key values
Function EditRow()
	On Error Resume Next
	Dim rs, sSql, sFilter
	Dim rsChk, sSqlChk, sFilterChk
	Dim bUpdateRow
	Dim rsold, rsnew
	sFilter = Customers.SqlKeyFilter
	If Not IsNumeric(Customers.CustomerID.CurrentValue) Then
		EditRow = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@CustomerID@", ew_AdjustSql(Customers.CustomerID.CurrentValue)) ' Replace key value
	If Customers.UserName.CurrentValue <> "" Then ' Check field with unique index
		sFilterChk = "([UserName] = '" & ew_AdjustSql(Customers.UserName.CurrentValue) & "')"
		sFilterChk = sFilterChk & " AND NOT (" & sFilter & ")"
		Customers.CurrentFilter = sFilterChk
		sSqlChk = Customers.SQL
		Set rsChk = conn.Execute(sSqlChk)
		If Err.Number <> 0 Then
			Session(EW_SESSION_MESSAGE) = Err.Description
			rsChk.Close
			Set rsChk = Nothing
			EditRow = False
			Exit Function
		ElseIf Not rsChk.Eof Then
			Session(EW_SESSION_MESSAGE) = "Duplicate value for index or primary key -- [UserName], value = " & Customers.UserName.CurrentValue
			rsChk.Close
			Set rsChk = Nothing
			EditRow = False
			Exit Function
		End If
		rsChk.Close
		Set rsChk = Nothing
	End If
	If Security.CurrentUserID <> "" And Not Security.IsAdmin Then ' Non system admin
		sFilter = Customers.AddUserIDFilter(sFilter, Security.CurrentUserID) ' Add user id filter
		Customers.CurrentFilter = sFilter
	End If
	Customers.CurrentFilter  = sFilter
	sSql = Customers.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = EW_CURSORLOCATION
	rs.Open sSql, conn, 1, 2
	If Err.Number <> 0 Then
		Session(EW_SESSION_MESSAGE) = Err.Description
		rs.Close
		Set rs = Nothing
		EditRow = False
		Exit Function
	End If

	' Clone old rs object
	Set rsold = ew_CloneRs(rs)
	If rs.Eof Then
		EditRow = False ' Update Failed
	Else

		' Field CustomerID
		' Field Inv_FirstName

		Call Customers.Inv_FirstName.SetDbValue(Customers.Inv_FirstName.CurrentValue, Null)
		rs("Inv_FirstName") = Customers.Inv_FirstName.DbValue

		' Field Inv_LastName
		Call Customers.Inv_LastName.SetDbValue(Customers.Inv_LastName.CurrentValue, Null)
		rs("Inv_LastName") = Customers.Inv_LastName.DbValue

		' Field Inv_Address
		Call Customers.Inv_Address.SetDbValue(Customers.Inv_Address.CurrentValue, Null)
		rs("Inv_Address") = Customers.Inv_Address.DbValue

		' Field Inv_Address2
		Call Customers.Inv_Address2.SetDbValue(Customers.Inv_Address2.CurrentValue, Null)
		rs("Inv_Address2") = Customers.Inv_Address2.DbValue

		' Field inv_City
		Call Customers.inv_City.SetDbValue(Customers.inv_City.CurrentValue, Null)
		rs("inv_City") = Customers.inv_City.DbValue

		' Field inv_Province
		Call Customers.inv_Province.SetDbValue(Customers.inv_Province.CurrentValue, Null)
		rs("inv_Province") = Customers.inv_Province.DbValue

		' Field inv_PostalCode
		Call Customers.inv_PostalCode.SetDbValue(Customers.inv_PostalCode.CurrentValue, Null)
		rs("inv_PostalCode") = Customers.inv_PostalCode.DbValue

		' Field inv_Country
		Call Customers.inv_Country.SetDbValue(Customers.inv_Country.CurrentValue, Null)
		rs("inv_Country") = Customers.inv_Country.DbValue

		' Field inv_PhoneNumber
		Call Customers.inv_PhoneNumber.SetDbValue(Customers.inv_PhoneNumber.CurrentValue, Null)
		rs("inv_PhoneNumber") = Customers.inv_PhoneNumber.DbValue

		' Field inv_EmailAddress
		Call Customers.inv_EmailAddress.SetDbValue(Customers.inv_EmailAddress.CurrentValue, Null)
		rs("inv_EmailAddress") = Customers.inv_EmailAddress.DbValue

		' Field inv_Fax
		Call Customers.inv_Fax.SetDbValue(Customers.inv_Fax.CurrentValue, Null)
		rs("inv_Fax") = Customers.inv_Fax.DbValue

		' Check recordset update error
		If Err.Number <> 0 Then
			Session(EW_SESSION_MESSAGE) = Err.Description
			rs.Close
			Set rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Call Row Updating event
		bUpdateRow = Customers.Row_Updating(rsold, rs)
		If bUpdateRow Then

			' Clone new rs object
			Set rsnew = ew_CloneRs(rs)
			rs.Update
			If Err.Number <> 0 Then
				Session(EW_SESSION_MESSAGE) = Err.Description
				EditRow = False
			Else
				EditRow = True
			End If
		Else
			rs.CancelUpdate
			If Customers.CancelMessage <> "" Then
				Session(EW_SESSION_MESSAGE) = Customers.CancelMessage
				Customers.CancelMessage = ""
			Else
				Session(EW_SESSION_MESSAGE) = "Update cancelled"
			End If
			EditRow = False
		End If
	End If

	' Call Row Updated event
	If EditRow Then
		Call Customers.Row_Updated(rsold, rsnew)
	End If
	rs.Close
	Set rs = Nothing
	If IsObject(rsold) Then
		rsold.Close
		Set rsold = Nothing
	End If
	If IsObject(rsnew) Then
		rsnew.Close
		Set rsnew = Nothing
	End If
End Function
%>
<%

' Page Load event
Sub Page_Load()

'***Response.Write "Page Load"
End Sub

' Page Unload event
Sub Page_Unload()

'***Response.Write "Page Unload"
End Sub
function getCustomerId(c)
	dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")

	rs.Open  "SELECT Customers.CustomerID FROM Customers WHERE (Customers.UserName=""" & session("") & """);=" & ItemId & ") ;", c, 1, 2 
	if(not rs.eof) then 
		if(q>4) then
			getPrice = rs("Price_Rebate")
		else
			getPrice = rs("Price")
		end if
	else
		getprice=-1
	end if
rs.close
set rs=nothing

end function
%>
