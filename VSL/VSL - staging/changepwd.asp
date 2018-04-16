<%
Const EW_PAGE_ID = "changepwd"
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
If Not Security.IsLoggedIn() Or Security.IsSysAdmin() Then Call Page_Terminate("login.asp")
Call Security.LoadCurrentUserLevel("Customers")
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
Dim bValidPwd, bPwdUpdated, sUsername, sOPwd, sNPwd, sCPwd, sEmail, sFilter, sSql, rs
If Request.Form("submit") <> "" Then
	bValidPwd = False
	bPwdUpdated = False

	' Setup variables
	sUsername = Security.CurrentUserName
	sOPwd = Request.Form("opwd")
	sNPwd = Request.Form("npwd")
	sCPwd = Request.Form("cpwd")
	If sNPwd = sCPwd Then
		sFilter = "([UserName] = '" & ew_AdjustSql(sUsername) & "')"

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Customers class, Customersinfo.asp

		Customers.CurrentFilter = sFilter
		sSql = Customers.SQL
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSql, conn, 1, 2
		If Not rs.Eof Then
			If sOPwd = rs("passwrd") Then
				rs("passwrd") = sNPwd ' Change Password
				rs.Update
				bValidPwd = True
				bPwdUpdated = True
			End If
		End If
		rs.Close
		Set rs = Nothing
	End If
	If bPwdUpdated Then
		Session(EW_SESSION_MESSAGE) = "Password Changed" ' set up message
		Call Page_Terminate("default.asp") ' exit page and clean up
	End If
Else
	bValidPwd = True
End If
%>
<!--#include file="header.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
// To include another .js script, use:
// ew_ClientScriptInclude("my_javascript.js"); 
//-->
</script>
<script type="text/javascript">
<!-- start JavaScript

function  ew_ValidateForm(fobj) {
	if  (!ew_HasValue(fobj.opwd)) {
		if  (!ew_OnError(fobj.opwd, "Please enter password"))
			return false;
	}
	if  (!ew_HasValue(fobj.npwd)) {
		if  (!ew_OnError(fobj.npwd, "Please enter password"))
			return false;
	}
	if  (fobj.npwd.value != fobj.cpwd.value) {
		if  (!ew_OnError(fobj.cpwd, "Mismatch Password"))
			return false;
	}
	return true;
}
// end JavaScript -->
</script>

<div align="right">
  <table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
    <tr>
      <td width="680" rowspan="2"><img src="images/title_account.png" width="410" height="75"></td>
    <!--  <td width="28" valign="top"><img src="images/fontsize.png" alt="" width="78" height="27" border="0"> </td>
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
      <td width="26"  valign="top"> <a href="#"
				onmouseover="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseout="changeImages('login_15', 'images/font3.png'); return true;"
				onmousedown="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseup="changeImages('login_15', 'images/login_15-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Large'); 
return false;"><img name="login_15" src="images/font3.png" width="24" height="27" border="0" alt=""></a></td>
    </tr>
    <tr>
      <td colspan="4" valign="top"><div align="right">
        <p><a href="french/changepwd.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
        </div></td>-->
      </tr>
  </table>
     </div> 
<% If Not bValidPwd Then %>
<p><span class="ewmsg">Invalid Password</span></p>
<% End If %>
<div class="t">
  <div align="right"><span class="vslcss"><a href="VSLOrderForm.asp">Back to Products</a> :  <a href="vslCart.asp">View Cart</a> :
            <%

	if (Not Security.IsLoggedIn()) then%>
	      
          <a href="login.asp">Login</a>
            <%else%>
	        <a href="Customersedit.asp">Edit account</a> :
		     <a href="changepwd.asp">Change Password</a> :
             <a href="logout.asp">Logout</a>
             <%end if
	set Security =nothing %>
             <img src="images/spacer.gif" width="25" height="10">
        </span>
  </div>
  <form action="changepwd.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<p class="subheading">Change Password </p>
 <table width="384" border="0" cellpadding="4" cellspacing="0">
	<tr>
		<td height="40"><span class="vslCss">Old Password</span></td>
		<td><span class="vslCss"><input type="password" name="opwd" id="opwd" size="20"></span></td>
	</tr>
	<tr>
		<td height="40"><span class="vslCss">New Password</span></td>
		<td><span class="vslCss">
		  <input type="password" name="npwd" id="npwd" size="20">
		</span></td>
	</tr>
	<tr>
		<td height="40"><span class="vslCss">Confirm Password</span></td>
		<td><span class="vslCss"><input type="password" name="cpwd" id="cpwd" size="20"></span></td>
	</tr>
	<tr>
		<td height="40">&nbsp;</td>
		<td><span class="vslCss"><input type="submit" name="submit" id="submit" value="Change Password"></span></td>
	</tr>
</table>
</form>
</div>
<br>
<!--#include file="footer.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your startup script here
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

' Page Load event
Sub Page_Load()

'***Response.Write "Page Load"
End Sub

' Page Unload event
Sub Page_Unload()

'***Response.Write "Page Unload"
End Sub
%>
