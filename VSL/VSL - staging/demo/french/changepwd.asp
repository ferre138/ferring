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
		if  (!ew_OnError(fobj.opwd, "S'il vous plaît entrez le Mot de passe"))
			return false;
	}
	if  (!ew_HasValue(fobj.npwd)) {
		if  (!ew_OnError(fobj.npwd, "S'il vous plaît entrez le Mot de passe"))
			return false;
	}
	if  (fobj.npwd.value != fobj.cpwd.value) {
		if  (!ew_OnError(fobj.cpwd, "Mot de passe mal assortis"))
			return false;
	}
	return true;
}
// end JavaScript -->
</script>

<div align="left">
  <table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
    <tr>
      <td  width="680" rowspan="2"><img src="images/title_account_fr.png" width="410" height="75"></td>
      
    </tr>
    <tr>
      <td colspan="4" valign="top"><div align="right">
        <p class="bodycopy_small" style="color: #1070A3"><a href="../changepwd.asp">english  &gt;</a></p>
      </div></td>
      </tr>
  </table>
     </div> 
<% If Not bValidPwd Then %>
<p><span class="ewmsg">Invalid Password</span></p>
<% End If %>
<div class="t">
  <div align="right"><span class="vslcss"><a href="VSLOrderForm.asp">Retour aux produits </a> :  <a href="vslCart.asp">Visualiser le panier </a> :
            <%

	if (Not Security.IsLoggedIn()) then%>
	      
          <a href="login.asp">inscription</a>
            <%else%>
	        <a href="Customersedit.asp">Reviser votre compte</a> :
		     <a href="changepwd.asp">Changer le mot de passe</a> :
             <a href="logout.asp">Quitter</a>
             <%end if
	set Security =nothing %>
             <img src="images/spacer.gif" width="25" height="10">
        </span>
  </div>
  <form action="changepwd.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<p class="subheading">Changer le mot de passe </p>
 <table width="384" border="0" cellpadding="4" cellspacing="0">
	<tr>
		<td height="40"><span class="vslCss">Vieux mot de passe</span></td>
		<td><span class="vslCss"><input type="password" name="opwd" id="opwd" size="20"></span></td>
	</tr>
	<tr>
		<td height="40"><span class="vslCss">Nouveau mot de passe</span></td>
		<td><span class="vslCss">
		  <input type="password" name="npwd" id="npwd" size="20">
		</span></td>
	</tr>
	<tr>
		<td height="40"><span class="vslCss">Confirmer le mot de passe </span></td>
		<td><span class="vslCss"><input type="password" name="cpwd" id="cpwd" size="20"></span></td>
	</tr>
	<tr>
		<td height="40">&nbsp;</td>
		<td><span class="vslCss"><input type="submit" name="submit" id="submit" value="Changer le mot de passe"></span></td>
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
