<%
Const EW_PAGE_ID = "login"
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

' Common page loading event (in userfn60.asp)
Call Page_Loading()
%>
<%

' Page load event, used in current page
Call Page_Load()
%>
<%
Dim bValidate, bValidPwd, sUsername, sPassword
Dim sLoginType
Dim sLastUrl
sLastUrl = Security.LastUrl ' Get Last Url
If sLastUrl = "" Then sLastUrl = "checkout.asp"
bValidate = False
If Request.Form("submit") <> "" Then

	' Setup variables
	sUsername = Request.Form("Username")
	sPassword = Request.Form("Password")
	sLoginType = LCase(Request.Form("rememberme"))
	bValidate = True
Else
	If Security.IsLoggedIn() Then
		If Session(EW_SESSION_MESSAGE) = "" Then Page_Terminate(sLastUrl) ' Return to last accessed page
	End If
End If
If bValidate Then
	bValidPwd = False

	' Call loggin in event
	bValidate = User_LoggingIn(sUsername, sPassword)
	If bValidate Then
		bValidPwd = Security.ValidateUser(sUsername, sPassword)
		If Not bValidPwd Then Session(EW_SESSION_MESSAGE) = "Incorrect user ID or password" ' Invalid user id/password
	Else
		If Session(EW_SESSION_MESSAGE) = "" Then Session(EW_SESSION_MESSAGE) = "Login cancelled" ' Login cancelled
	End If
	If bValidPwd Then

		' Write cookies
		If sLoginType = "a" Then ' Auto login
			Response.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" ' Set up autologin cookies
			Response.Cookies(EW_PROJECT_NAME)("username") = sUsername ' Set up user name cookies
			Response.Cookies(EW_PROJECT_NAME)("password") = ew_Encode(TEAencrypt(sPassword, EW_RANDOM_KEY)) ' Set up password cookies
			Response.Cookies(EW_PROJECT_NAME).Expires = DateAdd("d", 365, Date) ' Change the expiry date of the cookies here
		ElseIf sLoginType = "u" Then ' Remember user name
			Response.Cookies(EW_PROJECT_NAME)("autologin") = "rememberUsername" ' Set up remember user name cookies
			Response.Cookies(EW_PROJECT_NAME)("username") = sUsername ' Set up user name cookies
			Response.Cookies(EW_PROJECT_NAME).Expires = DateAdd("d", 365, Date) ' Change the expiry date of the cookies here
		Else
			Response.Cookies(EW_PROJECT_NAME)("autologin") = "" ' Clear autologin cookies
		End If

		' Call loggedin event
		Call User_LoggedIn(sUsername)
		Call Page_Terminate(sLastUrl) ' Return to last accessed url
	End If
End If
%>
<!--#include file="header.asp"-->

<table width="805"  border="0" cellpadding="0" cellspacing="0" id="Table_01">
            <tr>
            <td width="680" rowspan="2"><img src="images/title_order_fr.png" width="410" height="75"></td>
             
            </tr>
            <tr>
              <td colspan="4" valign="top"><div align="right">
                <p class="bodycopy_small" style="color: #1070A3"><a href="../login.asp" class="bodycopy_small">english  &gt;</a></p>
              </div></td>
              </tr>
        </table>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
// To include another .js script, use:
// ew_ClientScriptInclude("my_javascript.js"); 
//-->
</script>
<script type="text/javascript">
<!--

function ew_ValidateForm(fobj) {
	if (!ew_HasValue(fobj.username)) {
		if  (!ew_OnError(fobj.username, "Please enter user ID"))
			return false;
	}
	if (!ew_HasValue(fobj.password)) {
		if (!ew_OnError(fobj.password, "Please enter password"))
			return false;
	}
	return true;
}
//-->
</script>


<%
If Session(EW_SESSION_MESSAGE) <> "" Then
%>
<p><span class="ewmsg"><%= Session(EW_SESSION_MESSAGE) %></span></p>
<%
	Session(EW_SESSION_MESSAGE) = "" ' Clear message
End If
%>
<form action="login.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<div class="t">
  <p>
  <div align="right" class="vslcss"><a href="VSLOrderForm.asp">Retour aux produits </a> : <a href="vslCart.asp">Visualiser le panier </a> </div>
  </p>
  <table width="775" border="0" align="center" cellpadding="4" cellspacing="0" class="ewTable">
	    <tr>
	      <td>&nbsp;</td>
	      <td>&nbsp;</td>
	      <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td width="50%" rowspan="3"><div align="center">Vous n’avez pas de compte ? <br>
  Il est facile de s’inscrire  </div></td>
	    <td rowspan="3">&nbsp;</td>
	    <td height="15"><span class="vslcss">Nom d’utilisateur</span></td>
		<td><span class="vslcss">
		  <input type="text" name="username" id="username" size="20" value="<%= Request.Cookies(EW_PROJECT_NAME)("username") %>"></span></td>
	</tr>
	<tr>
	  <td height="15"><span class="vslcss">Mot de passe</span></td>
		<td><span class="vslcss"><input type="password" name="password" id="password" size="20"></span></td>
	</tr>
	<tr>
	  <td height="15">&nbsp;</td>
		<td><span class="vslcss"><a href="forgetpwd.asp">Mot de passe oublié ?</a>
		</span></td>
	</tr>
	<tr>
		<td align="center"><p><span class="vslcss">
<a href="register.asp">Inscrivez-vous ici  </a>
</span></p></td>
	    <td align="center">&nbsp;</td>
	    <td colspan="2" align="center"><span class="vslcss">
	      <input type="submit" name="submit" id="submit" value="    Ouverture de session   ">
	    </span></td>
      </tr>
</table>
</div>
</form>

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
<%

' User Logging In event
Function User_LoggingIn(usr, pwd)
	On Error Resume Next

	' Enter your code here
	' To cancel, set return value to False

	User_LoggingIn = True
End Function

' User Logged In event
Sub User_LoggedIn(usr)

	' Response.Write "User Logged In"
End Sub
%>
