<%
Const EW_PAGE_ID = "forgetpwd"
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
Dim bValidEmail, sEmail, sUserName, sPassword, sFilter, sSql, rs
If Request.Form("email") <> "" Then
	bValidEmail = False

	' Setup variables
	sEmail = Request.Form("email")

	' Set up filter (Sql Where Clause) and get Return Sql
	' Sql constructor in Customers class, Customersinfo.asp

	sFilter = "[inv_EmailAddress] = '" & ew_AdjustSql(sEmail) & "'"
	Customers.CurrentFilter = sFilter
	sSql = Customers.SQL
	Set rs = conn.Execute(sSql)
	If Not rs.Eof Then
		sUserName = rs("UserName")
		sPassword = rs("passwrd")
		bValidEmail = True
	End If
	rs.Close
	Set rs = Nothing
	If bValidEmail Then
		Dim Email
		Set Email = New cEmail
		Email.Load("txt/forgetpwd.txt")
		Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
		Email.ReplaceRecipient(sEmail) ' Replace Recipient
		Email.ReplaceContent "<!--$UserName-->", sUserName
		Email.ReplaceContent "<!--$Password-->", sPassword
		Email.Send()
		Set Email = Nothing
		Session(EW_SESSION_MESSAGE) = "Votre mot de passe à été envoyé à votre adresse de courriel " ' Set success message
		Call Page_Terminate("login.asp") ' Return to login page
	End If
Else
	bValidEmail = True
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
	if  (!ew_HasValue(fobj.email)) {
		if  (!ew_OnError(fobj.email, "S'il vous plaît entrez un adresse de courriel valide"))
			return false;
	}
	if  (!ew_CheckEmail(fobj.email.value)) {
		if  (!ew_OnError(fobj.email, "S'il vous plaît entrez un adresse de courriel valide"))
			return false;
	}
	return true;
}
// end JavaScript -->
</script>
<p>
<table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
  <tr>
    <td width="680" rowspan="2"><img src="images/title_account_fr.png" width="410" height="75"></td>
    
  </tr>
  <tr>
    <td colspan="4" valign="top"><div align="right">
      <p class="bodycopy_small" style="color: #1070A3"><a href="../forgetpwd.asp" class="bodycopy_small">english  &gt;</a></p>
    </div></td>
    </tr>
</table>
<br>
<a href="login.asp">Remettre à Page d' Établissement</a>
</span></p>
<% If Not bValidEmail Then %>
<p><span class="ewmsg">Invalid Email</span></p>
<% End If %>
<form action="forgetpwd.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<div class="t">
<table height="85" border="0" cellpadding="4" cellspacing="0" class="ewTableNoBorder">
	<tr>
	  <td colspan="3"><span class="vslcss"> si vous avez perdu ou oublié votre mot de passe, entrez  votre adresse de courriel est et cliquez « Envoyer mot de passe »  </span></td>
		</tr>
	<tr>
	  <td width="42"><span class="vslcss">
	  </span></td>
		<td width="333"><span class="vslcss">
		  <input type="text" name="email" id="email" value="<%= sEmail %>" size="30" maxlength="50">
		</span></td>
		<td width="452"><span class="vslcss"><input type="submit" name="submit" id="submit" value="Envoyer mot de passe"></span></td>
	</tr>
</table>
</div>
</form>
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
