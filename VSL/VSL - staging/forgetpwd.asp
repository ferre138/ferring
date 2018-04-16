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

const EW_SMTPSERVER="smtp.1and1.com"
const EW_SMTPSERVER_USERNAME="info@vsl3.ca"
const EW_SMTPSERVER_PASSWORD="web4colon"
const EW_SMTPSERVER_PORT=25





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
        Dim sMail
		sFrEmail = "info@ferring.ca"
		sSubject = "Password Recovery From VSL3.ca " 
        

        sbody = "Dear Sir/Madam, \n\n Please see below for the requested information: \n\n User Name: " & sUserName & "\n Password: " & sPassword & "\n \n Please feel free to contact us in case of further queries. \n \n Best Regards, \n Support"
        sbody = replace(sbody, "\n", vbCrLf)
        sbody = replace(sbody, vbcrlf,"<br>" )
        sbody = replace(sbody, "/r/n","<br>" )
        sTextBody = sBody 
        sHTMLBody = sBody
        sbody = replace(sbody,"                   " & vbcrlf ,"")
        sMail = sBody

        Call Email_1and1(sFrEmail, sEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)

        Session(EW_SESSION_MESSAGE) = "Password sent to your email. Please check your email inbox." ' Set success message
		Call Page_Terminate("login.asp") ' Return to login page
	End If
Else
	bValidEmail = True
End If





Sub Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat,sRtn,aAtt)


Dim objMail
Set objMail = Server.CreateObject("CDO.Message")
Set objConfig = Server.CreateObject("CDO.Configuration")


objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.1and1.com"
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vsl3.ca" 'Enter YOUR E-MAIL ADDRESS here
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "web4colon" 'Enter the PASSWORD for your email address
objConfig.Fields.Update

Set objMail.Configuration = objConfig

objMail.From = sFrEmail 'Enter the FROM ADDRESS
objMail.To = sToEmail 'Enter the TO ADDRESS
objMail.Subject = sSubject 'Enter a SUBJECT

objMail.TextBody = sMail 'Enter the BODY of the message
objMail.HTMLBody = sMail

objMail.Send


If Err.Number = 0 Then
Response.Write("Mail sent!<br>")
Response.Write(sSubject & "<br>")
Response.Write("<hr>")
Else
Response.Write("Error sending mail. Code: " & Err.Number)
Err.Clear
End If
Set objMail=Nothing
Set objConfig=Nothing

End Sub






Sub Email_1and1x(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat,sRtn,aAtt)

Dim objMail
Set objMail = Server.CreateObject("CDO.Message")
Set objConfig = Server.CreateObject("CDO.Configuration")


objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.1and1.com"
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vsl3.ca" 'Enter YOUR E-MAIL ADDRESS here
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "web4colon" 'Enter the PASSWORD for your email address
objConfig.Fields.Update

Set objMail.Configuration = objConfig
objMail.From = sFrEmail 'Enter the FROM ADDRESS
objMail.To = sToEmail 'Enter the TO ADDRESS
objMail.Subject =sSubject 'Enter a SUBJECT
objMail.TextBody=sMail 'Enter the BODY of the message
objMail.HTMLBody = sMail

objMail.Send

If Err.Number = 0 Then
Response.Write("Mail sent!<br>")
Response.Write(sSubject & "<br>")
Response.Write("<hr>")
Else
Response.Write("Error sending mail. Code: " & Err.Number)
Err.Clear
End If
Set objMail=Nothing
Set objConfig=Nothing

End Sub
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
		if  (!ew_OnError(fobj.email, "Please enter valid Email Address!"))
			return false;
	}
	if  (!ew_CheckEmail(fobj.email.value)) {
		if  (!ew_OnError(fobj.email, "Please enter valid Email Address!"))
			return false;
	}
	return true;
}
// end JavaScript -->
</script>
<p>
<table width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
  <tr>
    <td width="680" rowspan="2"><img src="images/title_account.png" width="410" height="75"></td>
  <!--  <td width="28" valign="top"><img src="images/fontsize.png" border="0" alt=""></td>
    <td width="24" valign="top"> <a href="#"
				onMouseOver="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onMouseOut="changeImages('login_13', 'images/font1.png'); return true;"
				onMouseDown="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onMouseUp="changeImages('login_13', 'images/login_13-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('default'); 
return false;"> <img name="login_13" src="images/font1.png" width="24" height="27" border="0" alt=""></a></td>
    <td width="24"  valign="top"> <a href="#"
				onMouseOver="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onMouseOut="changeImages('login_14', 'images/font2.png'); return true;"
				onMouseDown="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onMouseUp="changeImages('login_14', 'images/login_14-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Medium'); 
return false;"> <img name="login_14" src="images/font2.png" width="24" height="27" border="0" alt=""></a></td>
    <td width="26"  valign="top"> <a href="#"
				onMouseOver="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onMouseOut="changeImages('login_15', 'images/font3.png'); return true;"
				onMouseDown="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onMouseUp="changeImages('login_15', 'images/login_15-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Large'); 
return false;"><img name="login_15" src="images/font3.png" width="24" height="27" border="0" alt=""></a></td>
  </tr>
  <tr>
    <td colspan="4" valign="top"><div align="right">
      <p><a href="french/forgetpwd.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
    </div></td>-->
    </tr>
</table>
<br>
<a href="login.asp">Back to Login Page</a>
</span></p>
<% If Not bValidEmail Then %>
<p><span class="ewmsg">Invalid Email</span></p>
<% End If %>
<form action="forgetpwd.asp" method="post" onSubmit="return ew_ValidateForm(this);">
<div class="t">
<table height="85" border="0" cellpadding="4" cellspacing="0" class="ewTableNoBorder">
	<tr>
	  <td colspan="3"><span class="vslcss"> If you have lost or forgotten your  password, enter your e-mail address below and click "Send Password." </span></td>
		</tr>
	<tr>
	  <td width="42"><span class="vslcss">
	  </span></td>
		<td width="333"><span class="vslcss">
		  <input type="text" name="email" id="email" value="<%= sEmail %>" size="30" maxlength="50">
		</span></td>
		<td width="452"><span class="vslcss"><input type="submit" name="submit" id="submit" value="Send Password"></span></td>
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
