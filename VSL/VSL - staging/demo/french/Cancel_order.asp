
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%

orderid = Request.Querystring("token")
custid = Request.Querystring("txt")
if((orderid & "x"="x") or (custid & "x"="x"))then response.redirect "VSLOrderForm.asp"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
		
strSQL = "UPDATE Orders SET payment_status='Cancelled', payment_date=Now() " 
strSQL = strSQL & " WHERE payment_status='WIP' and orderid =" & orderid & " and Orders.CustomerId=" & custid & " ;"  
 
 conn.Execute(strSQL)
'if orderid = Session("orderid") Then conn.Execute(strSQL)

conn.Close ' Close Connection

%>

<!--#include file="header.asp"--><script language="JavaScript" type="text/javascript">
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
<table width="820" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="700" height="50" align="left" valign="middle">&nbsp;</td>
    <td width="120" align="left" valign="top"><div align="right" class="bodycopy_small">
      <p>&nbsp;</p>
      <p><a href="../Cancel_order.asp?<%=request.QueryString%>">english &gt;</a></p>
    </div></td>
  </tr>
  <tr>
    <td height="300" colspan="2" align="left" valign="top"><table width="820" border="0" cellspacing="0" cellpadding="0">
      <tr align="left" valign="top">
        <td width="620" class="bodycopy"><%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
          <a href="login.asp">inscription</a>
          <%else%>
          <a href="Customersedit.asp">Changement au compte </a> : <a href="changepwd.asp">Changement au mot de passe</a> : <a href="logout.asp">Quitter</a>
          <%end if
	 %>
          <table width="610" border="1" cellspacing="0" cellpadding="9" class="invoice">
          
          </table>
          <i>
            <div style="font-family: Verdana;font-size: 14px; padding-top:10px;" ><strong>Transaction annulée</strong></div>
            <div style="font-family: Verdana;font-size: 12px;" class="invoiceTotal"></div>
            </i>
          <div style="padding-left: 3px;padding-top:24px;"></div></td>
        <td width="200" ><img src="images/packshot_v2.png" width="200" height="200"></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2"></td>
  </tr>
</table><!--#include file="footer.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your startup script here
// document.write("page loaded");
//-->
</script>

