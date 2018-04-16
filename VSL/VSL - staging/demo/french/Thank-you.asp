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
      <p><a href="../Thank-you.asp">english &gt;</a></p>
    </div></td>
  </tr>
  <tr>
    <td height="300" colspan="2" align="left" valign="top"><table width="820" border="0" cellspacing="0" cellpadding="0">
      <tr align="left" valign="top">
        <td width="620" class="bodycopy"><%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
          <a href="login.asp">login</a>
          <%else%>
          <a href="Customersedit.asp">Edit account</a> : <a href="changepwd.asp">Change Password</a> : <a href="logout.asp">logout</a>
          <%end if
	 %>
          <table width="610" border="1" cellspacing="0" cellpadding="9" class="invoice">
            <%

					If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
					If Not Security.IsLoggedIn() Then
						Call Security.SaveLastUrl()
						response.redirect ("login.asp")
					End If

					Session.Timeout = 120
					Dim strSQL, rst, orderid, totalAmount, counter
					counter = 0
					orderid = Request.Querystring("token")
					Session("pOrderid") = orderid
					
					if(IsNumeric(orderid)) then
					
					Set conn = Server.CreateObject("ADODB.Connection")
					conn.Open EW_DB_CONNECTION_STRING
					
					strSQL = "SELECT Orders.Ship_FirstName, Orders.Ship_LastName, Orders.Ship_Address, Orders.Ship_Address2, Orders.Ship_City, Orders.Ship_Province, Orders.Ship_Postal, Orders.Ship_Country, Orders.Ship_Phone, productId,Quantity, Amount, Tax, Shipping,price FROM orders,orderdetails "
					strSQL = strSQL & "WHERE orders.orderId = orderdetails.orderId and orders.orderId =" & orderid & " And customerId = " & Security.CurrentUserID & ";"

					
					Set rst = conn.Execute(strSQL)
					totalAmount = CDbl(rst("Amount"))
					Tx = CDbl(rst("Tax"))
					ship = CDbl(rst("Shipping"))
					Set rs = Server.CreateObject("ADODB.Recordset")
					Do While not rst.EOF
						qProd = "SELECT ItemNo, description FROM Products WHERE ItemId =" & rst("ProductId") & ";"
						rs.open qProd,conn
						if counter = 0 Then
							Response.write "<tr height='30'>"
							Response.write "<td align='center'>Item #</td>"						
							Response.write "<td align='center'>Description</td>"						
							Response.write "<td align='center'>Quantity</td>"
							Response.write "<td align='center'>Unit Price</td>"						
							Response.write "<td align='center'>Total</td>"						
							Response.write "</tr>"
						End if 
						q= cdbl(rst("Quantity"))
						p= cdbl(rst("price"))
						
						Response.write "<tr height='25'>"
						Response.write "<td style='padding-left:5px'>" & rs("ItemNo") & "</td>"
						Response.write "<td style='padding-left:5px'>" & rs("description") & "</td>"
						Response.write "<td align='center'>" & q & "</td>"
						Response.write "<td align='right' style='padding-right:3px'>$" & p & "</td>"
						Response.write "<td align='right' style='padding-right:3px'>$" & (q * p)  & "</td>"
						Response.write "</tr>"
						Subtotal = Subtotal + CDbl(p * q) 
						counter = 1
						rs.close
						rst.MoveNext						
					Loop
					
					Response.write "<tr height='30'>"
					Response.write "<td colspan=5 align='right' style='padding-right:3px'><b>Sub-total:&nbsp;&nbsp;$" & FormatNumber(Subtotal,2) & "</b></td>"
					Response.write "</tr>"
	
					Response.write "<tr height='30'>"
					Response.write "<td colspan='4' align='right' style='padding-right:3px'>Shipping and handling:</td><td align='right'>$" & FormatNumber(ship,2) & "</td>"
					Response.write "</tr>"

					Response.write "<tr>"
					Response.write "<td colspan='4' align='right' style='padding-right:3px'>Tax:</td><td align='right'>$" & FormatNumber(tx,2) & "</td>"
					Response.write "</tr>"
				
					Response.write "<tr height='36'>"
					Response.write "<td colspan='4' align='right' style='padding-right:3px'><b>Total Amount: </td><td align='right'>$" & FormatNumber(totalAmount,2) & "</b></td>"
					Response.write "</tr>"
	
					rst.close
					conn.Close ' Close Connection
					
					Else
						Response.write "<tr><td>Invalid token</td></tr>"
					
					End If
					
					

					%>
          </table>
          <i>
            <div style="font-family: Verdana;font-size: 14px; padding-top:10px;" ><strong>Merci pour votre paiement</strong></div>
            <div style="font-family: Verdana;font-size: 12px;" class="invoiceTotal"> Votre transaction est complétée et vous recevrez un reçu pour votre commande de Paypal <br />
               </div>
            </i>
          <div style="padding-left: 3px;padding-top:24px;"> <a href="invoice.asp?token=<%=orderid %>" target="_blank"><span class="submitbutton"><u>Imprimer la facture</u></span></a> </div></td>
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
