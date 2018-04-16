<link href="../vsl.css" rel="stylesheet" type="text/css" media="screen">
<!--#include file="../ewcfg60.asp"-->
<!--#include file="../aspfn60.asp"-->
<table cellspacing="0" cellpadding="3" class="invoice" style="border-style:dashed;border-width:1px; " border='0'>
<%
Dim Security
	Set Security = New cAdvancedSecurity
	If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
	If Not Security.IsLoggedIn() Then
		Call Security.SaveLastUrl()
		response.redirect ("login.asp")
	End If
Session.Timeout = 120
Dim strSQL, rst, orderid, totalAmount, counter
counter = 0

orderid = Request.Querystring("token")


if(IsNumeric(orderid)) then

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING

strSQL = "SELECT Orders.Ship_FirstName, Orders.Ship_LastName, Orders.Ship_Address, Orders.Ship_Address2, Orders.Ship_City, Orders.Ship_Province, Orders.Ship_Postal, Orders.Ship_Country, Orders.Ship_Phone, productId,Quantity, Amount, Tax, Shipping,price FROM orders,orderdetails "
strSQL = strSQL & "WHERE orders.orderId = orderdetails.orderId and orders.orderId =" & orderid & " And customerId = " & Security.CurrentUserID & ";"

Set rst = conn.Execute(strSQL)
Set rs = Server.CreateObject("ADODB.Recordset")
if(not rst.EOF) then
	totalAmount = CDbl(rst("Amount"))
	Tx = CDbl(rst("Tax"))
	ship = CDbl(rst("Shipping"))

	Do While not rst.EOF
	
		qProd = "SELECT ItemNo, description, Price FROM Products WHERE ItemId =" & rst("ProductId") & ";"
		rs.open qProd,conn					
		if counter = 0 Then
			Response.write "<tr height='36'>"
			Response.write "<td align='left' width='80px' style='padding-left:5px'>Item #</td>"						
			Response.write "<td align='center' width='*'>Description</td>"						
			Response.write "<td align='center' width='70px'>Quantity</td>"
			Response.write "<td align='right' width='70px'>Price</td>"						
			Response.write "<td align='right' width='80px' style='padding-right:3px'>Amount</td>"						
			Response.write "</tr>"
			pShipName = rst("Ship_FirstName") & " " & rst("Ship_LastName")
			pShipAddr = rst("Ship_Address") & "<br>" & rst("Ship_Address2")
			pShipCity = rst("Ship_City")
			pShipProv = rst("Ship_Province")
			pShipPostal = rst("Ship_Postal")
			pShipCountry = rst("Ship_Country")
		End if 
		q= cdbl(rst("Quantity"))
		p= cdbl(rst("price"))
		
		Response.write "<tr height='25'>"
		Response.write "<td style='padding-left:5px'>" & rs("ItemNo") & "</td>"
		Response.write "<td style='padding-left:5px'>" & rs("description") & "</td>"
		Response.write "<td align='center'>" & q & "</td>"						
		Response.write "<td align='right' style='padding-right:3px'>$" & p  & "</td>"
		Response.write "<td align='right' style='padding-right:3px' class='invoice'>$" & FormatNumber(p * q,2) & "</td>"
		Response.write "</tr>"
		counter = 1
		rs.close
		rst.MoveNext						
	Loop
	Response.write "<tr height='30'>"
	Response.write "<td colspan='4' align='right' style='padding-right:3px'>Shipping and handling:</td><td align='right'>$" & FormatNumber(ship,2) & "</td>"
	Response.write "</tr>"

	Response.write "<tr>"
	Response.write "<td colspan='4' align='right' style='padding-right:3px'>Tax:</td><td align='right'>$" & FormatNumber(tx,2) & "</td>"
	Response.write "</tr>"
				
	Response.write "<tr height='36'>"
	Response.write "<td colspan='4' align='right' style='padding-right:3px'><b>Total Amount: </td><td align='right'>$" & FormatNumber(totalAmount,2) & "</b></td>"
	Response.write "</tr>"
else	
	Response.write "<tr><td>lexical invalide</td></tr>"
end if
rst.close  ' Close Recordset
conn.Close ' Close Connection
else
Response.write "<tr><td>lexical invalide</td></tr>"

end if
%>			  
</table>
<div style="padding-left:3px;padding-top:20px;">
<table border="0" cellspacing="0" cellpadding="1" class="invoiceTotal" >
<tr><td><i><span style="Verdana;font-size: 14px;" ><strong>Expédier à</strong></span></i></td></tr>
<tr><td><i>
<div style="font-family: Verdana;font-size: 12px;" class="invoiceTotal">
	<%=pShipName %><br />
	<%=pShipAddr %><br />
	<%=pShipCity %>, <%=pShipProv %><br />
	<%=pShipPostal %>, <%=pShipCountry %><br /></div></i>
	</td></tr>
	</table>
	<br /><br />
<i>
<div style="font-family: Verdana;font-size: 14px;" ><strong>Merci pour votre paiement</strong> </div>
	<div style="font-family: Verdana;font-size: 12px;" class="invoiceTotal">
		Votre transaction est complétée et vous recevrez un reçu pour votre commande de Paypal. </div>
</i>
<br />
<br />
<div style="padding-left:0px;font-family: Arial, Helvetica, sans-serif;font-size: 14px;">
<a href="javascript:window.print();">Imprimer</a>

</div>