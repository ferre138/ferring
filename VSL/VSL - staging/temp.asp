<!--#include file="ewcfg60.asp"-->
<% 
 Set conn = Server.CreateObject("ADODB.Connection")
 conn.Open EW_DB_CONNECTION_STRING

 response.write getEmailText(14148,conn)

function getEmailText(InvId,conn)
	
	sBody  = "New Payment updated at VSL3.ca on " & Now &   vbcrlf
	sBody  =sBody & "<table cellspacing='0' cellpadding='3' class='invoice' style='border-style:dashed;border-width:1px;' border='0'>"
	sBody  =sBody & "<tr height='36'>"
	sBody  =sBody & "<th align='left' width='80px' style='padding-left:5px'>Item #</th>"						
	sBody  =sBody & "<th align='center' width='*'>Description</th>"						
	sBody  =sBody & "<th align='center' width='70px'>Quantity</th>"
	sBody  =sBody & "<th align='right' width='70px'>Price</th>"						
	sBody  =sBody & "<th align='right' width='80px' style='padding-right:3px'>Amount</th>"						
	sBody  =sBody & "</tr>"
	strSQL = "SELECT products.ItemNo AS ItemNo,products.Description AS Description,orderdetails.Price AS Price,orderdetails.Quantity AS Quantity "
	strSQL = strSQL & " FROM orders,orderdetails,products WHERE orders.orderId = orderdetails.orderId and products.itemid = orderdetails.productId and orders.orderid = " & InvId & ";"
	Set rst = conn.Execute(strSQL)
	Do While not rst.EOF
		sBody  =sBody & "<tr>"
		sBody  =sBody & "<td align='left'>" & rst("ItemNo") & "</td>"						
		sBody  =sBody & "<td >" & rst("Description") & "</td>"						
		sBody  =sBody & "<td align='center'>" & rst("Quantity") & "</td>"
		sBody  =sBody & "<td align='right'>$" & rst("Price") & "</td>"						
		sBody  =sBody & "<td align='right'>$" & rst("Price") * rst("Quantity") & "</td>"						
		sBody  =sBody & "</tr>"		
		Subtotal = Subtotal + CDbl(rst("Price") * rst("Quantity")) 
		rst.MoveNext						
	Loop
	rst.close
	
	sBody  =sBody & "<tr>"
	sBody  =sBody & "<td align='right' colspan='4'><b>Sub-Total:</b></td>"						
	sBody  =sBody & "<td align='right'><b>$" & Subtotal & "</b></td>"						
	sBody  =sBody & "</tr>"	
	sBody  =sBody & "</table>"
	
	sBody  =sBody & "<br /><div style='font-family: Verdana;font-size: 14px; padding-top:10px;' ><a href='http://www.vsl3.ca/beta/admin/'><i>Please login to see the payment.</i></a></div>"
	
	getEmailText = sBody
End Function
%>