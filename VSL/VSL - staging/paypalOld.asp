<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->

<%

dim arr, prod, p, Amt, TotalAmt, subject, qty, Totalqty, InstQ, InvId, CustId, Security, applyDiscount, conn, qMax
applyDiscount= checkCustomer(c)
Set Security = New cAdvancedSecurity
CustId = Security.CurrentUserID

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Request.ServerVariables("APPL_PHYSICAL_PATH") & "\db\vsldb.mdb" & ";"
Set rs = Server.CreateObject("ADODB.Recordset")
qMax = "SELECT MAX(InvoiceId) AS InvId FROM Orders;"
rs.open qMax,conn
if(not rs.eof) then
	if IsNull(rs("InvId")) Then
		InvId = 1
	else
		InvId= CInt(rs("InvId")) + 1
	End if
	Session("InvoiceId") = InvId
end if
rs.close

p = 0
qty = 0
ind = 0
TotalAmt = 0
Totalqty = 0
subject = "VSL#3"
s_fname = request.form("ship_FirstName") 
s_lname = request.form("ship_LastName") 
s_addr1 = request.form("ship_Address")
s_addr2 = request.form("ship_Address2")
s_city  = request.form("ship_city")
s_prov  = request.form("ship_Province")
s_tel   = request.form("HomePhone")	
s_postal = request.form("ship_PostalCode")
s_country = request.form("ship_Country")

arr=Session("cart")

if(IsCartEmpty()) then  response.redirect "vslCart.asp"
if(uBound(arr)=-1) then
	response.redirect "vslCart.asp?cmd=resetall"
else
	For i=0 To UBound(arr)
		prod = arr(i,0)
		qty = CInt(arr(i,3)) 
		Totalqty = Totalqty + qty
		Amt = FormatNumber(CInt(arr(i,2)) * CInt(arr(i,3)),2)
		TotalAmt = TotalAmt + Amt
		
		InstQ = "INSERT INTO Orders(CustomerId,ProductId,InvoiceId,Quantity,Amount,Ship_FirstName,Ship_LastName,Ship_Address,Ship_Address2,Ship_City,Ship_Province,Ship_Postal,Ship_Country,Ship_Phone,Ship_Email) " 
		InstQ = InstQ & "VALUES("& CustId &","& prod &","& InvId &","& qty &","& Amt &",'"& s_fname &"','"& s_lname &"','"& s_addr1 &"','"& s_addr2 &"','"& s_city &"','"& s_prov &"','"& s_postal &"','"& s_country &"','"& s_tel &"','N/A')" 
		conn.execute InstQ
		
		Dim query2 As String = "Select @@Identity"
		conn.execute query2
		
		ind = ind + 1
		
    Next

	
	'VB
	Dim query As String = "Insert Into Categories (CategoryName) Values (?)"
	Dim query2 As String = "Select @@Identity"
	Dim ID As Integer
	Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|Northwind.mdb"
	Using conn As New OleDbConnection(connect)
	Using cmd As New OleDbCommand(query, conn)
	cmd.Parameters.AddWithValue("", Category.Text)
	conn.Open()
	cmd.ExecuteNonQuery()
	cmd.CommandText = query2
	ID = cmd.ExecuteScalar()
	End Using
	End Using

	
	
	If ind = 2 Then
		subject = "VSL#3 Flavoured and Unflavoured 30 sachets"
	Else
		If prod = 2 Then
			subject = "VSL#3 Flavoured 30 sachets"
		Else
			subject = "VSL#3 Unflavoured 30 sachets"
		End If
	End If
	email = "info@vsl3.com"	
	amt = FormatNumber(TotalAmt/Totalqty,2)
	
	' response.write subject & " == subject<br />"
	' response.write TotalAmt & " == TotalAmt<br />"
	' response.write Totalqty & " == Totalqty<br />"
	' response.write s_fname & " == First name<br />"
	' response.write s_lname & " == Last name<br />"
	' response.write InvId & " == Invoice Id<br />"
	' response.end 

%>

<p style="text-align:center; vertical-align:middle">Processing...Please wait...<img src="images/loading.gif"></p>
<form name="frmPayPal" action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post" >
    <input type="hidden" name="cmd" value="_ext-enter">
	<input type="hidden" name="redirect_cmd" value="_xclick">
	<input type="hidden" name="business" value="rtran_1302632991_biz@ravenshoegroup.com" />
	<input type="hidden" name="invoice" value="<%=InvId%>">  
	<input type="hidden" name="item_name" value="<%=subject%>" />
	<input type="hidden" name="item_number" value="<%=InvId%>">	
	<input type="hidden" name="custom" value="<%=subject%>">
	<input type="hidden" name="amount" value="<%=amt%>">
	<input type="hidden" name="mc_gross" value="<%=amt%>">
	<input type="hidden" name="quantity" value="<%=Totalqty%>" />
	<input type="hidden" name="address_override" value="1">
	<input type="hidden" name="first_name" value="<%=s_fname %>">
	<input type="hidden" name="last_name" value="<%=s_lname %>">
	<input type="hidden" name="email" value="<%=email%>">
	<input type="hidden" name="charset" value="iso-8859-2">
	<input type="hidden" name="currency_code" value="CAD">
	<input type="hidden" name="cbt" value="Complete your Order Confirmation">
	<input type="hidden" name="no_note" value="1">
	<input type="hidden" name="return" value="http://www.ravenshoegroup.ca/VSLPayPal/Thank-you.asp?token=<%=InvId%>">
	<!-- <input type="hidden" name="notify_url" value="http://www.ravenshoegroup.ca/Paypal_Confirmation.asp?token=<%=InvId%>" > -->
	<input type="hidden" name="cancel_return" value="http://www.ravenshoegroup.ca/VSLPayPal/Cancel_order.asp?token=<%=InvId%>">
	<input type="hidden" name="bn" value="osCommerce PayPal IPN v2.3.3">
	<input type="hidden" name="lc" value="CA">
</form>

<!-- 
<form name="frmPayPal" method="post" action="https://api-3t.sandbox.paypal.com/nvp">
<input type="hidden" name="USER" value="API_username">	 
<input type="hidden" name="PWD" value="API_password">	 
<input type="hidden" name="SIGNATURE" value="API_signature"> 	 
<input type="hidden" name="version" value="xx.0"> 	 
<input type="hidden" name="PAYMENTREQUEST_0_PAYMENTACTION" value="Sale">
<input type="hidden" name="PAYMENTREQUEST_0_AMT" value="19.95">
<input type="hidden" name="RETURNURL" value="http://www.vsl3.ca/Paypal_Confirmation.asp" >
<input type="hidden" name="CANCELURL" value="http://www.ravenshoegroup.com/Paypal_Confirmation.html">
</form>
-->
<script type="text/javascript" language="JavaScript">
<!--
document.frmPayPal.submit();
//-->
</script>
	
<%

end if
conn.Close ' Close Connection
function getPrice(ItemId,c,q)
	dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open  "SELECT PRICE,Price_Rebate FROM Products WHERE (Products.[ItemId]=" & ItemId & ") ;", conn, 1, 2 
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

Sub DisplayItems()
	dim arr
	arr=GetCart()
	if( IsCartEmpty()) then  
		response.redirect "vslCart.asp?cmd=resetall"
	else
		dim applyDiscount
		applyDiscount= checkCustomer(conn)
		response.Write "<div class='t' width=720px><b>My cart:</b>"

%>
	<table width="760px">
    <tr>
    <td><table width="720px"  class="ewTable">
        <tr >
          <td width="50%" class="ewTableHeader"><b>Products</b></td>
          <td width="16%" class="ewTableHeader"><div align="right"><b>Unit Price</b></div></td>
          <td width="12%" class="ewTableHeader"><div align="center"><b>Quantity</b></div></td>
          <td width="10%" class="ewTableHeader"><div align="right"><b>Total</b></div></td>
        </tr>
		
        <%For i=0 To UBound(arr)
			p=arr(i,2)
			t=arr(i,2)
				if(applyDiscount) then
					
				'	p=62.50
					
				'	t= "<p>Regular pricing : <s>" & arr(i,2) & "<br>"&vbCrLf 
				'	t=t & "</s><font color=""#FF0000"" size=""-1"">Your special price for this order only</font>"
				'	t=t & ": <font color=""#FF0000""><strong>" & "62.50"  &"</strong></font> </p>"&vbCrLf
					
				end if
		
		%>
        <tr >
          <td width="50%"><b><%= "</b>" & arr(i,1) %></b></td>
          <td width="16%"><div align="right"><b><%=t%></b></div></td>
          <td width="12%" align="center"><div align="center">
              <%
			  dim tc
			  tc=arr(i,3) 
			  if(tc>19) then 
				Response.write tc + 2
			else
				Response.write tc
			end if
			  %>
            </div></td>
          <td width="10%"><div align="right"><%=p *  arr(i,3)%></div></td>
        </tr>
        <%Next%>
		
        <tr bordercolor="#FFCC66">
          <td width="50%" class="ewTablePager">&nbsp;</td>
          <td width="16%" class="ewTablePager">&nbsp;</td>
          <td width="12%" class="ewTablePager"><div align="right">Total:</div></td>
          <td width="10%" class="ewTablePager"><div align="right"><%=FormatNumber(TotalPurchase(),2)%></div></td>
        </tr>
      </table></td>
    </tr>
    </table>
	<%if(applyDiscount) then response.write " This order will be Shipped free of charge"%>
	<%
	response.write "</div>"
	end if
End Sub

function checkCustomer(c)
	checkCustomer=false
	'disabled Nov 20 2009
	' dim rs,NewCustomer
	' Set rs = Server.Creatumber")),10)
		' rs.close	
		' rs.Open  "SELECT phonecode FROM phone WHERE phonecode='" & Inv_LastName   &  Inv_FirstName & inv_PhoneNumber & "';" , c, 1, 2 
		' if(not rs.eof) then NewCustomer = false
	' else
		' NewCustomer=false
	' end if
	' checkCustomer=NewCustomer
	'checkCustomer=false
	' rs.close
	' set rs=nothing
end function

function cleantxt(t)
	dim temp
	if(isnull(t)) then t=""
	if(t<>"") then
		t=LCase(t)
		t= replace (t,"-","")
		t= replace (t," ","")
		t= replace (t,"(","")
		t= replace (t,")","")
		t= replace (t,"'","")
		t= replace (t,".","")
	end if
	cleantxt=t
end function
%>