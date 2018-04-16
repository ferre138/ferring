<!-- Google Code for VSL#3 Sale Conversion Page -->
<script type="text/javascript">
/* <![CDATA[ */
var google_conversion_id = 980416341;
var google_conversion_language = "en";
var google_conversion_format = "3";
var google_conversion_color = "ffffff";
var google_conversion_label = "huyHCPvPvxAQ1e6_0wM";
var google_conversion_value = 0;
var google_remarketing_only = false;
/* ]]> */
</script>
<script type="text/javascript" src="//www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="//www.googleadservices.com/pagead/conversion/980416341/?value=0&amp;label=huyHCPvPvxAQ1e6_0wM&amp;guid=ON&amp;script=0"/>
</div>
</noscript>

<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->
<!--#include file="header.asp"--> 
<%
intLocale = SetLocale(4105)
Session.Timeout = 120			
dim c, OrderNum
Set c = Server.CreateObject("ADODB.Connection")
c.Open EW_DB_CONNECTION_STRING

Dim Security
Set Security = New cAdvancedSecurity
%> 
<%
If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
If Not Security.IsLoggedIn() Then
	Call Security.SaveLastUrl()
	Call Page_Terminate("login.asp")
End If

dim arr
arr=Session("cart")
if(  IsCartEmpty()) then  response.redirect "vslCart.asp"

if(uBound(arr)=-1) then
	'message "emptycart"
	response.redirect "vslCart.asp?cmd=resetall"
else

	'response.Write "<b>Your Cart:</b>"
%>

<!-- <FORM METHOD="POST" name="MyForm" ACTION="http://ww11.aitsafe.com/cf/pay.cfm">
  <input type="text" name="userid" value="D6155807">
  <input type="text" name="return" value="http://www.vsl.ca/test/cart.asp"><br> 
  Old code
  -->	
<%

CustId = Security.CurrentUserID
inv_fname =request.form("Inv_FirstName") 
inv_lname = request.form("Inv_LastName") 
inv_company =request.form("inv_company")
inv_addr1 =request.form("Inv_Address") 
inv_addr2 = request.form("Inv_Address2")
inv_city =request.form("Inv_city")
inv_state =request.form("inv_Province")
inv_zip =request.form("inv_PostalCode")
inv_country =request.form("inv_Country")
tel =request.form("inv_PhoneNumber")
fax =request.form("inv_Fax")
email =request.form("inv_EmailAddress")

del_fname = request.form("ship_FirstName") 
del_lname = request.form("ship_LastName") 
'del_company =request.form("del_company")

del_addr1 =request.form("ship_Address") 
del_addr2 =request.form("ship_Address2") 

del_city =request.form("ship_city")
del_state =request.form("ship_Province")
del_zip =request.form("ship_PostalCode")
del_country =request.form("ship_Country")
del_tel =request.form("HomePhone")

'if(request.form("CheckLocalPickup")="on") then
'	del_fname = request.form("Inv_FirstName") 
'	del_lname = request.form("Inv_LastName")
'	del_addr1 ="200 YorkLand" 
'	del_addr2 ="Suite:500"

'	del_city ="Toronto"
'	del_state ="ON"
'	del_zip ="M2J5C1"
'	del_country ="Canada"
'	del_tel ="416 642-0075"
'end if
gt=0

'**********************************
'	https://www.paypal.com/cgi-bin/webscr
'https://www.sandbox.paypal.com/cgi-bin/webscr
%>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
var pageTracker = _gat._getTracker("UA-1731303-15");
pageTracker._trackPageview();
</script>
  <form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="MyForm">
	<input type="hidden" name="cmd" value="_cart">
	<input type="hidden" name="upload" value="1">
	<input type="hidden" name="lc" value="FR">
	<input type="hidden" name="business" value="vsl3.orders@ferring.com" />
	<!--<input type="hidden" name="business" value="skarke_1247080008_biz@ravenshoegroup.com" />
	<input type="hidden" name="address_override" value="1">-->
	<input type="hidden" name="first_name" value="<%= inv_fname%>">
	<input type="hidden" name="last_name" value="<%= inv_lname%>">
<%
'**********************************
Set rs = Server.CreateObject("ADODB.Recordset")
' rs.CursorLocation = EW_CURSORLOCATION
rs.Open "Select * from orders where 1=2;", c, 1, 2
rs.AddNew
rs.fields ("CustomerId")=CustId
rs.fields ("Ship_FirstName")= del_fname
rs.fields ("Ship_LastName")= del_lname
rs.fields ("Ship_Address")= del_addr1
rs.fields ("Ship_Address2")= del_addr2
rs.fields ("Ship_City")= del_city
rs.fields ("Ship_Province")=del_state
rs.fields ("Ship_Postal")= del_zip
rs.fields ("Ship_Country")= del_country
rs.fields ("Ship_Phone")= del_tel
rs.fields ("Ship_Email")= email
OrderNum=rs.fields("OrderId")  'this is the new order Id
rs.update
rs.close
set rs=nothing

'**********************************
'RAMY insert ORDER table & get new ORDERID
' INSERT INTO Orders ( CustomerId, Ship_FirstName, Ship_LastName, Ship_Address, Ship_Address2, Ship_City, Ship_Province, Ship_Postal, Ship_Country, Ship_Phone, Ship_Email, Ordered_Date )
'******************************

dim applyDiscount
applyDiscount= checkCustomer(c,session("promocode"))

dim totalCnt

dim rs,strSql,taxrate,ship1,ship2

Set rst = Server.CreateObject("ADODB.Recordset")
strSql= " SELECT TaxRate, ShipRate_first, ShipRate_Rest  FROM Province WHERE (Province.Prov='"  & del_state &"');"
rst.open strSql,c
if(not rst.eof) then 
	taxrate=cdbl(rst("TaxRate"))
	m_taxrate=taxrate
	ship1=cdbl(rst("ShipRate_first"))
	ship2=cdbl(rst("ShipRate_Rest"))
end if
rst.close

'if((lcase(replace(del_zip," ", "")) ="m2j5c1") ) then
'	if((lcase(mid( replace(del_addr1," ",""),1,11))="200yorkland") and lcase(del_state)="on") then
'		ship1=0
'		ship2=0
'	end if
'end if

' response.write (lcase(replace(del_zip," ", ""))) & "   shipping 1 == " & ship1 & "   shipping 2 == " &ship2
' response.end
totalCnt=0
k=1


						
For i=0 To UBound(arr)
	strSql= "SELECT ItemId, fDescription, Price FROM Products WHERE (ItemId= "  & arr(i,0) & ");"
	Set rs = Server.CreateObject("ADODB.Recordset")
		
	rs.open strSql,c
				
	if(not rs.eof) then 
		p= arr(i,2)'getPrice(rs("ItemId"),c,arr(i,3))
		'arr(i,2)=p
		t=""
		if(applyDiscount) then
			if(session("specialxprice")<>""	) then 
				p= cdbl(session("specialxprice"))
				t= ": Prix spècial de promotion : " & session("promocode") & " : " & p
			end if
		end if
'p=0.10		
 %>
		<input type="hidden" name="item_name_<%=(i+1)%>" value="<%= rs("fDescription")%><%=t%>">
		<input type="hidden" name="amount_<%=(i+1)%>" value="<%=p%>">
		<input type="hidden" name="quantity_<%=(i+1)%>" size="3" value="<%=arr(i,3)%>"><br>
		<!--<input type="hidden" name="shipping_<%=(i+1)%>" size="3" value="<%=(ship1/(ubound(arr)+1))%>"><br>-->
<%

		
		gt=gt + cdbl(p) * cdbl(arr(i,3))
		if(Session("FreexQty")<>"") then
   	    'if(arr(i,3)>19) then 
			k=k+1
			arr(i,2)=0
			if(int(cdbl(arr(i,3))/cint(Session("FreexQty")) )>0) then
%>
			<input type="hidden" name="item_name_<%=UBound(arr)+k%>" value="<%= " gratuit .." & rs("fDescription")%>">
			<input type="hidden" name="amount_<%=UBound(arr)+k%>" value="0">
			<input type="hidden" name="quantity_<%=UBound(arr)+k%>" size="3" value="<%=int(cdbl(arr(i,3))/cint(Session("FreexQty")) )%>"><br>
			<!--<input type="hidden" name="shipping_<%=UBound(arr)+k%>" size="3" value="0"><br>-->
<%			InstQ = "INSERT INTO OrderDetails(OrderId,ProductId,Quantity,Price) " 
				InstQ = InstQ & "VALUES(" & OrderNum & "," & rs("ItemId") & ", " & int(cdbl(arr(i,3))/cint(Session("FreexQty")) ) & ",0)" 
				c.execute InstQ	

  	    end if
		

  	    end if
		
		'store itemid for later use
		totalCnt= totalcnt + cdbl(arr(i,3))
		totalTax= totalTax + round(rs("Price")*(taxrate)/100,2)
				
	end if
		
	'**********************************
	'RAMY insert ORDERDETAILS table using orderid
	'INSERT INTO OrderDetails ( ProductId, Quantity, Price,ORDERID ) 
	'******************************
	InstQ = "INSERT INTO OrderDetails(OrderId,ProductId,Quantity,Price) " 
	InstQ = InstQ & "VALUES(" & OrderNum & "," & rs("ItemId") & ", " & arr(i,3) & "," & p & ")" 
	c.execute InstQ	
	rs.close
Next

	if(totalCnt>6) then
	  ship1=ship2
	end if
	if(Session("freexship")) then  ship1=0
	'if(del_state="NT" or del_state="YT" or del_state="NU") then ship1=0

%>
<!-- Shipping  <input type="text" name="shipping9" value="<%=ship1%>"> -->
<%
 
taxrate= cdbl(taxrate) * (1 + cdbl(ship1)/cdbl(gt))
				
For i=0 To UBound(arr)
%>
 <input type="hidden" name="tax_<%=(i+1)%>" value="<%=round(cdbl((taxrate))* cdbl(arr(i,2))/100,2)%>">
 <input type="hidden" name="shipping_<%=(i+1)%>" size="3" value="<%=(ship1/(ubound(arr)+1))%>"><br>
<%
next
%>

<input type="hidden" name="tax_cart" value="<%=round((m_taxrate* (gt+ship1)/100),2)%>"> 
<!-- Total Tax -->

<input type="hidden" name="invoice" value="<%=OrderNum%>">
<input type="hidden" name="custom" value="<%if(applyDiscount) then response.write encode(session("promocode"))%>">
<input type="hidden" name="currency_code" value="CAD">
<input type="hidden" name="return" value="https://www.vsl3.ca/Thank-you.asp?token=<%=OrderNum%>">	
<input type="hidden" name="cancel_return" value="https://www.vsl3.ca/french/Cancel_order.asp?token=<%=OrderNum%>&txt=<%=CustId%>">
	

<!--	
<br> Invoice <br>
<input type="text" name="inv_name" value="<%=inv_name & " " & Inv_FirstName %>">
<input type="text" name="inv_company" value="">
<input type="text" name="inv_addr1" value="<%=inv_addr1%>">
<input type="text" name="inv_addr2" value="<%=inv_addr2%>">
<input type="text" name="inv_state" value="<%=inv_state%>">
<input type="text" name="inv_zip" value="<%=inv_zip%>">
<input type="text" name="inv_country" value="<%=inv_country%>">
<input type="text" name="tel" value="<%=tel%>">
<input type="text" name="email" value="<%=email%>">

<br> Shipping <br>
<input type="text" name="del_name" value="<%=del_name & " " & del_FirstName %>">
<input type="text" name="del_company" value="">
<input type="text" name="del_addr1" value="<%=del_addr1%>">
<input type="text" name="del_addr2" value="<%=del_addr2%>">
<input type="text" name="del_state" value="<%=del_state%>">
<input type="text" name="del_zip" value="<%=del_zip%>">
<input type="text" name="del_country" value="<%=del_country%>">
<input type="text" name="del_tel" value="<%=del_tel%>">

<br><br>
<input type="submit" name="btnSubmit" value="ShoppingCart99">
-->
<%
'**********************************
'RAMY update ORDER table using orderid
'INSERT INTO Order ( Amount, tax, shipping) using ORDERID 
'******************************
strSQL = "UPDATE Orders SET  Amount=" & gt + ship1 + round((m_taxrate* (gt+ship1)/100),2) &",tax="&round((m_taxrate* (gt+ship1)/100),2)& ",shipping=" & ship1
if(applyDiscount) then  strSQL = strSQL & ", PromoCodeUsed='" & session("promocode") & "'  "
strSQL = strSQL & " WHERE OrderId =" & OrderNum & ";" 
'response.write strsql
c.execute strSQL	

Session("freexship")=false
Session("specialxprice")=""
Session("FreexQty")=""
promomsg=""
session("promocode")=""

Session("orderid") = OrderNum
%>

<div class="t" align="center" valign="middle">
<table width="400" border="0" cellspacing="0" cellpadding="0">
<tr>
<td valign="middle" align="center">D&eacute;tournant. Attendez s.v.p</td><td valign="middle" align="center"><img src="images/loading.gif"></td>
</tr>
</table>
</div>
</form>
<%
c.execute  "update Customers set NewCustomer =false WHERE (((Customers.CustomerID)=" & Security.CurrentUserID & "));" 
%>
<script type="text/javascript" language="JavaScript"><!--
 document.MyForm.submit();
//--></script>
<%

end if
c.close



function getPrice(ItemId,c,q)
	dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")

	rs.Open  "SELECT PRICE,Price_Rebate FROM Products WHERE (Products.[ItemId]=" & ItemId & ") ;", c, 1, 2 
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


function checkCustomer(c,code)
	if(code<>"") then
		code = UCase(code)
		dim rs,NewCustomer
		Set rs = Server.CreateObject("ADODB.Recordset")
		strSql ="SELECT Discountcodes.DiscountCode, DiscountTypes.fDiscountTitle, DiscountTypes.DiscountType, DiscountTypes.freeShipping, DiscountTypes.FreePerQty, DiscountTypes.SpecialPrice, Orders.PromoCodeUsed "
		strSql = strSql & " FROM (Discountcodes INNER JOIN DiscountTypes ON Discountcodes.DiscountTypeId = DiscountTypes.DiscountTypeId) LEFT JOIN Orders ON Discountcodes.DiscountCode = Orders.PromoCodeUsed "
		strSql = strSql & " WHERE (((Discountcodes.DiscountCode)='"& mid(code,1,5) &"') AND ((Discountcodes.Active)=True)) AND ((DiscountTypes.StartDate)<Now()) AND ((DiscountTypes.EndDate)>Now()) "
		if(code<>"VSL14") then strSql = strSql & " AND((Discountcodes.used)=False) "
	'	response.write strsql
		rs.Open strSql, c, 1, 2 
		
		if(not rs.eof) then 
			promoCodeused=rs.fields("PromoCodeUsed")
			if(promoCodeused & "x"<>"x") then promoCodeused=UCase(promoCodeused)
			if((promoCodeused & "x"="x") or (promoCodeused ="VSL14"))  then
			'if(true)  then
				Session("freexship")=rs.fields("freeShipping")
				Session("specialxprice")=rs.fields("SpecialPrice")
				Session("FreexQty")=rs.fields("FreePerQty")
				checkCustomer=true
				promomsg="Promo code Applied : " & rs.fields("fDiscountTitle")
			else
				Session("freexship")=false
				Session("specialxprice")=""
				Session("FreexQty")=""
				checkCustomer=false
				promomsg="Promo Code is used and order is currently in process of being paid. If payment fails code will be unlocked within 24hrs."
				session("promocode")=""
			end if
		else
			Session("freexship")=false
			Session("specialxprice")=""
			Session("FreexQty")=""
			checkCustomer=false
			promomsg="Invalid promo Code"
			session("promocode")=""
		end if
		rs.close
		set rs=nothing
	else
		checkCustomer=false
	end if

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

' USE: location.href="nextpage.asp?" & encode("sParm=" & sData)
Function Encode(sIn)
dim x, y, abfrom, abto
Encode="": ABFrom = ""
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
If y = 0 Then
Encode = Encode & Mid(sin, x, 1)
Else
Encode = Encode & Mid(abto, y, 1)
End If
Next
End Function
%>