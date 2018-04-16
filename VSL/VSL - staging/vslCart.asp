<%@ Language=VBScript %>
<%
Const EW_PAGE_ID = "list"
Const EW_TABLE_NAME = "Products"%>
<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"--> 
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->
<!--#include file="userfn60.asp"-->
<% 
	dim errormsg, vNarc,NumberOfCells
	vNarc=False
	NumberOfCells=5

%>

<!--#include file="header.asp"-->

<%
dim shiphtml,linkhtml
shiphtml=""
linkhtml=""
Dim sLink,fldName
sLink= Request("link")

tqty=0

dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING

select case sLink
	case "del"
		RemoveItemFromCart(Request.querystring("ItemId"))
	case "clear"
		DestroyCart()
	case "checkout"
		VerifyCart()
	case "finalize"
		FinalizeCart()
	case "view"
		DisplayItems()
		writeFooter()	
		response.end
	case "holdOrder"
		SaveHeld()
	case "update" 
		if(IsCartEmpty()) then response.redirect "vslCart.asp"
	
		For Each fldName in Request.Form
			if(isnumeric(fldName)) then tqty=tqty + cint(request.form(fldName))
		Next
		
		call CheckQty(tqty,"vslCart.asp?link=view")
		
		For Each fldName in Request.Form
		
			if(isnumeric(request.form(fldName))) then 
				UpdateItemToCart fldName,request.form(fldName)
			end if
		Next
end select


dim strItemId,strUnit_Qty,strDESCRIPTION,strPrice

	If Not CartExists() Then CreateCart()
	
	For Each fldName in Request.Form
		if(left(fldName,6)="ItemId") then 
			strUnit_Qty=Request.Form(right(left(fldName,9),3) & "_Qty")
			if(not isnumeric(strUnit_Qty)) then	strUnit_Qty=0
			tqty=tqty + strUnit_Qty
		End if
		
	Next
	call CheckQty(tqty,"vslOrderForm.asp")
			
	For Each fldName in Request.Form
		if(left(fldName,6)="ItemId") then 
			strItemId = Request.Form(fldName)
			strUnit_Qty=Request.Form(right(left(fldName,9),3) & "_Qty")
			strDESCRIPTION=Request.Form(right(left(fldName,9),3) & "_DESC")
			'strPrice=getPrice (strItemId,conn)
			if(not isnumeric(strUnit_Qty)) then	strUnit_Qty=0
			
			
			'if(cint(strUnit_Qty)>1 or (isNewCustomer())) then 
				Call AddItemToCart(strItemId,strDESCRIPTION,strPrice,strUnit_Qty)
			'else
			''	call CheckQty(strUnit_Qty,"vslOderForm.asp")
			'end if
		End if
		
	Next
	CorrectPrice
	if( request.form("acart")="addcart") then response.redirect Request.ServerVariables("HTTP_REFERER") 
	DisplayItems()
	writeFooter()	

	
	
Sub CheckQty(strUnit_Qty,rUrl)

	if((cint(strUnit_Qty)=1) and (not isNewCustomer())) then 

		Session(EW_SESSION_MESSAGE)="Minimum Qty required is 2"
		response.redirect (rUrl)
	end if
End Sub

Sub DisplayItems()
	dim arr
	arr=GetCart()
	if(uBound(arr)=-1) then
		message "emptycart"
	else

		cartItemCount=UBound(arr)
	
		For i=0 to cartItemCount
			if(cdbl( arr(i,3))=0) then RemoveItemFromCart(arr(i,0))
		Next
		arr=GetCart()
		writeTopTable()
		'response.Write "<b>Your Cart:</b>"
%> 
<table width="820"  border="0" cellpadding="0" cellspacing="0" id="Table_01">
  <tr>
   <td width="680"><img src="images/title_yourcart.png" width="410" height="75"></td>
   <!-- <td width="28" valign="top"><img src="images/fontsize.png" border="0" alt=""></td>
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
      <p><a href="french/vslCart.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
    </div></td>-->
    </tr>
    <tr>
   <td width="680"><p style="color:#333333; font-weight: bold;">Please note:   VSL#3 orders are shipped Monday, Tuesday and Wednesday's only.<br/>  
Any orders received Wednesday after 10am EST; up to and including Friday, will ship the following Monday.<br/>
</p></td>
    </tr>
</table>

<form method="post" action="vslCart.asp"> 
<div class="t">
  <div align="right"><span class="vslcss"><a href="VSLOrderForm.asp">Back to Products</a> :
          <%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
	    
        <a href="login.asp">Login</a>
          <%else%>
	      <a href="Customersedit.asp">Edit account</a> :
		   <a href="changepwd.asp">Change Password</a> :
          <a href="logout.asp">Logout</a>
          <%end if
	set Security =nothing %>
          <img src="images/spacer.gif" width="25" height="10">
      </span>   </div>   <%
If Session(EW_SESSION_MESSAGE) <> "" Then
%>
  
  <p class="ewmsg"><%= Session(EW_SESSION_MESSAGE) %></p>
<%
	Session(EW_SESSION_MESSAGE) = "" ' Clear message
End If
%>
<table width="775"  class="ewTablecart" cellpadding="5" cellspacing="5" border="1">
  <tr >
    <td width="52%" class="ewTableHeader"> <b>Products</b></td>
    <td width="16%" class="ewTableHeader">
    <div align="right"><b>Unit Price</b></div></td>
    <td width="12%" class="ewTableHeader">
    <div align="center"><b>Quantity</b></div></td>
    <td width="10%" class="ewTableHeader"><div align="right"><b>Total</b></div></td>
    <td width="10%" class="ewTableHeader">&nbsp;</td>
  </tr>
  <%For i=0 To UBound(arr)%>
  <tr >
    <td width="52%" style="padding:5px;"><b><%= "</b>" & arr(i,1)%></b></td>
    <td width="16%" style="padding:5px;">
      <div align="right"><b><%=arr(i,2)%> </b></div></td>
    <td width="12%" align="center"style="padding:5px;">
      <div align="center">
        <%if(sLink="checkout" or sLink="finalize") then
	   		
				Response.write arr(i,3)
			
		else%>
			<input type="text" name="<%=arr(i,0)%>" size="2" value="<%=arr(i,3)%>">
        <%
		end if%>
    </div></td>
    <td width="10%"style="padding:5px;"><div align="right"><%=FormatCurrency(arr(i,2) *  arr(i,3),2)%></div></td>
    <td width="10%"style="padding:5px;">
      <div align="center"><b><a href='vslCart.asp?ItemId=<%=arr(i,0)%>&link=del' onClick="return confirm('delete item: <%=arr(i,0)%> from cart ?')"><b>Delete</b></a>
            <input type="hidden" name="txtproductID" value="<%=arr(i,0)%>">
    </b></div></td>
  </tr>
  <%
  'call extraUnits(arr(i,3),"<tr><td colspan=5 class='ewmsg' align='right'>+2 free packages will be added to your shipment  when you checkout</td></tr>")
  Next%>
  <tr bordercolor="#FFCC66">
    <td width="52%" class="ewTablePager">
    <% if (isNewCustomer()) then%> 
    	<span class="ewmsg">Please note ,first box will be charged 99 instead of 110 when you go to the paypal payment.</span>
    <%end if %>
    	
    </td>
    <td width="16%" class="ewTablePager">&nbsp;</td>
    <td width="12%" class="ewTablePager"><div align="right">Total:</div></td>
    <td width="10%" class="ewTablePager"><div align="right"><%=FormatCurrency(FormatNumber(TotalPurchase(),2),2)%></div></td>
    <td width="10%" class="ewTablePager">&nbsp;</td>
  </tr>
  <tr bordercolor="#FFCC66">
    <td><input type="hidden" name="link" value="update"> 
        <input name="Submit" type="image" class="InputNoBorder" value="Recalculate Total"  src="images/recalculate.gif" width="135" height="32"></td>
    <td colspan="4" nowrap><a href="vslCart.asp?link=checkout"><img src="images/checkout.gif" width="115" height="32" border="0"></a> <a href="vslCart.asp?link=clear" onClick="return confirm(' Empty all items from the cart ?');"><img src="images/cancelorder.gif" width="139" height="32" border="0"></a></td>
    </tr>
</table>
</div>
</form> 
<%
WriteBottomTable()
end if
End Sub
%> 

<%
Sub CreateCart()
	dim arrCart
	ReDim arrCart(-1,NumberOfCells)
	Session("cart")=arrCart
End Sub

Sub EmptyCart()
	DestroyCart()
	CreateCart()
End Sub

Sub DestroyCart()
	dim arrCart
	ReDim arrCart(-1,NumberOfCells)
	Session("cart")=arrCart
End Sub

Sub VerifyCart()
	isCartLogin()
	'MinQtyCheck()
	if(QtyCheck) then 
		response.redirect ("vslCart.asp?link=finalize")
	else
		EmptyCart()
		response.redirect ("VSLOrderForm.asp")
	end if
End Sub

function QtyCheck()
	
	Dim arrCart,cartItemCount,TotalQty
	arrCart=Session("cart")
	cartItemCount=UBound(arrCart)
	TotalQty=0
	For i=0 to cartItemCount
		TotalQty=TotalQty+ cdbl( arrCart(i,3))
	Next
	QtyCheck=not (TotalQty=0) 
End function

Sub isCartLogin()
	' Verify Login
	Dim Security
	Set Security = New cAdvancedSecurity
	
	If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
	If Not Security.IsLoggedIn() Then
		Call Security.SaveLastUrl()
		Call Page_Terminate("login.asp")
	End If
End Sub

Sub AddItemToCart(strItemId,strDESCRIPTION,strPrice,strUnit_Qty)
	Dim arrTemp	
	Dim arrCart
	dim cartItemCount,blnItemExists
	If Not CartExists() Then CreateCart()
	arrCart=Session("cart")
	cartItemCount=UBound(arrCart)

	strPrice=getPrice(strItemId,conn,cint(strUnit_Qty))
	if(strPrice="-1") then exit sub
' Create a error message for the add from saved cart.

	blnItemExists=False
	intItemIndex=-1
	For i=0 to cartItemCount
		If arrCart(i,0)=strItemId Then 
			blnItemExists=True
			intItemIndex=i
			Exit For
		End If
	Next
	
	If Not blnItemExists Then
		ReDim arrTemp(Ubound(arrCart)+1,NumberOfCells)

		For i=0 To cartItemCount
			arrTemp(i,0)=arrCart(i,0)
			arrTemp(i,1)=arrCart(i,1)
			arrTemp(i,2)=arrCart(i,2)
			arrTemp(i,3)=arrCart(i,3)
		Next
		arrTemp(cartItemCount+1,0)=strItemId
		arrTemp(cartItemCount+1,1)=strDESCRIPTION
		arrTemp(cartItemCount+1,2)=strPrice
		arrTemp(cartItemCount+1,3)=cdbl(strUnit_Qty)
		Session("cart")=arrTemp
	Else
		arrCart(intItemIndex,0)=strItemId
		arrCart(intItemIndex,1)=strDESCRIPTION
		arrCart(intItemIndex,2)=strPrice
		arrCart(intItemIndex,3)=cdbl(strUnit_Qty)
		Session("cart")=arrCart
	End If

End Sub

Sub UpdateItemToCart(strItemId,strUnit_Qty)

	Dim arrTemp	
	Dim arrCart
	dim cartItemCount,blnItemExists
	
	arrCart=Session("cart")
	cartItemCount=UBound(arrCart)

	blnItemExists=False
	intItemIndex=-1
	For i=0 to cartItemCount
		If arrCart(i,0)=strItemId Then 
			blnItemExists=True
			intItemIndex=i
			Exit For
		End If
	Next
	
	If  blnItemExists Then
		arrCart(i,3) =cdbl(strUnit_Qty)
		arrCart(i,2) =getPrice(strItemId,conn,cdbl(strUnit_Qty))
		Session("cart")=arrCart
	end if
		
End Sub

Sub RemoveItemFromCart(strID)

	Dim arrTemp, arrCart
	Dim j, blnMatch

	If Not IsCartEmpty() Then
		arrCart=Session("cart")
		ReDim arrTemp(UBound(arrCart),NumberOfCells)
		j=0			
		blnFoundItem=False
		For i=0 To UBound(arrCart)
			If arrCart(i,0)=strID Then
				'skip this item
				blnFoundItem=True
			Else
				arrTemp(j,0)=arrCart(i,0)
				arrTemp(j,1)=arrCart(i,1)
				arrTemp(j,2)=arrCart(i,2)
				arrTemp(j,3)=arrCart(i,3)
				j=j+1				
			End If
		Next

		If blnFoundItem Then 
			ReDim arrCart(UBound(arrTemp)-1,NumberOfCells)
			For i=0 To UBound(arrCart)
				arrCart(i,0)=arrTemp(i,0)
				arrCart(i,1)=arrTemp(i,1)
				arrCart(i,2)=arrTemp(i,2)
				arrCart(i,3)=arrTemp(i,3)
			Next
			Session("cart")=arrCart
		End If
	End If	

End Sub

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

Sub FinalizeCart()
	
	isCartLogin()
	PrepareCart()
'generate form ship

	
End Sub


Sub PrepareCart()

	response.redirect("checkout.asp")
%>
<form action="http://ww11.aitsafe.com/cf/addmulti.cfm" name="qForm" method="post"> 
    <input type="hidden" name="useridsd" value="D6155807"> 
    <input type="hidden" name="return" value="http://localhost/vsl/cart.asp"> 
	<%For i=0 To UBound(arr)%>
	<%=arr(i,0)%>
    <input type="text" name="qty1" size="3" value="<%=arr(i,0)%>"> q
    <input type="hidden" name="product1" value="<%=arr(i,0)%>"> 
    <input type="hidden" name="price1" value="<%=arr(i,1)%>"> 
    <input type="hidden" name="tax1" value="<%=arr(i,0)%>"> 
	
<%next%>
	<p><input name="submit" type="image" value="Buy Now!" src="images/finalizeorder.gif" width="139" height="32">
	</p> 
</form>

<%

End Sub

Sub PrepareCartx()
dim arr
	arr=GetCart()
	
%>
<form action="http://ww11.aitsafe.com/cf/addmulti.cfm" name="qForm" method="post"> 
    <input type="hidden" name="useridsd" value="D6155807"> 
    <input type="hidden" name="return" value="http://localhost/vsl/cart.asp"> 
	<%For i=0 To UBound(arr)%>
	<%=arr(i,0)%>
    <input type="text" name="qty1" size="3" value="<%=arr(i,0)%>"> 
    <input type="hidden" name="product1" value="<%=arr(i,0)%>"> 
    <input type="hidden" name="price1" value="<%=arr(i,1)%>"> 
    <input type="hidden" name="tax1" value="<%=arr(i,0)%>"> 
	
<%next%>
	<p>
	  <input name="submit2" type="image" value="Buy Now!" src="images/finalizeorder.gif" width="139" height="32">
	</p> 
</form>
<script language="JavaScript" type="text/JavaScript">
document.qForm.submit();
</script>
<%

End Sub

sub writeFooter()

%>

  </div>
</div>

<!-- InstanceEndEditable -->


<!-- Footer -->
<div class="footer-wrapper">
<div class="container">
<div class="row "> 
 
 <div class="col-sm-12 col-md-12 text-center tb-pad">
 		<h6 style="line-height:20px;"><large>Questions? Call 1.800.263.4057 or <br id="brView"><a href="contact-us.php" class="blue-link">Fill Out This Form</a></large><br>
        <a href="full-Product-information.html" class="footerlink">Full Product Information</a> | <a href="http://www.ferring.ca" target="_blank" class="footerlink">About Ferring</a> | <a href="legal-notice.html" class="footerlink">Legal Notice</a> | <a href="sitemap.html" class="footerlink">Site Map</a> | <a href="letters-to-insurance.html" class="footerlink">Letters to insurance companies for reimbursement</a> | <a href="contact-us.php" class="footerlink">Contact Us</a><br>
          <span class="bluehighlight">This website is intended only for Canadian residents.</span><br>
		  <span>Natural Product Number NPN 80037590</span></h6>
      </div>
   <div class="col-sm-6 col-md-6"> 
     <div id="ferringSpace">&nbsp;</div>
  <div id="ferringLogo"><img src="_images/ferring-logo.jpg" width="80" alt="Ferring Pharmaceuticals"></div>
  <div id="ferringCopy"><h6>Copyright © 2014. Ferring Canada. All rights reserved. <br>
    200 Yorkland Boulevard, Suite 500 North York <br>
    Ontario Canada M2J 5C1</h6></div>
</div>
  
  <div class="col-sm-6 col-md-6"> 
  	  <div id="ferringSpaceR">&nbsp;</div>
  	  <div id="ferringRight"><h6>VSL#3<sup>®</sup> and The Living Shield are registered trademarks of VSL Pharmaceuticals Inc.<br id="brView">
		VSL#3<sup>®</sup> is a probiotic blend that is intended to be used under the supervision of a doctor.<br id="brView">
		Please consult with your doctor before trying VSL#3<sup>®</sup></h6></div>
  </div>
<div class="clearfix"></div>

 <div class="col-sm-12 col-md-12 text-right rsgpad"><h6>Website Design By <a href="http://www.ravenshoegroup.com" target="_blank" class="blue-link">Ravenshoe Group</a></h6></div>

</div>
</div>
</div><!-- footer end -->



</div><!-- wrapper end -->

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="_js/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="_js/bootstrap.min.js"></script>

<!-- InstanceBeginEditable name="footer" -->
<!-- InstanceEndEditable -->

</body>
<!-- InstanceEnd --></html>

<% 
end sub

function iif(c,t,f)
if(c) then
	iif=t
else
	iif=f
end if
end function

sub WriteTopTable()

end sub

sub WriteBottomTable()
exit sub
Response.Write("</td><td background=""images/Box_mr.gif""  ></td>" & vbcrlf) 
Response.Write("  </tr>" & vbcrlf) 
Response.Write("  <tr>" & vbcrlf) 
Response.Write("    <td><img src=""images/Box_bl.gif"" width=""16"" height=""16""></td>" & vbcrlf) 
Response.Write("    <td background=""images/Box_bm.gif""><img src=""images/Box_bm.gif"" width=""24"" height=""16""></td>" & vbcrlf) 
Response.Write("    <td><img src=""images/Box_br.gif"" width=""16"" height=""16""></td>" & vbcrlf) 
Response.Write("  </tr>" & vbcrlf) 
Response.Write("</table>") 

end sub

sub prx(m)
response.write m
response.End()
end sub
function getSearchVis(adv)
if((request.QueryString("psearch")="") Eqv  (adv)) then
 getSearchVis="Visible"
else
 getSearchVis="hidden"
end if
end function



Sub Message(t)
	'Response.Write("<table width=""790"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf) 
	'Response.Write("                <tr>" & vbcrlf) 
	'Response.Write("                  <td align='right'><a href=""french/vslCart.asp"" class=""bodycopy_small"">en fran&ccedil;ais &gt;</a></td>" & vbcrlf) 
	'Response.Write("                </tr>" & vbcrlf) 
	'Response.Write("                <tr>" & vbcrlf) 
	'Response.Write("                  <td><div class='t' align='right'>") 
	    
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then
	%>
       <a href="login.asp">Login</a>
    <% else %>
	    <a href="Customersedit.asp">Edit account</a> :
		<a href="changepwd.asp">Change Password</a> :
        <a href="logout.asp">Logout</a>
    <%
	end if
	set Security =nothing 
	
	select case  t
	case "emptycart"
		Response.Write("<table width=""790"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF""> " & vbcrlf) 
		Response.Write("        <tr> " & vbcrlf) 
		Response.Write("          <td width=""14"" height=""40"">&nbsp;</td> " & vbcrlf) 
		Response.Write("          <td width=""701"" align=""left"" valign=""top"" class=""bodybold"">Your Cart is empty. Please add items.<br> " & vbcrlf) 
		Response.Write("            <br> " & vbcrlf) 
		Response.Write("            <a href='VSLOrderForm.asp?cmd=resetall'><img src=""images/clicktoreturn.gif""  border=0></a><br> " & vbcrlf) 
		Response.Write("            </p> </td> " & vbcrlf) 
		Response.Write("        </tr> " & vbcrlf) 
		Response.Write("      </table>") 
	case "saveheld"
		Response.Write("<table width=""701"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF""> " & vbcrlf) 
		Response.Write("        <tr> " & vbcrlf) 
		Response.Write("          <td width=""14"" height=""40"">&nbsp;</td> " & vbcrlf) 
		Response.Write("          <td width=""701"" align=""left"" valign=""top"" class=""bodybold"">Your Cart has been saved.<br> Date:" & Now & "<br> " & vbcrlf) 
		Response.Write("            <br> <br>" & vbcrlf) 
		Response.Write("            <a href='vslCart.asp?link=getHeldOrders'><img src=""images/clicktoreturn.gif"" width=""181"" height=""23"" border=0></a><br> " & vbcrlf) 
		Response.Write("            </p> </td> " & vbcrlf) 
		Response.Write("        </tr> " & vbcrlf) 
		Response.Write("      </table>") 
	end select
	
	Response.Write("</div>                  </td>" & vbcrlf) 
	Response.Write("                </tr>" & vbcrlf) 
	Response.Write("            </table>") 
end sub

%> 
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



	conn.Close ' Close Connection
	Set conn = Nothing
	Set Security = Nothing
	Set Customers = Nothing

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

Function CorrectPrice()
	Dim arrTemp	
	Dim arrCart
	dim cartItemCount,blnItemExists
	If Not CartExists() Then CreateCart()
	arrCart=Session("cart")
	cartItemCount=UBound(arrCart)

' Create a error message for the add from saved cart.
	t=0
	intItemIndex=-1
	For i=0 to cartItemCount
		if(cdbl(arrCart(i,2))>0) then  t= cdbl(arrCart(i,3)) + t
	Next
		
	if(t>4) then
		For i=0 to cartItemCount
			if(cdbl(arrCart(i,2))>0) then	arrCart(i,2) =getPrice(arrCart(i,0),conn,t)
		Next
		
	end if
	Session("cart")=arrCart
end Function

%>