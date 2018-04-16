<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->

<html>
<head>
<title>VSL#3&reg; - The Living Shield&reg;</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Description" content="Canada's source of information on VSL#3, a high potency probiotic blend specially formulated for care of patients with serious intestinal disorders." />
<style type="text/css">
<!--
body {
	background-image: url(images/bgnd.jpg);
}
A {
	color:#1070A3;
	text-decoration:underline;
	font-weight: normal;
}
A:hover {color:#333333; text-decoration;}
-->
</style>
<link href="imagemenu.css" rel="stylesheet" type="text/css" media="screen">
<link href="vsl.css" rel="stylesheet" type="text/css" media="screen">
<script type="text/javascript" src="mootools.js"></script>
<script type="text/javascript" src="imagemenu.js"></script>

 

<style type="text/css">


#kwick li a {
        color:#FFFFFF; }

.style1 {font-size: 10px}
</style>
<script type="text/javascript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (concept_v1_build.psd) -->
<div align="center">
  <table width="960"  border="0" align="center" cellpadding="0" cellspacing="0" background="images/topTemp.jpg">
  <tr>
    <td height="134">&nbsp;</td>
    <td><a href="index.asp"><img src="images/hitspot.png" width="250" height="100" border="0"></a></td>
  </tr>
  <tr>
    <td width="72" height="201">&nbsp;</td>
    <td width="895"><div id="Layer1" ><span id="kwick" >
            <ul id="kwicks">
		    <li class="kwick opt1" style="background: transparent url(images/page1.jpg) repeat scroll 0%; width:137px;"><a class="active" href="about.html" id="active0"></a>
		      
		    </li> 
		    <li class="kwick opt1" style="background: transparent url(images/page2.jpg) repeat scroll 0%; width:136px;"><a class="active" href="how.html" id="active0"></a></li> 
		    <li class="kwick opt1" style="background: transparent url(images/page3.jpg) repeat scroll 0%; width:137px;"><a class="active" href="faq.html" id="active0"></a></li> 
		    <li class="kwick opt1" style="background: transparent url(images/page4.jpg) repeat scroll 0%; width:137px;"><a class="active" href="letters.html" id="active0"></a></li> 
		    <li class="kwick opt1" style="background: transparent url(images/page5.jpg) repeat scroll 0%; width:137px;"><a class="active" href="VSLOrderForm.asp" id="active0"></a></li>
			<li class="kwick opt1" style="background: transparent url(images/page6.jpg) repeat scroll 0%; width:137px;"><a class="active" href="journal.html" id="active0"></a></li>
    </ul> </span></div></td>
  </tr>
  <tr>
    <td height="8" ></td>
    <td></td>
  </tr>
</table>
  <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="90" height="25">&nbsp;</td>
          <td width="820" height="25"><table width="820" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="700" height="50" align="left" valign="middle">&nbsp;</td>
              <td width="120" align="left" valign="top"><div align="right" class="bodycopy_small">
                <p>&nbsp;</p>
                <p><a href="french/journal.html">en fran&ccedil;ais &gt;</a></p>
              </div></td>
            </tr>
            <tr>
              <td height="300" colspan="2" align="left" valign="top">
			  <table width="820" border="0" cellspacing="0" cellpadding="0">
                <tr align="left" valign="top">
                  <td width="620" class="bodycopy">
				  	    <%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
	
            <a href="login.asp">login</a>
            <%else%>
			<a href="Customersedit.asp">Edit account</a> : <a href="changepwd.asp">Change Password</a> :
            <a href="logout.asp">logout</a>
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
					<div style="font-family: Verdana;font-size: 14px; padding-top:10px;" ><strong>Thank you for your payment.</strong> </div>
					<div style="font-family: Verdana;font-size: 12px;" class="invoiceTotal">
						Your transaction has been completed, and a receipt for your purchase <br />
						has been emailed to you from Paypal.
					</div>
					</i>
					<div style="padding-left: 3px;padding-top:24px;">
					<a href="invoice.asp?token=<%=orderid %>" target="_blank"><span class="submitbutton"><u>Print Invoice</u></span></a>
					</div>
					
				  </td>
                  <td width="200" ><img src="images/packshot_v2.png" width="200" height="200"></td>
                </tr>
              </table></td>
            </tr>
            <tr><td colspan="2"></td></tr>
          </table></td>
          <td width="90">&nbsp;</td>
        </tr>
      </table>      </td>
    </tr>
  </table>
  <br /><br />
  <!-- End ImageReady Slices -->
  <table width="1000" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="25">&nbsp;</td>
      <td height="25" class="footerlink">&nbsp;</td>
      <td width="350" height="25" align="right" valign="middle" class="footer"><p>This website brought to you by: </p>
          <p>Canadian distributor of VSL#3&nbsp; </p></td>
      <td width="60" align="center" valign="middle" class="footer"><img src="images/ferring_logo_web.png" width="58" height="24" border="0"></td>
      <td height="25" class="footer">&nbsp;</td>
    </tr>
    <tr>
      <td width="90" height="25">&nbsp;</td>
      <td width="410" height="25" class="footerlink"><p>&nbsp;</p>
          <p><a href="legal.asp">Legal Notice</a></p></td>
      <td width="410" height="25" colspan="2" align="right" valign="middle" class="footer"><table width="283" height="28" border="0" cellpadding="0" cellspacing="0">
        <tr align="right" valign="middle">
          <td width="156" height="28" class="rsgslug"><a href="http://www.ravenshoegroup.com/custom-website-design.html" target="_blank" class="rsgstyle">Website Design</a> by: </td>
          <td width="27" height="28" align="right"><span class="rsgslug"><img src="images/rsglogo-icon.png" width="19" height="25" alt="Ravenshoe Group" /></span></td>
          <td width="100" height="28" align="left" class="rsgslug"><a href="http://www.ravenshoegroup.com/" target="_blank" class="rsgstyle">Ravenshoe Group</a></td>
        </tr>
      </table></td>
      <td width="90" height="25">&nbsp;</td>
    </tr>
  </table>
  <script type="text/javascript">
    var myMenu = new ImageMenu($$('#kwick .kwick'),{openWidth:261,start:3});
  </script>
</div>
<map name="Map">
  <area shape="rect" coords="34,16,258,107" href="index.asp">
</map>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
var pageTracker = _gat._getTracker("UA-1731303-15");
pageTracker._trackPageview();
</script>
</body>
</html>