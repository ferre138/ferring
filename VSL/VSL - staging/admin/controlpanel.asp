<!--METADATA TYPE="typelib"
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
NAME="CDO for Windows 2000 Library" -->
<!--METADATA TYPE="typelib"
UUID="00000205-0000-0010-8000-00AA006D2EA4"
NAME="ADODB Type Library" -->
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<script language="javascript" type="text/javascript" src="tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript">
tinyMCE.init({
        theme : "advanced",
        mode : "exact",
		elements : "EmailText",
		        theme_advanced_buttons1 : "bold,italic,underline,forecolor,backcolor",
        theme_advanced_buttons2 : "",
        theme_advanced_buttons3 : ""

});
function reloadtxt(t)
{etxt="";
	if(t.value=="confirm"){
		etxt=document.getElementById("ship_email").innerHTML;
	}
	else
	{
		etxt=document.getElementById("fail_email").innerHTML;
	}
	
tinyMCE.get('EmailText').setContent(etxt);
}
</script>
<style type="text/css">
.cptable
{
	border: thin solid #999;
	border-spacing: 0;
	border-collapse: collapse;
	empty-cells: show;
	width:300px;height:250px;
	font-family: Verdana; /* font name */
	font-size: small; /* font size */
}

.cptable .ewTableHeader  td {
	background-color: #003366;	/* header bgcolor */
	color: #FFFFFF; /* header font color */
	border-bottom: 1px solid; /* header border width */
	border-right: 1px solid; /* header border width */
	border-color: #4E4F51; /* header border color */
	background-image: url(../images/darkglass.png); /* header bg image */
	background-repeat: repeat-x;
	vertical-align: top;height:30px;
	
}
.cptable .ewTableCaption  td {
	background-color: #cce5f0;	/* header bgcolor */
	color: #003366; /* header font color */
	border-bottom: 1px solid; /* header border width */
	border-right: 1px solid; /* header border width */
	border-color: #4E4F51; /* header border color */
	background-image: url(../images/darkglass.png); /* header bg image */
	background-repeat: repeat-x;
	vertical-align: middle;height:40px;
	
}
.pending {background-color: #FFFFCC; border-bottom: 1px solid;}
.cptablesummary
{
	border: thin solid #999;
	border-spacing: 0;
	border-collapse: collapse;
	empty-cells: show;
	width:1000px;
	font-family: Verdana; /* font name */
	font-size: medium; /* font size */
	font-weight: bold;
}

.cptablesummary .ewTableHeader  td {
	background-color: #003366;	/* header bgcolor */
	color: #FFFFFF; /* header border color */
	background-image: url(../images/darkglass.png); /* header bg image */
	background-repeat: repeat-x;
	vertical-align: top;
	border: thin solid #4E4F51;	
}
.cptablesummary .ewTableCaption  td {
	background-color: #cce5f0;	/* header bgcolor */
	color: #003366; /* header border color */
	background-image: url(../images/darkglass.png); /* header bg image */
	background-repeat: repeat-x;
	vertical-align: middle;
	border: thin solid #4E4F51;	
}
.cp {
	border: thin solid #999;
}
.test {
	font-family: Georgia, "Times New Roman", Times, serif;
}
</style>
<!-- #INCLUDE FILE="Includes/FusionCharts.asp" -->
<%

' Define page object
Dim custompage
Set custompage = New ccustompage
Set Page = custompage

' Page init processing
Call custompage.Page_Init()

' Page main processing
Call custompage.Page_Main()
%>
<!--#include file="header.asp"-->
<% custompage.ShowMessage %>
<% 
Dim strSQL, rst, orderid, totalAmount, counter
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING

Set rs = Server.CreateObject("ADODB.Recordset")
'Response.Write StatsGraph(200)


	Dim strXML
	sSql = "Select * from (SELECT top 30  OrderedDate,DateValue(Orders.Ordered_Date) AS OrderedDate, Sum(Orders.payment_gross) AS orders    "
	sSql = sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) GROUP BY DateValue(Orders.Ordered_Date)"
	sSql = sSql & "  ORDER BY DateValue(Orders.Ordered_Date) desc) order by 1 asc;" & vbCrLf
	rs.open sSql,conn					
	strXML = ""
	strXML = strXML & "<graph caption='Daily orders' xAxisName='Date' yAxisName='Amount' decimalPrecision='0' formatNumberScale='0' >"
		
	Do While not rs.EOF
		strXML = strXML & "<set name='" & day(rs("OrderedDate")) & "/" & month(rs("OrderedDate")) &"' value='" & rs("orders") &"' color='AFD8F8' />"
		rs.MoveNext						
	Loop
	strXML = strXML & "<set name='/' value='0' color='AFD8F8' />"
	strXML = strXML & "<set name='/' value='0' color='AFD8F8' />"
	strXML = strXML & "<set name='/' value='0' color='AFD8F8' />"
	strXML = strXML & "</graph>"
	rs.close
   
   'Create the chart - Column 3D Chart with data from strXML variable using dataXML method
  strHtml=""
	sSql = "SELECT TOP 5 Products.Description, Sum(OrderDetails.Price*OrderDetails.Quantity) AS SumOfPrice, Sum(OrderDetails.Quantity) AS SumOfQuantity, Products.Image "
	sSql=sSql & " FROM Products INNER JOIN (Orders INNER JOIN OrderDetails ON Orders.OrderId = OrderDetails.OrderId) ON Products.ItemId = OrderDetails.ProductId "
	sSql=sSql & " WHERE (((Orders.payment_status)='Completed')) GROUP BY Products.Description,Products.Image;"

	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4><h3>Products Sold..  </h3></td>"
		strHtml = strHtml &  "</tr>"

		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='middle' valign='top'>Product</td>"
	
		strHtml = strHtml & "<td align='middle' valign='top'>Count</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Total</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Avg</td>"
		strHtml = strHtml &  "</tr>"
		tot_qty=0
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		'strHtml = strHtml & "<td align='middle' valign='top'><img src='../products/thumbs/" & rs("Image") & "'></td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("Description") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("SumOfQuantity") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>$" & rs("SumOfPrice") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>$" & round(rs("SumOfPrice")/rs("SumOfQuantity"),2) & "</td>"

		strHtml = strHtml &  "</tr>"
		tot_qty= tot_qty + rs("SumOfQuantity")
		rs.MoveNext		
		
	Loop
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4>(excl tax,ship)</td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "</table>"
	rs.close
	m_strHtml=strHtml
%>

<table width="1000" border="0" align="left">
<tr>
    <td colspan="3" width=800><% Call renderChartHTML("Charts/FCF_Column3D.swf", "", strXML, "myNext", 1000, 250)%></td>
    </tr>

   <tr>
    <td colspan="3"><%
	sSql = "SELECT Sum(Orders.payment_gross) AS SumOfpayment_gross, Sum(Orders.payment_fee) AS SumOfpayment_fee, Sum(Orders.Shipping) AS SumOfShipping, Sum(Orders.Tax) AS SumOfTax, Count(Orders.OrderId) AS CountOfOrderId "
	sSql=sSql & "  FROM Orders "
	sSql=sSql & " WHERE (((Orders.payment_status)='Completed'));"
	rs.open sSql,conn	
	strHtml=""	
	strHtml = strHtml &  "<table width='100%' class='cptablesummary'>"
		'strHtml = strHtml &  "<tr class='ewTableCaption'>"
		'strHtml = strHtml & "<td align='middle' valign='middle' colspan=6><h2>Control Panel</h2></td>"
		'strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='middle' valign='top'>Orders</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Total Payments</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Paypal fees</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Paypal %</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Shipping</td>"
		'strHtml = strHtml & "<td align='middle' valign='top'>Province</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Tax</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Net/Pack</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("CountOfOrderId") & "</a></td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & FormatCurrency(rs("SumOfpayment_gross")) & "</a></td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & FormatCurrency(rs("SumOfpayment_fee")) & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & round(rs("SumOfpayment_fee")/rs("SumOfpayment_gross")*100,2) & "%</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & FormatCurrency(rs("SumOfShipping")) & "</td>"
		'strHtml = strHtml & "<td align='middle' valign='top'>" & rs("Ship_Province") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & FormatCurrency( rs("SumOfTax")) & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & FormatCurrency( (rs("SumOfpayment_gross")-rs("SumOfTax")-rs("SumOfShipping")-rs("SumOfpayment_fee"))/tot_qty) & "</td>"
		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml & "<td align='middle' valign='middle' colspan=5>&nbsp;</td>"
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml
%></td>
    </tr>
   <tr>
     <td rowspan="2" valign="top"><%

	strHtml=""
	sSql = "SELECT top 12 Orders.Ship_FirstName  & ' ' & Orders.Ship_LastName AS fullname, Orders.Ship_City, Orders.Ship_Province, Orders.OrderId, Orders.payment_gross, Orders.payment_status,CustomerId ,EmailSent"
	sSql=sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) ORDER BY EmailSent, Orders.OrderId DESC;"
	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4><h3>Recent 10 orders.. </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='middle' valign='top'>Orders</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Name</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>City</td>"
		'strHtml = strHtml & "<td align='middle' valign='top'>Province</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Total</td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='top' colspan=4>Not Confirmed</td>"
		strHtml = strHtml &  "</tr>"
		firstConfirm=0
	Do While not rs.EOF
		if(rs("EmailSent") & "x"="x") then
			strHtml = strHtml &  "<tr class='pending' title='confirmation not sent'>"
		else
			if(firstconfirm=0) then
				strHtml = strHtml &  "<tr >"
				strHtml = strHtml & "<td align='middle' valign='top' colspan=4>None pending</td>"
				strHtml = strHtml &  "</tr>"
				firstconfirm=1
			
			end if

			if((firstconfirm=1) and(firstconfirm<>2)) then
				strHtml = strHtml &  "<tr class='ewTableCaption'>"
				strHtml = strHtml & "<td align='middle' valign='top' colspan=4>Confirmed</td>"
				strHtml = strHtml &  "</tr>"
				firstconfirm=2
			end if
			strHtml = strHtml &  "<tr>"
			
		end if
		if(firstconfirm=0) then firstconfirm=1
		strHtml = strHtml & "<td align='middle' valign='top'><a href='Orderslist.asp?t=Orders&z_OrderId=%3D&x_OrderId=" & rs("OrderId") & "&z_CustomerId=%3D&x_CustomerId=&z_payment_status=LIKE&x_payment_status=&z_Ordered_Date=%3D&x_Ordered_Date=&psearch=&Submit=Search+(*)&psearchtype='>" & rs("OrderId") & "</a></td>"
		'strHtml = strHtml & "<td align='middle' valign='top'><a href='OrderDetailslist.asp?showmaster=Orders&OrderId=" & rs("OrderId") & "'>" & rs("OrderId") & "</a></td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("fullname") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("Ship_City") & "</td>"
		'strHtml = strHtml & "<td align='middle' valign='top'>" & rs("Ship_Province") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("payment_gross") & "</td>"
		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml
%></td>
     <td height="250" align="center" valign="top"><%
	strHtml=""

	sSql = "SELECT A.DiscountType, A.DiscountTypeId, A.CountOfDiscountid AS Used_cnt, b.CountOfDiscountid AS Tot_cnt "
	sSql=sSql & " FROM (SELECT DiscountTypes.DiscountType, DiscountTypes.DiscountTypeId, Discountcodes.used, Count(Discountcodes.Discountid) AS CountOfDiscountid, "
	sSql=sSql & " Orders.payment_status FROM (Orders RIGHT JOIN Discountcodes ON Orders.PromoCodeUsed = Discountcodes.DiscountCode) INNER JOIN DiscountTypes ON "
	sSql=sSql & " Discountcodes.DiscountTypeId = DiscountTypes.DiscountTypeId WHERE (((Orders.payment_status)=""Completed"")) GROUP BY DiscountTypes.DiscountType, "
	sSql=sSql & " DiscountTypes.DiscountTypeId, Discountcodes.used, Orders.payment_status ORDER BY DiscountTypes.DiscountTypeId, Discountcodes.used)  AS A LEFT JOIN "
	sSql=sSql & " (SELECT DiscountTypes.DiscountType, DiscountTypes.DiscountTypeId, Count(Discountcodes.Discountid) AS CountOfDiscountid FROM (Orders RIGHT JOIN Discountcodes "
	sSql=sSql & " 	ON Orders.PromoCodeUsed = Discountcodes.DiscountCode) INNER JOIN DiscountTypes ON Discountcodes.DiscountTypeId = DiscountTypes.DiscountTypeId WHERE  "
	sSql=sSql & " (((Orders.payment_status) Is Null)) GROUP BY DiscountTypes.DiscountType, DiscountTypes.DiscountTypeId ORDER BY DiscountTypes.DiscountTypeId)  AS b ON "
	sSql=sSql & " A.DiscountTypeId = b.DiscountTypeId;"

	
	'dim p(4,4)

	rs.open sSql,conn	
	'i=0
	
on error resume next
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4><h3>Promo Codes.. </h3></td>"
		strHtml = strHtml &  "</tr>"

		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='middle' valign='top'>Promo Codes</td>"
	
		strHtml = strHtml & "<td align='middle' valign='top'>Avl</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Used</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>% </td>"
		strHtml = strHtml &  "</tr>"
	Do While not rs.EOF
			t=rs("Tot_cnt")
			u=rs("Used_cnt")
			if(u="") then u=0
			if( not IsNumeric(t)) then t=0
			strHtml = strHtml &  "<tr>"
			strHtml = strHtml & "<td align='middle' valign='top'>" & rs("DiscountType") & "</td>"

			strHtml = strHtml & "<td align='middle' valign='top'>" & t & "</td>"
			strHtml = strHtml & "<td align='middle' valign='top'>" & u & "</td>"
			strHtml = strHtml & "<td align='middle' valign='top'>" & round((u/(u+t)*100),2) & "%</td>"
			
		strHtml = strHtml &  "</tr>"
			strHtml = strHtml &  "<tr><td bgcolor='#999' height=1px colspan=4> </td></tr>"
			rs.MoveNext		

	Loop
	rs.close
	

	
	strHtml = strHtml &  "</table>"

	Response.write strHtml	

%></td>
     <td rowspan="2" align="right" valign="top"><%	
	strHtml=""
	sSql = "SELECT top 8 Count(Orders.OrderId) AS CountOfOrderId, Customers.Inv_FirstName, Customers.Inv_LastName, Customers.CustomerID ,Sum(Orders.payment_gross) AS SumOfpayment_gross":
	sSql = sSql & " FROM Orders INNER JOIN Customers ON Orders.CustomerId = Customers.CustomerID WHERE (((Orders.payment_status)=""Completed""))"
	sSql = sSql & " GROUP BY Customers.Inv_FirstName, Customers.Inv_LastName, Customers.CustomerID ORDER BY Sum(Orders.payment_gross) DESC;"
	'sSql = sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) GROUP BY  datevalue(Orders.Ordered_Date) ORDER BY  datevalue(Orders.Ordered_Date) DESC;"

	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=3><h3>Top 8 Customers.. </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='left' valign='top'>Customer</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Orders</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Amount</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		strHtml = strHtml & "<td align='left' valign='top'>" & rs("Inv_FirstName") & " " & rs("Inv_LastName") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("CountOfOrderId") & "</td>"
		strHtml = strHtml & "<td align='right' valign='top'>" & FormatCurrency(rs("SumOfpayment_gross")) & "</td>"

		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml

%>
       <%	
	strHtml=""
	sSql = "SELECT top 8 Count(Orders.OrderId) AS CountOfOrderId, Customers.Inv_FirstName, Customers.Inv_LastName, Customers.CustomerID ,Sum(Orders.payment_gross) AS SumOfpayment_gross":
	sSql = sSql & " FROM Orders INNER JOIN Customers ON Orders.CustomerId = Customers.CustomerID WHERE (((Orders.payment_status)=""Completed""))"
	sSql = sSql & " GROUP BY Customers.Inv_FirstName, Customers.Inv_LastName, Customers.CustomerID having Count(Orders.OrderId)>1 ORDER BY Count(Orders.OrderId) desc ,Sum(Orders.payment_gross) DESC;"
	'sSql = sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) GROUP BY  datevalue(Orders.Ordered_Date) ORDER BY  datevalue(Orders.Ordered_Date) DESC;"

	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=3><h3>Repeat Customers.. </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='left' valign='top'>Customer</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Orders</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Amount</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		strHtml = strHtml & "<td align='left' valign='top'>" & rs("Inv_FirstName") & " " & rs("Inv_LastName") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("CountOfOrderId") & "</td>"
		strHtml = strHtml & "<td align='right' valign='top'>" & FormatCurrency(rs("SumOfpayment_gross")) & "</td>"

		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml

%></td>
   </tr>
   <tr>
    <td height="250" align="center" valign="top"><%	
	strHtml=""
	sSql = " SELECT TOP 5 Year([Orders].[Ordered_Date]) & '-' & Right('0' & Month([Orders].[Ordered_Date]),2) AS Expr1, Count(Orders.OrderId) AS CountOfOrderId, Sum(Orders.Amount) AS SumOfAmount"
	sSql = sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed""))"
	sSql = sSql & "  GROUP BY Year([Orders].[Ordered_Date]) & '-' & Right('0' & Month([Orders].[Ordered_Date]),2) ORDER BY 1 DESC;"
	'sSql = sSql & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) GROUP BY  datevalue(Orders.Ordered_Date) ORDER BY  datevalue(Orders.Ordered_Date) DESC;"
'response.write sSql
	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4><h3>Summary 5 Months.. </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='left' valign='top'>Month</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Orders</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Amount</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Avg</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		strHtml = strHtml & "<td align='left' valign='top'>" &  (rs("Expr1"))  & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("CountOfOrderId") & "</td>"
		strHtml = strHtml & "<td align='right' valign='top'>" & FormatCurrency(rs("SumOfAmount")) & "</td>"
		strHtml = strHtml & "<td align='right' valign='top'>" & FormatCurrency(rs("SumOfAmount")/rs("CountOfOrderId")) & "</td>"
		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml

%></td>
    </tr>
  <tr>
    <td height="250" valign="top"><%
	strHtml=""
	sSql = "SELECT  Orders.Ship_Province, Sum(Orders.payment_gross) AS SumOfpayment_gross, Count(Orders.OrderId) AS CountOfOrderId FROM Orders "
	sSql=sSql & " WHERE (((Orders.payment_status)='Completed')) GROUP BY Orders.Ship_Province; "

	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=3><h3>Sales by Province..   </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='middle' valign='top' >Province</td>"
	
		strHtml = strHtml & "<td align='middle' valign='top'>Count</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Total</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("Ship_Province") & "</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>" & rs("CountOfOrderId") & "</td>"
		strHtml = strHtml & "<td align='right' valign='top'>" & FormatCurrency(rs("SumOfpayment_gross")) & "</td>"

		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml	
%></td>
    <td height="250" align="center" valign="top"><%

	
	Response.write m_strHtml	

%></td>
    <td height="250" align="right" valign="top"><%
	strHtml=""
	sSql = "SELECT TOP 5 Customers.Inv_FirstName, Customers.Inv_LastName,inv_city, Customers.inv_Province,CustomerID "
	sSql=sSql & " FROM Customers ORDER BY Customers.CustomerID desc; "

	rs.open sSql,conn					
		strHtml = strHtml &  "<table class='cptable'>"
		strHtml = strHtml &  "<tr class='ewTableCaption'>"
		strHtml = strHtml & "<td align='middle' valign='middle' colspan=4><h3>Recent signups.. </h3></td>"
		strHtml = strHtml &  "</tr>"
		strHtml = strHtml &  "<tr class='ewTableHeader'>"
		strHtml = strHtml & "<td align='left' valign='top' >Name</td>"
	
		strHtml = strHtml & "<td align='middle' valign='top'>City</td>"
		strHtml = strHtml & "<td align='middle' valign='top'>Province</td>"
		strHtml = strHtml &  "</tr>"
		
	Do While not rs.EOF
		strHtml = strHtml &  "<tr>"
		
		strHtml = strHtml & "<td align='left' valign='top'><a href='Customersedit.asp?showdetail=&CustomerID= " & rs("CustomerID") & "'>" & rs("Inv_FirstName") & " " & rs("Inv_LastName") & "</a></td>"

		
		strHtml = strHtml & "<td align='middle' valign='top' >" & rs("inv_city") & "</td>"
		strHtml = strHtml & "<td align='r' valign='top' >" & rs("inv_Province") & "</td>"

		strHtml = strHtml &  "</tr>"
		rs.MoveNext						
	Loop
	strHtml = strHtml &  "</table>"
	rs.close
	Response.write strHtml	


%></td>
  </tr>
   <tr>
  <td height="250" align="right" valign="top">
</td>
  </tr>
  <tr>
    <td height="250" valign="top"></td></tr>
</table>
<%
conn.Close ' Close Connection
%>
<!-- Put your custom html here -->
<link href="css/vslpaypal.css" rel="stylesheet" type="text/css" />


  <!--#include file="footer.asp"-->
  <%

' Drop page object
Set custompage = Nothing
%>
  <%

' -----------------------------------------------------------------
' Page Class
'
Class ccustompage

	' Page ID
	Public Property Get PageID()
		PageID = "custompage"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "custompage"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
	End Property

	' Message
	Public Property Get Message()
		Message = Session(EW_SESSION_MESSAGE)
	End Property

	Public Property Let Message(v)
		Dim msg
		msg = Session(EW_SESSION_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_MESSAGE) = msg
	End Property

	Public Property Get FailureMessage()
		FailureMessage = Session(EW_SESSION_FAILURE_MESSAGE)
	End Property

	Public Property Let FailureMessage(v)
		Dim msg
		msg = Session(EW_SESSION_FAILURE_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_FAILURE_MESSAGE) = msg
	End Property

	Public Property Get SuccessMessage()
		SuccessMessage = Session(EW_SESSION_SUCCESS_MESSAGE)
	End Property

	Public Property Let SuccessMessage(v)
		Dim msg
		msg = Session(EW_SESSION_SUCCESS_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_SUCCESS_MESSAGE) = msg
	End Property

	' Show Message
	Public Sub ShowMessage()
		Dim sMessage
		sMessage = Message
		If sMessage <> "" Then Response.Write "<p class=""ewMessage"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		If sSuccessMessage <> "" Then Response.Write "<p class=""ewSuccessMessage"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		If sErrorMessage <> "" Then Response.Write "<p class=""ewErrorMessage"">" & sErrorMessage & "</p>"
		Session(EW_SESSION_FAILURE_MESSAGE) = "" ' Clear message in Session
	End Sub

	' -----------------------------------------------------------------
	'  Class initialize
	'  - init objects
	'  - open ADO connection
	'
	Private Sub Class_Initialize()
		If IsEmpty(StartTimer) Then StartTimer = Timer ' Init start time

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize user table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "custompage"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Init
	'  - called before page main
	'  - check Security
	'  - set up response header
	'  - call page load events
	'
	Sub Page_Init()
		Set Security = New cAdvancedSecurity

		' Uncomment codes below for security
		'
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then Call Page_Terminate("login.asp")
		' Global page loading event (in userfn7.asp)

		Call Page_Loading()
	End Sub

	' -----------------------------------------------------------------
	'  Class terminate
	'  - clean up page object
	'
	Private Sub Class_Terminate()
		Call Page_Terminate("")
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Terminate
	'  - called when exit page
	'  - clean up ADO connection and objects
	'  - if url specified, redirect to url
	'
	Sub Page_Terminate(url)

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set ObjForm = Nothing

		' Go to url if specified
		If sReDirectUrl <> "" Then
			If Response.Buffer Then Response.Clear
			Response.Redirect sReDirectUrl
		End If
	End Sub

	'
	'  Subroutine Page_Terminate (End)
	' ----------------------------------------
	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		'SuccessMessage = "Welcome " & CurrentUserName
		' Put your custom codes here

	End Sub
End Class

Function getFileString(txtFilename)
	Set fs=Server.CreateObject("Scripting.FileSystemObject")

	Set f=fs.OpenTextFile(Server.MapPath(txtFilename), 1)
	getFileString =(f.ReadAll)
	f.Close
	Set f=Nothing
	Set fs=Nothing
end function

Sub Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat,sRtn,aAtt)
	Dim objMail
Set objMail = Server.CreateObject("CDO.Message")
Set objConfig = Server.CreateObject("CDO.Configuration")


objConfig.Fields(cdoSendUsingMethod) = 2
objConfig.Fields(cdoSMTPServer)="smtp.1and1.com"
objConfig.Fields(cdoSMTPServerPort)=25
objConfig.Fields(cdoSMTPAuthenticate)=1
objConfig.Fields(cdoSendUserName) = "info@vsl3.ca" 'Enter YOUR E-MAIL ADDRESS here
objConfig.Fields(cdoSendPassword) = "web4colon" 'Enter the PASSWORD for your email address
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




Function StatsGraph(Height)
	Dim SQL , RSg

	'get last 30 records from statistics.
	SQL = "SELECT TOP 30 DateValue(Orders.Ordered_Date) AS OrderedDate, Sum(Orders.payment_gross) AS orders  "
	SQL = SQL & " FROM Orders WHERE (((Orders.payment_status)=""Completed"")) GROUP BY DateValue(Orders.Ordered_Date)"
	SQL = SQL & " ORDER BY DateValue(Orders.Ordered_Date) Asc;" & vbCrLf
	'Order them by first column (date) ascending
	'SQL = "select * from (" & SQL & ") As a order by 1 ASC"


	'get recordset from SQL
	Set RSg = conn.Execute(SQL)

	Dim HTML
	HTML = HTMLGraf(RSg, Height, 30) 
	StatsGraph = HTML
End Function 



'Simple HTML bar chart
'2003 Antonin Foller, http://www.motobit.com
'function accepts recordset with first column As X values
'and other columns As data series.
'The data series must be aligned To graph height.
Function HTMLGraf(RSg, height, SerieWidth)
  Dim Col, Row, cCol 
  Dim Colors, nCols
  nCols = RSg.Fields.Count - 1

  'Colors For series
  Colors = Array("Blue", "LightGreen", "Yellow", "Black", "Magenta", "Red")
  
  'Chart border table
  Dim aHTML, rowHTML
  aHTML = "<Table border=1 bordercolor=blue CellPadding=0 CellSpacing=0 height=" _
    & height & "><tr><td vAlign=Center>"

  'Series labels.
  For cCol = 1 To nCols
    Color = colors(cCol-1)
    aHTML = aHTML & "<Table width=100% border=0 CellPadding=0 CellSpacing=0 height=" _
      & SerieWidth & "><tr><td bgColor=" & Color & _
      " style=""color:white;font-size:8pt""> " & _
      RSg.Fields(cCol).Name & " </td></tr></table>"
  Next
  
  aHTML = aHTML & "</td></tr><tr><td vAlign=Center>"
  
  'Each row of recordset is one point.
  Do While Not rsg.eof
    rowHTML = "<Table Height=" & height & _
      " Cellpadding=0 CellSpacing=0 border=0 Align=Left><TR><TD>" 
    For cCol = 1 To nCols
      Dim Value, Color
      'get serie color
      Color = colors(cCol-1)

      on error resume Next
      'get value
      Value = (rsg(cCol))/15

      If isnull(Value) Or err > 0 Then 
        'empty Or problematic value
        rowHTML = rowHTML & "<Table Height=1 Cellpadding=0 CellSpacing=0 border=0 Width=" & _
          SerieWidth & " Align=Left><tr><td></td></tr></Table>" 
      Else
        'create one data point
        'the data point size is defined by td height
        rowHTML = rowHTML & "<Table Height=" & height 
        rowHTML = rowHTML & " Cellpadding=0 CellSpacing=0 border=0 Width="
        rowHTML = rowHTML & SerieWidth & " Align=Left><tr><td height=" 
        rowHTML = rowHTML & (100-Value) & "%" 
        rowHTML = rowHTML & "></td></tr><tr><td bgColor=" 
        rowHTML = rowHTML & Color  & " height=" & Value 
        rowHTML = rowHTML & "%" & "></td></tr></Table>"
      End If
    Next
    
    'X-labels.
    rowHTML = rowHTML & "</td></tr><tr><td Align=Center NoWrap>"
    rowHTML = rowHTML & "<DIV style=""FONT-SIZE: 10pt; WIDTH: 10pt; WRITING-MODE: tb-rl""> " 
    rowHTML = rowHTML & rsg(0) & "</DIV>"
    rowHTML = rowHTML & "</td></tr></table>" & vbCrLf
    aHTML = aHTML & rowHTML
    rsg.MoveNext
  Loop

  'Closing border table tag.
  aHTML = aHTML & "</td></tr></table>"
  HTMLGraf = aHTML
End Function

%>
</p>
