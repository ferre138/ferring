<form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post" name="frmPayPal">
	<input type="hidden" name="cmd" value="_cart">
	<input type="hidden" name="upload" value="1">
	<!-- <input type="hidden" name="add" value="1">
	<input type="hidden" name="display" value="1">
	<input type="hidden" name="redirect_cmd" value="_xclick"> -->
	<input type="hidden" name="business" value="rtran_1302632991_biz@ravenshoegroup.com" />
	<input type="hidden" name="address_override" value="1">
	<input type="hidden" name="first_name" value="Ramy">
	<input type="hidden" name="last_name" value="Testing">
	<input type="hidden" name="email" value="rtran_1302632967_per@ravenshoegroup.com">
	
	<input type="hidden" name="item_name_1" value="Item Name 1">
	<INPUT TYPE="hidden" name="quantity_1" value="6">
	<input type="hidden" name="amount_1" value="1.00">
	<input type="hidden" name="shipping_1" value="1.00">
	<input type="hidden" name="tax_1" value="1.00">
	<INPUT TYPE="hidden" name="weight_1" value="1.5">
	
	<input type="hidden" name="item_name_2" value="Item Name 2">
	<input type="hidden" name="amount_2" value="2.00">	
	<input type="hidden" name="shipping_2" value="1.00">
	<input type="hidden" name="tax_2" value="1.00">
	<INPUT TYPE="hidden" name="weight_2" value="1.5">
	<INPUT TYPE="hidden" name="quantity_2" value="6">
	<input type="hidden" name="currency_code" value="CAD">
	
	<input type="hidden" name="custom" value="55five">
	<input type="hidden" name="invoice" value=5995> 
	<input type="hidden" name="return" value="http://www.ravenshoegroup.ca/VSLPayPal/Thank-you.asp?token=55">
	<input type="hidden" name="cancel_return" value="http://www.ravenshoegroup.ca/VSLPayPal/Cancel_order.asp?token=55">
	<input type="submit" name="btnSubmit" value="ShoppingCart">	
</form>

<script type="text/javascript" language="JavaScript">
<!--
// document.frmPayPal.submit();
//-->
</script>