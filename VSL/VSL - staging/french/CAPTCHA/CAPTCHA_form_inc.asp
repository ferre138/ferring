<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz CAPTCHA(TM)
'**  http://www.webwizCAPTCHA.com
'**                                                              
'**  Copyright (C)2005-2008 Web Wiz(TM). All rights reserved.  
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM 'WEB WIZ'.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN 'WEB WIZ' IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwizguide.com/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwizguide.com
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************            



'*************************** SOFTWARE AND CODE MODIFICATIONS **************************** 
'**
'** MODIFICATION OF THE FREE EDITIONS OF THIS SOFTWARE IS A VIOLATION OF THE LICENSE  
'** AGREEMENT AND IS STRICTLY PROHIBITED
'**
'** If you wish to modify any part of this software a license must be purchased
'**
'****************************************************************************************






%>
<script language="javaScript">
function reloadCAPTCHA() {
	document.getElementById('CAPTCHA').src='CAPTCHA/CAPTCHA_image.asp?'+Date();
}
</script>           
<table width="100%" border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td><img src="CAPTCHA/CAPTCHA_image.asp" alt="Code Image - Please contact webmaster if you have problems seeing this image code" id="CAPTCHA" />&nbsp;<a href="javascript:reloadCAPTCHA();"><% = strTxtLoadNewCode %></a></td>
 </tr>
 <tr>
 <td>&nbsp;</td>
 </tr>
 <tr>
  <td>
   <input type="hidden" name="CAPTCHA_Postback" id="CAPTCHA_Postback" value="true" />
   <input type="text" class="form-control whitefield" name="securityCode" id="securityCode" size="14" maxlength="12" autocomplete="off" />
  </td>
 </tr>
</table>