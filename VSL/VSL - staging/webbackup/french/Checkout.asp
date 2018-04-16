<%@ Language=VBScript %>
<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->
<!--#include file="header.asp"--> 
              <script language="JavaScript" type="text/JavaScript">
function copyAddress()
{
 formCheckout.ship_FirstName.value=formCheckout.Inv_FirstName.value;
  formCheckout.ship_LastName.value=formCheckout.Inv_LastName.value;
   formCheckout.ship_Address.value=formCheckout.Inv_Address.value;
    formCheckout.ship_Address2.value=formCheckout.Inv_Address2.value;
	 formCheckout.ship_City.value=formCheckout.inv_City.value;
	 // formCheckout.ship_Province.value=formCheckout.inv_Province.value;
	 
	  formCheckout.ship_Province.selectedIndex=formCheckout.inv_Province.selectedIndex;
	   formCheckout.ship_PostalCode.value=formCheckout.inv_PostalCode.value;
	    formCheckout.ship_Country.value=formCheckout.inv_Country.value;
		
		formCheckout.HomePhone.value= formCheckout.inv_PhoneNumber.value;
}
</script> 
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' doit contenir une adresse d\'E-mail.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+'  requis.\n'; }
  }
  
   p=errors.indexOf('ship_FirstName');
  if(p>0)errors=errors.substring(0,p)+ ' Prénom' + errors.substring(p+('ship_FirstName').length);
  
   p=errors.indexOf('ship_LastName');
  if(p>0)errors=errors.substring(0,p)+ ' nom de famille' + errors.substring(p+('ship_LastName').length);
   
   p=errors.indexOf('ship_Address');
  if(p>0)errors=errors.substring(0,p)+ ' Adresse' + errors.substring(p+('ship_Address').length);
   
  p=errors.indexOf('ship_emailAddress');
  if(p>0)errors=errors.substring(0,p)+ ' courriel' + errors.substring(p+('ship_emailAddress').length);
  
  p=errors.indexOf('ship_telephone');
  if(p>0)errors=errors.substring(0,p)+ ' Numéro de téléphone' + errors.substring(p+('ship_telephone').length);
  
   p=errors.indexOf('ship_Reason');
  if(p>0)errors=errors.substring(0,p)+ ' Raison' + errors.substring(p+('ship_Reason').length);

   p=errors.indexOf('ship_City');
  if(p>0)errors=errors.substring(0,p)+ ' Ville' + errors.substring(p+('ship_City').length);
  
     p=errors.indexOf('ship_Country');
  if(p>0)errors=errors.substring(0,p)+ ' Pays' + errors.substring(p+('ship_Country').length);
  
  p=errors.indexOf('ship_Province');
  if(p>0)errors=errors.substring(0,p)+ ' Province' + errors.substring(p+('ship_Province').length);
  
    p=errors.indexOf('ship_PostalCode');
  if(p>0)errors=errors.substring(0,p)+ ' Code postal' + errors.substring(p+('ship_PostalCode').length);
  


  if (errors) alert('Les erreurs suivantes se sont produites:\n'+errors);
  document.MM_returnValue = (errors == '');
  

}

//-->
</script>
              <table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01"> 
                <tr> 
                  <td width="680" rowspan="2"><img src="images/title_shipping_fr.png" width="410" height="75"></td> 
                  <td width="28" valign="top"><img src="images/FontSize_fr.png" border="0" alt=""> </td> 
                  <td width="24" valign="top"> <a href="#"
				onmouseover="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseout="changeImages('login_13', 'images/font1.png'); return true;"
				onmousedown="changeImages('login_13', 'images/login_13-over.jpg'); return true;"
				onmouseup="changeImages('login_13', 'images/login_13-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('default'); 
return false;"> <img name="login_13" src="images/font1.png" width="24" height="27" border="0" alt=""></a></td> 
                  <td width="24"  valign="top"> <a href="#"
				onmouseover="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseout="changeImages('login_14', 'images/font2.png'); return true;"
				onmousedown="changeImages('login_14', 'images/login_14-over.jpg'); return true;"
				onmouseup="changeImages('login_14', 'images/login_14-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Medium'); 
return false;"> <img name="login_14" src="images/font2.png" width="24" height="27" border="0" alt=""></a></td> 
                  <td width="26"  valign="top"> <a href="#"
				onmouseover="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseout="changeImages('login_15', 'images/font3.png'); return true;"
				onmousedown="changeImages('login_15', 'images/login_15-over.jpg'); return true;"
				onmouseup="changeImages('login_15', 'images/login_15-over.jpg'); return true;" onClick="javascript:setActiveStyleSheet('Large'); 
return false;"><img name="login_15" src="images/font3.png" width="24" height="27" border="0" alt=""></a></td> 
                </tr>
                <tr>
                  <td colspan="4" valign="top"><div align="right">
                    <p class="bodycopy_small" style="color: #1070A3"><a href="../Checkout.asp" class="bodycopy_small">english  &gt;</a></p>
                  </div></td>
                </tr> 
              </table> 
              <div align="right"><span class="vslcss"><a href="VSLOrderForm.asp">Retour aux produits </a> : <a href="vslCart.asp">Voir le panier </a> : <a href="Customersedit.asp">Changement au compte </a> : <a href="changepwd.asp">Changement au mot de passe </a> : <a href="logout.asp">Quitter</a><img src="images/spacer.gif" width="65" height="10"> 
                <%
			
dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Request.ServerVariables("APPL_PHYSICAL_PATH") & "\db\vsldb.mdb" & ";"

Dim Security
Set Security = New cAdvancedSecurity
%> 
                <%
If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
	If Not Security.IsLoggedIn() Then
		Call Security.SaveLastUrl()
		Call Page_Terminate("login.asp")
	End If
'response.write " response" & request.Form("SubmitSecure")
if(request.Form("SubmitSecure")="Proceed to Secure Checkout") then
	preparedata(conn)
else
'response.end
	call getAddress (conn)
end if


function getAddress(c)
	dim rs,strSql
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSql= "SELECT QAddress.Inv_FirstName, QAddress.Inv_LastName, QAddress.Inv_Address, QAddress.Inv_Address2, "
	strSql= strSql & " QAddress.inv_City, QAddress.inv_Province, QAddress.inv_PostalCode, QAddress.inv_Country, "
	strSql= strSql & " QAddress.inv_PhoneNumber, QAddress.inv_EmailAddress, QAddress.inv_Fax, QAddress.ship_FirstName,"
	strSql= strSql & " 	QAddress.ship_LastName, QAddress.ship_Address, QAddress.ship_Address2, QAddress.ship_City, "
	strSql= strSql & " QAddress.ship_Province, QAddress.ship_PostalCode, QAddress.ship_Country, QAddress.HomePhone, "
	strSql= strSql & " QAddress.Customers.CustomerId FROM QAddress "
	strSql= strSql & " WHERE (((QAddress.Customers.CustomerId)=" & Security.CurrentUserID & "));"

	rs.Open strSql, c, 1, 2 
	if(not rs.eof) then %> 
                </span> </div> 
              <form  method="post" action="Checkout.asp" name="formCheckout" id="formCheckout" onSubmit="MM_validateForm('Inv_FirstName','','R','Inv_LastName','','R','Inv_Address','','R','inv_City','','R','inv_Province','','R','inv_PostalCode','','R','inv_Country','','R','inv_PhoneNumber','','R','inv_EmailAddress','','RisEmail','ship_FirstName','','R','ship_LastName','','R','ship_Address','','R','ship_City','','R','ship_Province','','R','ship_PostalCode','','R','ship_Country','','R');return document.MM_returnValue"> 
                <table width="720"  border="0" cellspacing="0" cellpadding="0"> 
                  <tr> 
                    <td><li class="t"> 
                        <table width="360"  border="0" cellpadding="0" cellspacing="0" class="ewTableNoBorder"> 
                          <tr> 
                            <td colspan="2"><strong><font size="+1">Adresse de facturation</font></strong></td> 
                          </tr> 
                          <tr> 
                            <td colspan="2" height="10px"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Prénom  </td> 
                            <td width="244"><input name="Inv_FirstName" type="text" id="Inv_FirstName" value="<%=rs("Inv_FirstName")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Nom </td> 
                            <td width="244"><input name="Inv_LastName" type="text" id="Inv_LastName" value="<%=rs("Inv_LastName")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116"> Adresse </td> 
                            <td width="244"><input name="Inv_Address" type="text" id="Inv_Address" value="<%=rs("Inv_Address")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">&nbsp;</td> 
                            <td width="244"><input name="Inv_Address2" type="text" id="Inv_Address2" value="<%=rs("Inv_Address2")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Ville</td> 
                            <td width="244"><input name="inv_City" type="text" id="inv_City" value="<%=rs("inv_City")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Province</td> 
                            <td width="244"> 
                              <select id='inv_Province' name='inv_Province'> 
                                <option value="" > Please Select </option> 
                                <option value="AL" <%if(rs("inv_Province")="AL") then response.Write "Selected"%>> Alberta </option> 
                                <option value="BC" <%if(rs("inv_Province")="BC") then response.Write "Selected"%>> Colombie britannique</option> 
                                <option value="MB" <%if(rs("inv_Province")="MB") then response.Write "Selected"%>> Manitoba </option> 
                                <option value="NB" <%if(rs("inv_Province")="NB") then response.Write "Selected"%>> Nouveau-Brunswick </option> 
                                <option value="NL" <%if(rs("inv_Province")="NL") then response.Write "Selected"%>> Terre-Neuve et Labrador </option> 
                                <option value="NT" <%if(rs("inv_Province")="NT") then response.Write "Selected"%>> (territoires du) Nord-Ouest </option> 
                                <option value="NS" <%if(rs("inv_Province")="NS") then response.Write "Selected"%>> Nouvelle-ƒcosse </option> 
                                <option value="NU" <%if(rs("inv_Province")="NU") then response.Write "Selected"%>> Nunavut </option> 
                                <option value="ON" <%if(rs("inv_Province")="ON") then response.Write "Selected"%>> Ontario </option> 
                                <option value="PE" <%if(rs("inv_Province")="PE") then response.Write "Selected"%>> l'”le du Prince-ƒdouard</option> 
                                <option value="QC" <%if(rs("inv_Province")="QC") then response.Write "Selected"%>> QuŽbec </option> 
                                <option value="SK" <%if(rs("inv_Province")="SK") then response.Write "Selected"%>> Saskatchewan </option> 
                                <option value="YT" <%if(rs("inv_Province")="YT") then response.Write "Selected"%>> (territoire du) Yukon </option> 
                              </select> </td> 
                          </tr> 
                          <tr> 
                            <td width="116">Code postal </td> 
                            <td width="244"><input name="inv_PostalCode" type="text" id="inv_PostalCode" value="<%=rs("inv_PostalCode")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Pays</td> 
                            <td width="244"><input name="inv_Country" type="text" id="inv_Country" value="<%=rs("inv_Country")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Numéro de téléphone                                                                          </td> 
                            <td width="244"><input name="inv_PhoneNumber" type="text" id="inv_PhoneNumber" value="<%=rs("inv_PhoneNumber")%>" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Adresse courriel </td> 
                            <td width="244"><input name="inv_EmailAddress" type="text" id="inv_EmailAddress" value="<%=rs("inv_EmailAddress")%>" size="20"></td> 
                          </tr> 
                        </table> 
                      </li></td> 
                    <td valign="top"><li class="t"> 
                        <table width="360"  border="0" cellpadding="0" cellspacing="0" class="ewTableNoBorder"> 
                          <tr> 
                            <td colspan="2"><strong><font size="+1">Adresse d’expédition</font></strong></td> 
                          </tr> 
                          <tr> 
                            <td colspan="2" height="10px"><input name="copyaddress" type="button" onClick="javascript:copyAddress();" value="adresse de facturation"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Prénom </td> 
                            <td width="237"><input name="ship_FirstName" type="text" id="ship_FirstName" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Nom</td> 
                            <td width="237"><input name="ship_LastName" type="text" id="ship_LastName" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123"> Adresse </td> 
                            <td width="237"><input name="ship_Address" type="text" id="ship_Address" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">&nbsp;</td> 
                            <td width="237"><input name="ship_Address2" type="text" id="ship_Address2" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Ville</td> 
                            <td width="237"><input name="ship_City" type="text" id="ship_City" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Province</td> 
                            <td width="237">						      <select id='ship_Province' name='ship_Province'> 
                                  <option value="" selected>--Aucun--</option>
                                    <option value="AB">Alberta</option>
                              <option value="BC">Colombie britannique</option>
                              <option value="MB">Manitoba</option>
                              <option value="NB">Nouveau-Brunswick</option>
                              <option value="NL">Terre-Neuve et Labrador</option>

                              <option value="NT">(territoires du) Nord-Ouest</option>
                              <option value="NS">Nouvelle-ƒcosse</option>
                              <option value="NU">Nunavut</option>
                              <option value="ON">Ontario</option>
                              <option value="PE">l'”le du Prince-ƒdouard</option>
                              <option value="QC">QuŽbec</option>

                              <option value="SK">Saskatchewan</option>
                              <option value="YT">(territoire du) Yukon</option>
                              </select></td></tr> 
                          <tr> 
                            <td width="123">Code postal </td> 
                            <td width="237"><input name="ship_PostalCode" type="text" id="ship_PostalCode" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Pays</td> 
                            <td width="237"><input name="ship_Country" type="text" id="ship_Country"  size="20">                            </td> 
                          </tr> 
                          <tr> 
                            <td width="123">Numéro de téléphone</td> 
                            <td width="237"><input name="HomePhone" type="text" id="HomePhone" size="20"></td> 
                          </tr> 
                          <tr> 
                            <td>&nbsp;</td> 
                            <td>&nbsp;</td> 
                          </tr> 
                        </table> 
                      </li></td> 
                  </tr> 
                  <tr> 
                      
                    <td colspan="2" class="vslcss"><div class="t">
					
					
						<%if(Month(now()) & year(now())="102009") then %>
					<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >Veuillez commander le nombre de boîtes pour lesquelles vous voulez être facturés, et vous recevrez votre confirmation par courriel qui vous donnera le nombre de boîtes gratuites qui seront expédiées avec votre commande.</font></p>
					<%end if%>
					<p> <font color="#FF0000" size="+1">Les coûts d'envoi pour les commandes seront calculés lorsque vous recevrez votre confirmation de commande.</font></p>
                      <p><font size="2">Vous recevrez une confirmation et vous serez avisé de la date de livraison de votre commande par courriel ou par téléphone. </font></p>
                      <p><font size="2">Les commandes reçues avant 13 H, du lundi au jeudi (sauf les jours fériés) seront livrées le jour suivant. Les commandes reçues 
                             après 13 H seront livrées en deux jours ouvrables. Les commandes reçues le vendredi seront expédiées le jour ouvrable suivant.
    
                        </font></p>
                      <p><font size="2">La commande sera expédiée dans un emballage isolé contenant des blocs réfrigérants pour maintenir la température
                              d’entreposage au niveau exigé et elle sera livrée par un service de messagerie.
      
                      </font></p>
                      <p><strong> Veuillez prendre note qu&rsquo;une personne devra se trouver &agrave; votre adresse le jour de la livraison pour recevoir le paquet et s&rsquo;assurer qu&rsquo;il reste r&eacute;frig&eacute;r&eacute;.</strong> </p>
                      <font size="2"></font></div></td> 
                  </tr> 
                  <tr> 
                    <td colspan="2"><% call DisplayItems() %></td> 
                  </tr> 
                  <tr> 
                    <td colspan="2"><p align="center"><a href="vslCart.asp"><img src="images/clicktoreturn_fr.gif" width="201" height="32" border="0"></a>&nbsp; 
                        <input type="hidden" name="SubmitSecure" id="SubmitSecure" value="Proceed to Secure Checkout">
                        <input name="Submit" type="image" class="InputNoBorder" id="Submit"  value="Proceed to Secure Checkout" src="images/finalizeorder_fr.gif" width="200" height="32">
                    &nbsp; <a href="logout.asp"><img src="images/logout_fr.gif" width="159" height="32" border="0"></a> </p></td> 
                  </tr> 
                </table> 
              </form> 
              <script type="text/javascript">
    var myMenu = new ImageMenu($$('#kwick .kwick'),{openWidth:261,start:4});
  </script> 
              <%
	else
		response.write "Error.."
	end if
rs.close
set rs=nothing

end function%> 
              <!--#include file="footer.asp"-->
<%Call Page_Terminate("")

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


Sub DisplayItems()
	dim arr
	arr=GetCart()
	if( IsCartEmpty()) then  
	response.redirect "vslCart.asp?cmd=resetall"
	
		'message "emptycart"
		'response.redirect "vslCart.asp?cmd=resetall"
	else
dim applyDiscount
applyDiscount= checkCustomer(conn)
		response.Write "<div class='t' width=720px><b>Panier d’achat:</b>"
%>
<table width="760px">
  <tr>
    <td><table width="720px"  class="ewTable">
        <tr >
          <td width="50%" class="ewTableHeader"><b>Produits</b></td>
          <td width="16%" class="ewTableHeader"><div align="right"><b>Prix unitaire</b></div></td>
          <td width="12%" class="ewTableHeader"><div align="center"><b>Quantité</b></div></td>
          <td width="10%" class="ewTableHeader"><div align="right"><b>Total</b></div></td>
        </tr>
        <%For i=0 To UBound(arr)
		
		p=arr(i,2)
			t=arr(i,2)
				if(applyDiscount) then
					
					p=62.50
					
					t= "<p>Prix régulier : <s>" & arr(i,2) & "<br>"&vbCrLf 
					t=t & "</s><font color=""#FF0000"" size=""-1"">Prix spécial pour cette commande seulement</font>"
					t=t & ": <font color=""#FF0000""><strong>" & "62.50"  &"</strong></font> </p>"&vbCrLf
					
				end if
		%>
        <tr >
          <td width="50%"><b><%= "</b>" & arr(i,1)%></b></td>
          <td width="16%"><div align="right"><b><%=t%></b></div></td>
          <td width="12%" align="center"><div align="center">
              <%
			  if(arr(i,3)>19) then 
				Response.write arr(i,3) + 2
			else
				Response.write arr(i,3)
			end if
			  %>
            </div></td>
          <td width="10%"><div align="right"><%=arr(i,2) *  arr(i,3)%></div></td>
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
<%if(applyDiscount) then response.write " Cette commande sera expédiée gratuitement"%>
<%
response.write "</div>"
end if
End Sub


Sub PrepareData(c)
	dim arr
	arr=Session("cart")
	if(  IsCartEmpty()) then  response.redirect "vslCart.asp"
	if(uBound(arr)=-1) then
		'message "emptycart"
		response.redirect "vslCart.asp?cmd=resetall"
	else

		'response.Write "<b>Your Cart:</b>"
%>
<FORM METHOD="POST" name="MyForm" ACTION="http://ww11.aitsafe.com/cf/pay.cfm">
  <input type="hidden" name="userid" value="D0238311">
  <input type="hidden" name="return" value="http://www.vsl.ca/test/cart.asp">
  <%
				
				inv_name =request.form("Inv_FirstName") & " " & request.form("Inv_LastName") 
inv_company =request.form("inv_company")
inv_addr1 =request.form("Inv_Address")
inv_addr2 =request.form("Inv_city")
inv_state =request.form("inv_Province")
inv_zip =request.form("inv_PostalCode")
inv_country =request.form("inv_Country")
tel =request.form("inv_PhoneNumber")
fax =request.form("inv_Fax")
email =request.form("inv_EmailAddress")
del_name =request.form("del_FirstName") & " " & request.form("del_LastName") 
del_company =request.form("del_company")
del_addr1 =request.form("ship_Address")
del_addr2 =request.form("ship_city")
del_state =request.form("ship_Province")
del_zip =request.form("ship_PostalCode")
del_country =request.form("ship_Country")
del_tel =request.form("HomePhone")

gt=0
dim applyDiscount
applyDiscount= checkCustomer(c)

				dim shipProv,totalCnt
				shipProv= request.Form("ship_Province")
				dim rs,strSql,taxrate,ship1,ship2
				Set rs = Server.CreateObject("ADODB.Recordset")
				
				strSql= " SELECT TaxRate, ShipRate_first, ShipRate_Rest  FROM Province WHERE (Province.Prov='"  & shipProv &"');"
				rs.open strSql,c
				if(not rs.eof) then 
					taxrate=cdbl(rs("TaxRate"))
					ship1=cdbl(rs("ShipRate_first"))
					ship2=cdbl(rs("ShipRate_Rest"))
				end if
				rs.close
				if((lcase(replace(del_zip," ", "")) ="m2j5c1") ) then
					if((lcase(mid( replace(del_addr1," ",""),1,11))="200yorkland") and lcase(del_state)="on") then
						ship1=0
						ship2=0
					end if
				end if
				'response.write (lcase(replace(del_zip," ", ""))) & ship1 & ship2
				'response.end
				totalCnt=0
				k=1
				For i=0 To UBound(arr)
				strSql= "SELECT ItemId, fDescription, Price FROM Products WHERE (ItemId= "  & arr(i,0) & ");"
				rs.open strSql,c
				
				if(not rs.eof) then 
				p=getPrice(rs("ItemId"),c,arr(i,3))
				
				if(applyDiscount) then	p=62.50
				%>
  
  <input type="hidden" name="product<%=(i+1)%>" value="<%=  rs("fDescription")%><%if(applyDiscount) then response.write " ..Cette commande sera expédiée gratuitement"%>">
  <input type="hidden" name="price<%=(i+1)%>" value="<%=p%>">

  <input type="hidden" name="qty<%=(i+1)%>" size="3" value="<%=arr(i,3)%>">
  <%
  gt=gt + cdbl(p) * cdbl(arr(i,3))
  if(arr(i,3)>19) then 
  k=k+1%>
  	<input type="hidden" name="product<%=UBound(arr)+k%>" value="<%="Gratuit.." & rs("fDescription")%>">
  	<input type="hidden" name="price<%=UBound(arr)+k%>" value="0">
  	<input type="hidden" name="tax<%=UBound(arr)+k%>" value="0">
  	<input type="hidden" name="qty<%=UBound(arr)+k%>" size="3" value="2">
  <%end if%>
  <% totalCnt= totalcnt + cdbl(arr(i,3))
				totalTax= totalTax + round(rs("Price")*(taxrate)/100,2)
				
				end if
				rs.close
				Next%>
  <%if(totalCnt>9) then
	 ' ship1=ship2
	  ship1=0
	end if
	ship1=0
	if(del_state="NT" or del_state="YT" or del_state="NU") then ship1=0
	%>
  <input type="hidden" name="shipping" value="<%=ship1%>">
  <%
taxrate= cdbl(taxrate) * (1 + cdbl(ship1)/cdbl(gt))
				
For i=0 To UBound(arr)
%>
 <input type="hidden" name="tax<%=(i+1)%>" value="<%=cdbl((taxrate))%>">
<%
next
				%>

  <input type="hidden" name="lg" value="8">
  <input type="hidden" name="totaltax" value="<%=totalTax%>">
  <input type="hidden" name="inv_name" value="<%=inv_name & " " & Inv_FirstName %>">
  
  <input type="hidden" name="inv_company" value="">
  
  <input type="hidden" name="inv_addr1" value="<%=inv_addr1%>">
  
  <input type="hidden" name="inv_addr2" value="<%=inv_addr2%>">
  
  <input type="hidden" name="inv_state" value="<%=inv_state%>">
  
  <input type="hidden" name="inv_zip" value="<%=inv_zip%>">
  
  <input type="hidden" name="inv_country" value="<%=inv_country%>">
  
  <input type="hidden" name="tel" value="<%=tel%>">
  
  <input type="hidden" name="email" value="<%=email%>">
  
  <input type="hidden" name="del_name" value="<%=del_name & " " & del_FirstName %>">
  
  <input type="hidden" name="del_company" value="">
  
  <input type="hidden" name="del_addr1" value="<%=del_addr1%>">
  
  <input type="hidden" name="del_addr2" value="<%=del_addr2%>">
  
  <input type="hidden" name="del_state" value="<%=del_state%>">
  
  <input type="hidden" name="del_zip" value="<%=del_zip%>">
  
  <input type="hidden" name="del_country" value="<%=del_country%>">
  
  <input type="hidden" name="del_tel" value="<%=del_tel%>">
   <div class="t" align="center" valign="middle">
  <table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="middle" align="center">Redirection .. Attendez S’il vous plaît ..</td><td valign="middle" align="center"><img src="images/loading.gif"></td>
  </tr>
</table>

  
  </div>
  
</form>
<%
c.execute  "update Customers set NewCustomer =false WHERE (((Customers.CustomerID)=" & Security.CurrentUserID & "));" 
%>
<%'response.end%>
<script type="text/javascript" language="JavaScript"><!--
document.MyForm.submit();
//--></script>
<%

end if
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

function checkCustomer(c)
checkCustomer=false
	' dim rs,NewCustomer
	' Set rs = Server.CreateObject("ADODB.Recordset")
	' NewCustomer=false
	' rs.Open  "SELECT NewCustomer,Inv_FirstName,Inv_LastName,inv_PhoneNumber FROM Customers WHERE (((Customers.CustomerID)=" & Security.CurrentUserID & "));" , c, 1, 2 
	' if(not rs.eof) then NewCustomer = (rs("NewCustomer"))
	' if(NewCustomer) then 
		' Inv_FirstName=""
		' Inv_LastName=""
		' inv_PhoneNumber=left(cleantxt(rs("inv_PhoneNumber")),10)
		' rs.close	
		' rs.Open  "SELECT phonecode FROM phone WHERE phonecode='" & Inv_LastName   &  Inv_FirstName & inv_PhoneNumber & "';" , c, 1, 2 
		' if(not rs.eof) then NewCustomer = false
	' else
		' NewCustomer=false
	' end if
	' checkCustomer=NewCustomer
	'response.write NewCustomer
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
