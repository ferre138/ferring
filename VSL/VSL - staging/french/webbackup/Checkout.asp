<%@ Language=VBScript %>
<!--#include file="vslConfig.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="cartinc.asp"-->
<!--#include file="header.asp"--> 
<!--#include file="userfn60.asp"-->
              <script language="JavaScript" type="text/JavaScript">
function copyAddress()
{
	//if(!formCheckout.CheckLocalPickup.checked)
	//{

        var unit = formCheckout.Inv_Address2.value;
        var arr = [' buzzer',' buzz',' Buzzer',' Buzz'];

        if(unit == ''){        
            //DO NOTHING  
        }else{
            for(i = 0; i < arr.length; i++) { 
                var found = unit.includes(arr[i]);
                console.log(found);
                if(found){            
                    //BUZZER FOUND
                    //console.log('Found');
                    var unit_array = unit.split(arr[i]);
                    formCheckout.ship_Address2.value=unit_array[0];
                    formCheckout.Buzzer.value=(unit_array[1]).replace(":","");  
                    break;                  
                }else{
                    //BUZZER NOT FOUND
                    //console.log('Not Found');
                    formCheckout.ship_Address2.value=formCheckout.Inv_Address2.value;
                }
            }
        }

		formCheckout.ship_FirstName.value=formCheckout.Inv_FirstName.value;
		formCheckout.ship_LastName.value=formCheckout.Inv_LastName.value;
		formCheckout.ship_Address.value=formCheckout.Inv_Address.value;		
		formCheckout.ship_City.value=formCheckout.inv_City.value;
		// formCheckout.ship_Province.value=formCheckout.inv_Province.value;
	 
		formCheckout.ship_Province.selectedIndex=formCheckout.inv_Province.selectedIndex;
		formCheckout.ship_PostalCode.value=formCheckout.inv_PostalCode.value;
	    formCheckout.ship_Country.value=formCheckout.inv_Country.value;
		
		formCheckout.HomePhone.value= formCheckout.inv_PhoneNumber.value;
	//}
}
function localAddress()
{
	/*if(formCheckout.CheckLocalPickup.checked)
	{
		formCheckout.ship_FirstName.value=formCheckout.Inv_FirstName.value;
		formCheckout.ship_FirstName.disabled=true ;
		formCheckout.ship_LastName.value=formCheckout.Inv_LastName.value;
		formCheckout.ship_LastName.disabled=true ;
		formCheckout.ship_Address.value="200 YorkLand";
		formCheckout.ship_Address.disabled=true ;
		formCheckout.ship_Address2.value="Suite:500";
		formCheckout.ship_Address2.disabled=true ;
		formCheckout.ship_City.value="Toronto";
		formCheckout.ship_City.disabled=true ;
	 // formCheckout.ship_Province.value=formCheckout.inv_Province.value;
	 
		formCheckout.ship_Province.selectedIndex=9;
		formCheckout.ship_Province.disabled=true ;
		formCheckout.ship_PostalCode.value="M2J5C1";
		formCheckout.ship_PostalCode.disabled=true ;
	    formCheckout.ship_Country.value="Canada";
		formCheckout.ship_Country.disabled=true ;
		
		formCheckout.HomePhone.value= "416 642-0075";
		formCheckout.HomePhone.disabled=true ;
		
		
	}
	else
	{
		formCheckout.ship_FirstName.disabled=false ;
		formCheckout.ship_LastName.disabled=false ;
		formCheckout.ship_Address.disabled=false ;
		formCheckout.ship_Address2.disabled=false ;
		formCheckout.ship_City.disabled=false ;
		formCheckout.ship_Province.disabled=false ;
		formCheckout.ship_PostalCode.disabled=false ;
		formCheckout.ship_Country.disabled=false ;
		formCheckout.HomePhone.disabled=false ;
	}
	*/
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
    var currentBuzzer = $('#Buzzer').val();
    // var add2 = $('#ship_Address2').val();
    // var newStr = add2 + ' Buzz:' + currentBuzzer;
    // $('#ship_Address2').val(newStr);

    
    var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;

    if(currentBuzzer === '' || currentBuzzer === null){
        errors += '- Numéro de sonnerie est requis.\n';
    }else{
        var add2 = $('#ship_Address2').val();
        var newStr = add2 + ' Buzz:' + currentBuzzer;
        $('#ship_Address2').val(newStr);
    }

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
function PaypalSubmit(){
MM_validateForm('Inv_FirstName','','R','Inv_LastName','','R','Inv_Address','','R','inv_City','','R','inv_Province','','R','inv_PostalCode','','R','inv_Country','','R','inv_PhoneNumber','','R','inv_EmailAddress','','RisEmail','ship_FirstName','','R','ship_LastName','','R','ship_Address','','R','ship_City','','R','ship_Province','','R','ship_PostalCode','','R','ship_Country','','R');

if(document.MM_returnValue) {
document.formCheckout.action="paypal.asp";
document.formCheckout.onSubmit="";
document.formCheckout.submit();
}
return false;
	}


function checkAddress(){
    var currentBuzzer = $('#Buzzer').val();
    if(needBuzzer() && currentBuzzer == ""){
       // $('#Buzzer').val('N/A');
        alert('Il semble que vous pourriez avoir besoin d\'inclure un numéro de bruiteur. Si non, veuillez indiquer N/A');
    }else{
        //Ignore
    }
}


function needBuzzer(){
    var str = $('#ship_Address').val();
    var str2 = $('#ship_Address2').val();
    var res = str.toLowerCase();
    var res2 = str2.toLowerCase();
    var arr = ['apartment','appartement','suite','bureau','unit','unite','unitÃ©'];
    if(res2 == ''){        
        for(i = 0; i < arr.length; i++) { 
            var found = res.includes(arr[i]);
            if(found){
                return found;
            }
        }  
    }else{
        for(i = 0; i < arr.length; i++) { 
            var found = res2.includes(arr[i]);
            console.log(found);
            if(found){            
                return found;
            }
        }
    $('#ship_Address2').val('Unit '+ res2);
    return true;    
    }      
}



//-->
</script>
              <table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01"> 
                <tr> 
                  <td width="680" rowspan="2"><img src="images/title_shipping_fr.png" width="410" height="75"></td> 
                 
                </tr>
                <tr>
                  <td colspan="4" valign="top"><div align="right">
                    <p class="bodycopy_small" style="color: #1070A3"><a href="../Checkout.asp" class="bodycopy_small">english  &gt;</a></p>
                  </div></td>
                </tr> 
              </table> 
              <div align="right"><span class="vslcss"><a href="VSLOrderForm.asp">Retour aux produits </a> : <a href="vslCart.asp">Voir le panier </a> : <a href="Customersedit.asp">Changement au compte </a> : <a href="changepwd.asp">Changement au mot de passe </a> : <a href="logout.asp">Quitter</a><img src="images/spacer.gif" width="65" height="10"> 
                <%

Dim promomsg			
dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING


conn.execute " UPDATE Orders SET Orders.PromoCodeUsed = '' WHERE (((Orders.Ordered_Date)<DateAdd('h',-24,Now())) AND ((Orders.payment_status)='WIP')) ;"
'conn.execute " delete * from orderdetails where orderid in( select orderid FROM Orders WHERE (((Orders.Ordered_Date)<DateAdd('h',-168,Now())) AND ((Orders.payment_status)='WIP')) OR (((Orders.payment_status)='Cancelled')));"
'conn.execute "delete * FROM Orders WHERE (((Orders.Ordered_Date)<DateAdd('h',-168,Now())) AND ((Orders.payment_status)='WIP')) OR (((Orders.payment_status)='Cancelled')); "

conn.execute " delete * from orderdetails where orderid in( select orderid FROM Orders WHERE  (((Orders.payment_status)='Cancelled')));"
conn.execute "delete * FROM Orders WHERE  (((Orders.payment_status)='Cancelled')); "


Dim Security
Set Security = New cAdvancedSecurity

%> 
                <%
If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
	If Not Security.IsLoggedIn() Then
		Call Security.SaveLastUrl()
		Call Page_Terminate("login.asp")
	End If

	shipvalues =request.form("shipvalues")
	if(shipvalues<>"") then
		shipvalue= split(shipvalues,"||")
		ship1=shipvalue(0)
		ship2=shipvalue(1)
		ship3=shipvalue(2)
		ship4=shipvalue(3)
		ship5=shipvalue(4)
		ship6=shipvalue(5)
		ship7=shipvalue(6)
		ship8=shipvalue(7)
		ship9=shipvalue(8)
	else
		ship1=""
		ship2=""
		ship3=""
		ship4=""
		ship5=""
		ship6=""
		ship7=""
		ship8=""
		ship9=""
		
	end if

	call getAddress (conn)

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
        <form  method="post" action="paypal.asp" name="formCheckout" id="formCheckout" onSubmit="MM_validateForm('Inv_FirstName','','R','Inv_LastName','','R','Inv_Address','','R','inv_City','','R','inv_Province','','R','inv_PostalCode','','R','inv_Country','','R','inv_PhoneNumber','','R','inv_EmailAddress','','RisEmail','ship_FirstName','','R','ship_LastName','','R','ship_Address','','R','ship_City','','R','ship_Province','','R','ship_PostalCode','','R','ship_Country','','R');return document.MM_returnValue"> 
                <table width="720"  border="0" cellspacing="0" cellpadding="0"> 
                  <tr> 
                    <td><li class="t"> 
                        <table width="360"  border="0" cellpadding="0" cellspacing="0" class="ewTableNoBorder"> 
                          <tr> 
            <td colspan="2"><strong><font size="+1">Adresse de facturation </font></strong> <a href="Customersedit.asp">(éditer)</a>
              <input type="hidden" name="SubmitSecure" id="SubmitSecure" value="Proceed to Secure Checkout">
            </td> 
                          </tr> 
                          <tr> 
                            <td colspan="2" height="10px"></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Prénom  </td> 
                            <td width="244"><input name="Inv_FirstName" type="text" id="Inv_FirstName" value="<%=rs("Inv_FirstName")%>" size="20" readonly></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Nom </td> 
                            <td width="244"><input name="Inv_LastName" type="text" id="Inv_LastName" value="<%=rs("Inv_LastName")%>" size="20" readonly></td> 
                          </tr> 
                          <tr> 
                            <td width="116"> Adresse </td> 
                            <td width="244"><input name="Inv_Address" type="text" id="Inv_Address" value="<%=rs("Inv_Address")%>" size="20" disabled></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Suite</td> 
                            <td width="244"><input name="Inv_Address2" type="text" id="Inv_Address2" value="<%=rs("Inv_Address2")%>" size="20" disabled ></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Ville</td> 
                            <td width="244"><input name="inv_City" type="text" id="inv_City" value="<%=rs("inv_City")%>" size="20" disabled></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Province</td> 
                            <td width="244"> 
                              <select id='inv_Province' name='inv_Province' disabled> 
                                <option value="" > Please Select </option> 
                                <option value="AB" <%if(rs("inv_Province")="AB") then response.Write "Selected"%>> Alberta </option> 
                                <option value="BC" <%if(rs("inv_Province")="BC") then response.Write "Selected"%>> Colombie-Britannique</option> 
                                <option value="MB" <%if(rs("inv_Province")="MB") then response.Write "Selected"%>> Manitoba </option> 
                                <option value="NB" <%if(rs("inv_Province")="NB") then response.Write "Selected"%>> Nouveau-Brunswick </option> 
                                <option value="NL" <%if(rs("inv_Province")="NL") then response.Write "Selected"%>> Terre-Neuve et Labrador </option> 
                                <option value="NT" <%if(rs("inv_Province")="NT") then response.Write "Selected"%>> (territoires du) Nord-Ouest </option> 
                                <option value="NS" <%if(rs("inv_Province")="NS") then response.Write "Selected"%>> Nouvelle-Écosse </option> 
                                <option value="NU" <%if(rs("inv_Province")="NU") then response.Write "Selected"%>> Nunavut </option> 
                                <option value="ON" <%if(rs("inv_Province")="ON") then response.Write "Selected"%>> Ontario </option> 
                                <option value="PE" <%if(rs("inv_Province")="PE") then response.Write "Selected"%>> l'île du Prince-Édouard</option> 
                                <option value="QC" <%if(rs("inv_Province")="QC") then response.Write "Selected"%>> Québec </option> 
                                <option value="SK" <%if(rs("inv_Province")="SK") then response.Write "Selected"%>> Saskatchewan </option> 
                                <option value="YT" <%if(rs("inv_Province")="YT") then response.Write "Selected"%>> (territoire du) Yukon </option> 
                              </select> </td> 
                          </tr> 
                          <tr> 
                            <td width="116">Code postal </td> 
                            <td width="244"><input name="inv_PostalCode" type="text" id="inv_PostalCode" value="<%=rs("inv_PostalCode")%>" size="20" disabled></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Pays</td> 
                            <td width="244"><input name="inv_Country" type="text" id="inv_Country" value="<%=rs("inv_Country")%>" size="20" disabled></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Numéro de téléphone                                                                          </td> 
                            <td width="244"><input name="inv_PhoneNumber" type="text" id="inv_PhoneNumber" value="<%=rs("inv_PhoneNumber")%>" size="20" disabled></td> 
                          </tr> 
                          <tr> 
                            <td width="116">Adresse courriel </td> 
                            <td width="244"><input name="inv_EmailAddress" type="text" id="inv_EmailAddress" value="<%=rs("inv_EmailAddress")%>" size="20" disabled></td> 
                          </tr> 
                        </table> 
                      </li></td> 
                    <td valign="top"><li class="t"> 
                        <table width="360"  border="0" cellpadding="0" cellspacing="0" class="ewTableNoBorder"> 
                          <tr> 
                            <td colspan="2"><strong><font size="+1">Adresse d’expédition</font></strong></td> 
                          </tr> 
                          <tr> 
                <td colspan="2" height="10px"><table align="left" border="0" cellspacing="0" cellpadding="0" width="100%">

	<tr align="left" valign="top">
		<td width="50%"><input name="copyaddress" type="button" onClick="javascript:copyAddress();" value="adresse de facturation"></td>
		<td width="50%"><!-- <input type="checkbox" name="CheckLocalPickup" id="CheckLocalPickup" onClick="javascript:localAddress();"><span class="aspmaker">Recueil local à Toronto seulement</span>--></td>
	</tr>
</table>
                           </td> 
                          </tr> 
                          <tr> 
                            <td width="123">Prénom </td> 
                            <td width="237"><input name="ship_FirstName" type="text" id="ship_FirstName" size="20" value="<%=ship1%>"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Nom</td> 
                            <td width="237"><input name="ship_LastName" type="text" id="ship_LastName" size="20" value="<%=ship2%>"></td> 
                          </tr> 
                          <tr> 
                            <td width="123"> Adresse </td> 
                            <td width="237"><input name="ship_Address" type="text" id="ship_Address" size="20" value="<%=ship3%>"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Numéro de l'unité</td> 
                            <td width="237"><input name="ship_Address2" type="text" id="ship_Address2" onblur="checkAddress()" size="20" value="<%=ship4%>"></td> 
                          </tr> 
                          <tr> 
                                 <td width="116">Buzzer**</td> 
                          <td width="244"><input name="Buzzer" type="text" id="Buzzer"  onblur="checkAddress()"  value="" size="20"></td> 
                          </tr>
                            <tr> 
                            <td width="123">Ville</td> 
                            <td width="237"><input name="ship_City" type="text" id="ship_City" size="20" value="<%=ship5%>"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Province</td> 
                            <td width="237">						      <select id='ship_Province' name='ship_Province'> 
                                  <option value="" selected>--Aucun--</option>
                                    <option value="AB">Alberta</option>
                              <option value="BC">Colombie-Britannique</option>
                              <option value="MB">Manitoba</option>
                              <option value="NB">Nouveau-Brunswick</option>
                              <option value="NL">Terre-Neuve et Labrador</option>

                              <option value="NT">(territoires du) Nord-Ouest</option>
                              <option value="NS">Nouvelle-Écosse</option>
                              <option value="NU">Nunavut</option>
                              <option value="ON">Ontario</option>
                              <option value="PE">l'île du Prince-Édouard </option>
                              <option value="QC">Québec</option>

                              <option value="SK">Saskatchewan</option>
                              <option value="YT">(territoire du) Yukon</option>
                              </select></td></tr> 
                          <tr> 
                            <td width="123">Code postal </td> 
                            <td width="237"><input name="ship_PostalCode" type="text" id="ship_PostalCode" size="20" value="<%=ship7%>"></td> 
                          </tr> 
                          <tr> 
                            <td width="123">Pays</td> 
                            <td width="237"><input name="ship_Country" type="text" id="ship_Country"  size="20" value="<%=ship8%>">
                            </td> 
                          </tr> 
                          <tr> 
                            <td width="123">Numéro de téléphone</td> 
                            <td width="237"><input name="HomePhone" type="text" id="HomePhone" size="20" value="<%=ship9%>"></td> 
                          </tr> 
                          <tr> 
                            <td>&nbsp;</td> 
                            <td>&nbsp;</td> 
                          </tr> 
                        </table> 
                        <p style="font-size:12px;color:#333333;">**Must be set if the address contains a suite number</p>
                      </li></td> 
                  </tr>  </table></form><form name="promo" method="post" id="promoform" onSubmit="javascript:fillship();"><input type="hidden" name="shipvalues" id="shipvalues" value="" /><table>
                  <tr> 
                    <td colspan="2" class="vslcss"><div class="t">
	<%if((Month(now())=12) and (day(now())<16) and (year(now())=2016)) then %>
				<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >Vente anniversaire </font></p>
				<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >Jusqu’au 15 décembre ou jusqu’à l’épuissement des stocks</font></p>
				<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >&Agrave;&nbsp;tous nos chers clients, Ferring Produits Pharmaceutiques &agrave; le plaisir de vous  annoncer &nbsp;une <em>vente anniversaire pour le VSL#3.</em></font></p>
    <p style="background-color: #FFFFCC;"><font color="#FF0000"  >Jusq&ugrave;au 15 décembre, 2016,  avec chaque achat de 3  bo&icirc;tes de VSL#3, vous recevrez une 4i&egrave;me bo&icirc;te sans frais, jusq&ugrave;&agrave; l’&egrave;puissement des stocks <br />
  </font></p>
					<%end if%>
       <!--<span class="ewmsg"> ATTENTION! <br>
La dernière journée pour placer votre commande de VSL#3 pour une livraison le lendemain sera mercredi le 21 décembre 2016. Toutes commandes placées après cette date seront livrées lorsque le service de livraison pour le lendemain reprends le 3 janvier 2017.
</span>--><br>
					<p><font size="2">
          Vous recevrez la confirmation de votre commande par courriel.</font></p>
Les boîtes postales ne peuvent pas être acceptées comme adresse de livraison
Veuillez noter que vous aurez besoin d'avoir quelqu'un à l'adresse de livraison le jour de la livraison pour recevoir le colis et s'assurer qu'il est réfrigéré.

					  Vous recevrez une confirmation et vous serez avisé de la date de livraison de votre commande par courriel ou par téléphone.
                      <p><font size="2"><strong>Les boîtes postales ne peuvent pas être acceptées comme adresse de livraison  </strong></font></p>
                      <p><font size="2">Veuillez noter que vous aurez besoin d'avoir quelqu'un à l'adresse de livraison le jour de la livraison pour recevoir le colis et s'assurer qu'il est réfrigéré.
    
                        </font></p>
                      
                      <font size="2"></font></div></td> 
                  </tr> 
                  <tr> 
                    <td colspan="2" class="vslcss"><p>
                      <% call DisplayItems() %>
                    </p>
				<!--	<div class="t">   
                      <label for="PromoCode">Code promotionnel :</label>
                      <input name="PromoCode" type="text" id="PromoCode" value="<%=session("promocode")%>">
                      <input type="submit" name="promosubmit" id="promosubmit" value="Accèpte">
                    <span id="promomsg" class="ewmsg"><br /><%=promomsg%></span></div>
-->
                    </td> 
                  </tr> 
                  <tr> 
                    <td colspan="2" align="center" valign="top"><p><a href="vslCart.asp"><img src="images/clicktoreturn_fr.gif" width="201" height="32" border="0" style="vertical-align:top!important;"></a>&nbsp; &nbsp; 
                        <input name="Submit2" type="image" class="InputNoBorder" id="Submit2"  value="Proceed to Secure Checkout" src="images/checkout_fr.gif" onclick="javascript:return PaypalSubmit();" style="border-style:none!important;border-color:transparent!important;boder-width:0px!important;">
                    &nbsp; <a href="logout.asp"><img src="images/logout_fr.gif" width="159" height="32" border="0" style="vertical-align:top!important;"></a> </p></td> 
                  </tr> 
                </table> 
              </form> 
              <script type="text/javascript">
			  function fillship() {
			  document.getElementById("shipvalues").value = document.getElementById("ship_FirstName").value + '||' 
					+ document.getElementById("ship_LastName").value + '||'
					+ document.getElementById("ship_Address").value + '||'
					+ document.getElementById("ship_Address2").value + '||'
					+ document.getElementById("ship_City").value + '||'
					+ document.getElementById("ship_Province").value + '||'
					+ document.getElementById("ship_PostalCode").value + '||'
					+ document.getElementById("ship_Country").value + '||'
					+ document.getElementById("HomePhone").value ;
					
			  }
	document.getElementById("ship_Province").value="<%=ship6%>";
    var myMenu = new ImageMenu($$('#kwick .kwick'),{openWidth:261,start:4});
	
  </script> 
              <%
	else
		response.write "Error.."
	end if
rs.close
set rs=nothing

end function
%> 

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
mcode= "" 'request.form("promocode")

If(mcode & "x"<>"x") then mcode=UCase(mcode)
if((mcode ="") and (session("promocode")<>"")) then mcode=session("promocode")
session("promocode")=mcode
applyDiscount= checkCustomer(conn,mcode)
		response.Write "<div class='t' width=720px><b>Panier d’achat:</b>"

%>
<table width="760px" >
  <tr>
    <td><table width="720px"  class="ewTable" cellspacing="5" cellpadding="5" border="1">
        <tr >
          <td width="50%" class="ewTableHeader"><b>Produits</b></td>
          <td width="16%" class="ewTableHeader"><div align="right"><b>Prix unitaire</b></div></td>
          <td width="12%" class="ewTableHeader"><div align="center"><b>Quantité</b></div></td>
          <td width="10%" class="ewTableHeader"><div align="right"><b>Total</b></div></td>
        </tr>
        <%GT=0
		For i=0 To UBound(arr)
			p=arr(i,2)
			t=arr(i,2)
				if(applyDiscount) then
				
					if(session("specialxprice")<>""	) then 
						p= cdbl(session("specialxprice"))
						t= "<p>Prix régulier : <s>" & arr(i,2) & "<br>"&vbCrLf 
						t=t & "</s><font color=""#FF0000"" size=""-1"">Prix spécial pour cette commande seulement</font>"
						t=t & ": <font color=""#FF0000""><strong>" & p  &"</strong></font> </p>"&vbCrLf
					end if
				end if
		
		%>
        <tr >
          <td width="50%"><b><%= "</b>" & arr(i,1) %></b></td>
          <td width="16%"><div align="right"><b><%=t%></b></div></td>
          <td width="12%" align="center"><div align="center">
              <%
			  dim tc
			  tc=arr(i,3) 
			  
			

			if(Session("FreexQty")<>"") then
				Response.write tc & ("( + " &  int(tc/cint(Session("FreexQty")) ) & " gratuit )" )
			else
				Response.write tc
				'call extraUnits(tc,"<br />+2 gratuitement")
				'if(tc>19) then  Response.write "<br />+2 gratuitement"
			end if
			
			  
			  %>
            </div></td>
          <td width="10%"><div align="right"><%=p *  arr(i,3)%></div></td>
        </tr>
        <%
		GT=GT+ p *  arr(i,3)
		Next%>
        <tr bordercolor="#FFCC66">
          <td width="50%" class="ewTablePager">
    <% if (isNewCustomer()) then%>
        <span class="ewmsg">Notez, la première boîte sera facturée à 99 $ aulieu du 110 $ lors de la transaction de PayPal.</span>
    <%end if %>
          </td>
          <td width="16%" class="ewTablePager">&nbsp;</td>
          <td width="12%" class="ewTablePager"><div align="right">Total:</div></td>
          <td width="10%" class="ewTablePager"><div align="right"><%=FormatNumber(GT,2)%></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<%if(applyDiscount) then 
	if(Session("freexship")) then response.write " Cette commande sera expédiée gratuitement"
end if
%>
<%
response.write "</div>"
end if
End Sub





function checkCustomer(c,code)
if(session("invalidtry") & "x"="x") then
'nothing
else
if(cint(session("invalidtry"))>10) then 
		Session("freexship")=false
		Session("specialxprice")=""
		Session("FreexQty")=""
		checkCustomer=false
		session("promocode")=""
		promomsg="Invalid promo codes entered too many times.. "
		exit function
end if
end if

	if(code<>"") then
		code = UCase(code)
		dim rs,NewCustomer
		Set rs = Server.CreateObject("ADODB.Recordset")
		strSql ="SELECT Discountcodes.DiscountCode, DiscountTypes.fDiscountTitle, DiscountTypes.DiscountType, DiscountTypes.freeShipping, DiscountTypes.FreePerQty, DiscountTypes.SpecialPrice, Orders.PromoCodeUsed "
		strSql = strSql & " FROM (Discountcodes INNER JOIN DiscountTypes ON Discountcodes.DiscountTypeId = DiscountTypes.DiscountTypeId) LEFT JOIN Orders ON Discountcodes.DiscountCode = Orders.PromoCodeUsed "
		strSql = strSql & " WHERE (((Discountcodes.DiscountCode)='"& mid(code,1,5) &"') AND ((Discountcodes.Active)=True)) AND ((DiscountTypes.StartDate)<Now()) AND ((DiscountTypes.EndDate)>Now()) "
		if(code<>"VSL14") then strSql = strSql & " AND((Discountcodes.used)=False) "
		
		'response.write strsql
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
				promomsg="Code promotionnel est accepté : " & rs.fields("fDiscountTitle")
			else
				Session("freexship")=false
				Session("specialxprice")=""
				Session("FreexQty")=""
				checkCustomer=false
				promomsg="Le code promotionel est accepté et la commande ce fait payé.  Ci le payment n'est pas réussit, le code sera dégage á moins de 24 heures."
				session("promocode")=""
			end if
			session("invalidtry")=0
		else
			Session("freexship")=false
			Session("specialxprice")=""
			Session("FreexQty")=""
			checkCustomer=false
			promomsg="Invalid promo Code"
			session("promocode")=""
			if(session("invalidtry") & "x"="x") then session("invalidtry")=0
			session("invalidtry") = cint(session("invalidtry"))+1
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



%>
