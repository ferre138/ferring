<%
Const EW_PAGE_ID = "register"
%>
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%

' Open connection to the database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
%>
<%
Dim Security
Set Security = New cAdvancedSecurity
%>
<%

' Common page loading event (in userfn60.asp)
Call Page_Loading()
%>
<%

' Page load event, used in current page
Call Page_Load()
%>
<%
Dim bUserExists
Dim captcha
Response.Buffer = True

' Create form object
Dim objForm
Set objForm = New cFormObj
If objForm.GetValue("a_register")&"" <> "" Then

	' Get action
	Customers.CurrentAction = objForm.GetValue("a_register")
	Call LoadFormValues() ' Get form values
Else
	Customers.CurrentAction = "I" ' Display blank record
	Call LoadDefaultValues() ' Load default values
End If
If Customers.CurrentAction <> "I" And Customers.CurrentAction <> "C" Then

	' Get captcha value
	captcha = objForm.GetValue("captcha")

	' Check captcha value from form
	If captcha <> Session("CAPTCHA") Then ' Captcha matched
		Session(EW_SESSION_MESSAGE) = "Please enter the validation code shown" ' Set message
		Customers.CurrentAction = "I" ' Reset action, do not insert if captcha unmatched
	End If
End If

' Close form object
Set objForm = Nothing
Select Case Customers.CurrentAction
	Case "I" ' Blank record, no action required
	Case "A" ' Add

		' Check for Duplicate User ID
		Dim sFilter, sUserSql, rs
		sFilter = "([UserName] = '" & ew_AdjustSql(Customers.UserName.CurrentValue) & "')"

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in Customers class, Customersinfo.asp

		Customers.CurrentFilter = sFilter
		sUserSql = Customers.SQL
		Set rs = conn.Execute(sUserSql)
		If Not rs.Eof Then
			bUserExists = True
			Call RestoreFormValues() ' Restore form values
			Session(EW_SESSION_MESSAGE) = "User Already Exists!" ' Set user exist message
		End If
		rs.Close
		Set rs = Nothing
		If Not bUserExists Then
			Customers.SendEmail = True ' Send email on add success
			If AddRow() Then ' Add record
				Session(EW_SESSION_MESSAGE) = "Registration Successful" ' Register success
				Call Page_Terminate("login.asp") ' Go to login page
			Else
				Call RestoreFormValues() ' Restore form values
			End If
		End If
End Select

' Render row
Customers.RowType = EW_ROWTYPE_ADD ' Render add
Call RenderRow()
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
var EW_PAGE_ID = "register"; // Page id
//-->
</script>
<script type="text/javascript">
<!--
function ew_ValidateForm(fobj) {
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = (fobj.key_count) ? Number(fobj.key_count.value) : 1;
	for (i=0; i<rowcnt; i++) {
		infix = (fobj.key_count) ? String(i+1) : "";
		elm = fobj.elements["x" + infix + "_Inv_FirstName"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - First Name"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_Inv_LastName"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Last Name"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_Inv_Address"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Billing Address"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_City"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - City"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_Province"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Province"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_PostalCode"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Postal Code"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_PhoneNumber"];
		if (elm && !ew_CheckPhone(elm.value)) {
			if (!ew_OnError(elm, "Incorrect phone number - Phone Number"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_EmailAddress"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - Email Address"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_EmailAddress"];
		if (elm && !ew_CheckEmail(elm.value)) {
			if (!ew_OnError(elm, "Please Enter a Valid Email"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_inv_Fax"];
		if (elm && !ew_CheckPhone(elm.value)) {
			if (!ew_OnError(elm, "Incorrect phone number - Fax"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_UserName"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - User Name"))
				return false;
		}
		elm = fobj.elements["x" + infix + "_passwrd"];
		if (elm && !ew_HasValue(elm)) {
			if (!ew_OnError(elm, "Please enter required field - passwrd"))
				return false;
		}
		if (fobj.x_passwrd && !ew_HasValue(fobj.x_passwrd)) {
			if (!ew_OnError(fobj.x_passwrd, "Please enter password"))
				return false; 
		}
		if (fobj.c_passwrd.value != fobj.x_passwrd.value) {
			if (!ew_OnError(fobj.c_passwrd, "Mismatch Password"))
				return false; 
		}
	}
		if (fobj.captcha && !ew_HasValue(fobj.captcha)) {
			if (!ew_OnError(fobj.captcha, "Please enter the validation code shown"))
				return false;
		}
	return true;
}

//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script type="text/javascript">
<!--
var ew_MultiPagePage = "Page"; // multi-page Page Text
var ew_MultiPageOf = "of"; // multi-page Of Text
var ew_MultiPagePrev = "Prev"; // multi-page Prev Text
var ew_MultiPageNext = "Next"; // multi-page Next Text
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
// To include another .js script, use:
// ew_ClientScriptInclude("my_javascript.js"); 
//-->
</script>
<table  border="0" cellpadding="0" cellspacing="0" id="Table_01">
            <tr>
            <td  width="699" height="75" rowspan="2"><span class="Header">Politique de Ferring sur la protection des renseignements personnels</span></td>
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
                <p><a href="../privacy.asp" class="bodycopy_small">english &gt;</a></p>
              </div></td>
              </tr>
        </table>
<p align="justify" class="vslcss">Ferring vous remercie de visiter ce site web. Parce que nous sommes engag&eacute;s &agrave;   respecter la confidentialit&eacute; des visiteurs de notre site web, la politique sur   la protection des renseignements personnels de ce site d&eacute;crit les renseignements   que nous pouvons recueillir de vous pendant votre visite &agrave; notre site, de quelle   fa&ccedil;on nous pouvons les utiliser et comment nous prot&eacute;gerons les renseignements   que vous pouvez choisir de nous fournir. Cette politique explique les efforts   pour &eacute;quilibrer nos int&eacute;r&ecirc;ts d'affaires en recueillant et utilisant les   renseignements que nous recevons de vous avec votre besoin pour une gestion   convenable de tout renseignement vous identifiant personnellement et que vous   partagez avec nous. <br>
    <br>
    <span class="subheading">Renseignements vous identifiant personnellement 
</span> <br>
    Les renseignements vous identifiant personnellement comprennent vos nom,   adresse, num&eacute;ro de t&eacute;l&eacute;phone, adresse courriel ou tout autre renseignement   pouvant &ecirc;tre utilis&eacute; pour vous identifier personnellement. Ferring recueille les   renseignements vous identifiant personnellement quand vous visitez le site web   et seulement quand ils sont volontairement fournis. Autrement, Ferring ne   recueillera aucun renseignement de vous sur notre site web. Quand Ferring re&ccedil;oit   des renseignements vous identifiant personnellement, nous nous r&eacute;servons le   droit de les utiliser pour des fins d'affaires raisonnables comme peuvent le   faire toutes entreprises autrement que par Internet. En fournissant &agrave; Ferring   via ce site web, des renseignements vous identifiant personnellement, vous   consentez automatiquement &agrave; ce que nous les utilisions &agrave; de telles fins. Par   exemple, nous pouvons utiliser ces renseignements pour communiquer avec vous via   courriel ou poste ordinaire afin de vous fournir de l'information que nous   croyons &ecirc;tre dans votre int&eacute;r&ecirc;t. Ces renseignements peuvent aussi &ecirc;tre utilis&eacute;s   pour compiler des donn&eacute;es et faire des analyses afin de comprendre et servir les   besoins de notre client&egrave;le. Les donn&eacute;es sont &eacute;galement compil&eacute;es pour &eacute;valuer   l'usage et l'utilit&eacute; des services que nous fournissons en ligne. Vos   renseignements peuvent &ecirc;tre transf&eacute;r&eacute;s et utilis&eacute;s par des installations Ferring   dans d'autres pays et r&eacute;gions du monde. Les renseignements vous identifiant   personnellement ne seront pas vendus, lou&eacute;s ou &eacute;chang&eacute;s &agrave; l'ext&eacute;rieur du Groupe   Ferring &agrave; moins que l'usager n'en soit d'abord avis&eacute; et donne son consentement   expr&egrave;s &agrave; un tel transfert. <br><br>
    </p>
<p align="justify" class="vslcss"><span class="subheading">L'assemblage des données</span> <br>
  Le site web de Ferring peut aussi recueillir des renseignements de vous ne vous   identifiant pas personnellement. Par exemple, nous pouvons retracer la date et   l'heure o&ugrave; les visiteurs ont acc&eacute;d&eacute; &agrave; notre site, le type d'explorateur web   qu'ils utilisent et le site web duquel ils ont connect&eacute; &agrave; notre site. Nos sites   web recueillent ces renseignements en d&eacute;posant certaines pi&egrave;ces d'information   appel&eacute;es " t&eacute;moins " dans l'ordinateur d'un visiteur. Cette technologie ne   recueillent pas les renseignements permettant d'identifier personnellement un   visiteur individuel. Ces renseignements sont plut&ocirc;t recueillis dans une forme   assembl&eacute;e. Les t&eacute;moins peuvent nous dire comment et quand les pages d'un site   web ont &eacute;t&eacute; visit&eacute;es et par combien de personnes. L'assemblage de ces donn&eacute;es va   nous permettre d'am&eacute;liorer nos sites web afin de mieux vous servir et informer.   Cela peut &eacute;galement vous permettre d'acc&eacute;der plus rapidement &agrave; des points   d'int&eacute;r&ecirc;ts de notre site web quand vous entrez de nouveau dans notre syst&egrave;me. La   fixation de cet appareil sur votre syst&egrave;me n'a aucun effet sur son rendement.</p>
<p align="justify" class="vslcss">&nbsp;</p>
<p align="justify" class="vslcss"><span class="subheading">Liens à d'autres sites</span> <br>
  Comme ressources pour nos visiteurs, Ferring peut fournir des liens &agrave;  d'autres sites web. Nous essayons de s&eacute;lectionner soigneusement des  sites web que nous croyons utiles et qui satisfont &agrave; nos normes &eacute;lev&eacute;es  pour l'exactitude et l'utilit&eacute; de l'information. Toutefois, le contenu  et la conception d'un site web pouvant changer rapidement, nous ne  pouvons garantir les normes de chaque site web auquel nous nous lions.  Dans la m&ecirc;me veine, nous ne sommes pas responsable du contenu de tout  site autre que celui de Ferring. Nous ne pouvons &eacute;galement pas garantir  les politiques de confidentialit&eacute; de ces autres sites et recommandons  que vous v&eacute;rifiiez les politiques sur la confidentialit&eacute; directement  avec ces autres sites.<br>
  <br>
</p>
<p align="justify" class="vslcss"><span class="subheading">Choix
</span> <br>
Si vous souhaitez cesser de recevoir tout courriel ou autre communication de   Ferring pouvant vous &ecirc;tre envoy&eacute; dans l'avenir suite &agrave; votre demande pour cette   information ou si vous avez soumis des renseignements vous identifiant   personnellement par le biais d'un site web Ferring et que vous souhaitiez que   cette information soit supprim&eacute;e de nos dossiers, veuillez en aviser l'Agent du   Service de la protection de la vie priv&eacute;e de Ferris. Vos souhaits en ces   mati&egrave;res seront respect&eacute;s. <br>
        <br>
        <span class="subheading">Exactitude</span> <br>
        Ferring fera tous les efforts pour maintenir l'exactitude et la confidentialit&eacute;   de tout renseignement personnel que vous pouvez nous fournir. Si vous souhaitez   rajouter ou effectuer des corrections aux renseignements que vous nous avez   envoy&eacute;, veuillez aviser l'Agent du Service de la protection de la vie priv&eacute;e de   Ferring.<br>
        <br>
        <span class="subheading">Sécurité</span> <br>
        Tous les renseignements transmis &agrave; ce site web de Ferring sont s&ucirc;rs dans la   mesure du possible avec l'utilisation de la technologie existante. Il ne devrait   pas &ecirc;tre possible &agrave; des tierces parties d'acc&eacute;der &agrave; ces renseignements pendant   la transmission. Nous conserverons de fa&ccedil;on s&ucirc;re les renseignements que vous   partagez avec nous et prendrons les mesures appropri&eacute;es pour les prot&eacute;ger d'un   acc&egrave;s ou d'une diffusion non autoris&eacute;. Bien qu'aucune mesure de s&eacute;curit&eacute; ne   puisse offrir une protection totale, nous utilisons une technologie et des   syst&egrave;mes ultramodernes pour pr&eacute;venir l'acc&egrave;s non autoris&eacute; aux renseignements que   nous d&eacute;tenons. Nous limiterons l'acc&egrave;s &agrave; ces renseignements aux employ&eacute;(e)s de   Ferring pour lesquels cela est n&eacute;cessaire. Nous instruisons notre personnel de   leur devoir de prot&eacute;ger ce qui vous est confidentiel.<br>
        <br>
        <span class="subheading">Enfants</span> <br>
        Ferring ne recueille ni ne conserve intentionnellement des renseignements   identifiant personnellement des enfants qui n'ont pas atteint l'&acirc;ge de dix-huit   ans. Si votre enfant peut nous avoir soumis des renseignements sans avoir   indiquer son &acirc;ge actuel et que vous souhaitiez faire retirer ces renseignements,   veuillez en aviser l'Agent du Service de la protection de la vie priv&eacute;e de   Ferring et nous supprimerons ces renseignements imm&eacute;diatement. <br>
        <br>
        <span class="subheading">Changements</span> <br>
</p>
<p>Tout changement &agrave; cette politique sur la protection des renseignements   personnels sera promptement communiqu&eacute; sur ce site. Veuillez v&eacute;rifier   r&eacute;guli&egrave;rement la page sur la protection des renseignements personnels pour   examiner tout changement pouvant avoir &eacute;t&eacute; apport&eacute;.<br><br></p>
  <p>Merci de votre visite sur le site web de Ferring. Nous   appr&eacute;cions votre int&eacute;r&ecirc;t et vos id&eacute;es. Si vous avez des commentaires ou   pr&eacute;occupations en ce qui concerne l'utilisation des renseignements fournis &agrave;   Ferring via un site Internet, veuillez communiquer avec l'Agent du service de la   protection de la vie priv&eacute;e.</p>
  <p align="justify" class="vslcss"><br>
      <br>
  </p>
  <p><span class="subheading">Poste régulière
</span> <br>
  Agent du Service de la protection de la vie privée
 <br>
  Ferring Inc. <br>
  200 Yorkland Blvd., Suite 500 <br>
  Toronto, Ontario <br>
  M2J 5C1 </p>
<p>Courriel <br>
    <a href="mailto:Privacy.officer@ferring.com">Privacy.officer@ferring.com </a></p>
<!--#include file="footer.asp"-->
<script language="JavaScript" type="text/javascript">
<!--
// Write your startup script here
// document.write("page loaded");
//-->
</script>
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

	' Page unload event, used in current page
	Call Page_Unload()

	' Global page unloaded event (in userfn60.asp)
	Call Page_Unloaded()
	conn.Close ' Close Connection
	Set conn = Nothing
	Set Security = Nothing

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

%>
<%

' Load default values
Function LoadDefaultValues()
	Customers.inv_Country.CurrentValue = "Canada"
End Function
%>
<%

' Load form values
Function LoadFormValues()

	' Load from form
	Customers.Inv_FirstName.FormValue = objForm.GetValue("x_Inv_FirstName")
	Customers.Inv_LastName.FormValue = objForm.GetValue("x_Inv_LastName")
	Customers.Inv_Address.FormValue = objForm.GetValue("x_Inv_Address")
	Customers.Inv_Address2.FormValue = objForm.GetValue("x_Inv_Address2")
	Customers.inv_City.FormValue = objForm.GetValue("x_inv_City")
	Customers.inv_Province.FormValue = objForm.GetValue("x_inv_Province")
	Customers.inv_PostalCode.FormValue = objForm.GetValue("x_inv_PostalCode")
	Customers.inv_Country.FormValue = objForm.GetValue("x_inv_Country")
	Customers.inv_PhoneNumber.FormValue = objForm.GetValue("x_inv_PhoneNumber")
	Customers.inv_EmailAddress.FormValue = objForm.GetValue("x_inv_EmailAddress")
	Customers.inv_Fax.FormValue = objForm.GetValue("x_inv_Fax")
	Customers.UserName.FormValue = objForm.GetValue("x_UserName")
	Customers.passwrd.FormValue = objForm.GetValue("x_passwrd")
End Function

' Restore form values
Function RestoreFormValues()
	Customers.Inv_FirstName.CurrentValue = Customers.Inv_FirstName.FormValue
	Customers.Inv_LastName.CurrentValue = Customers.Inv_LastName.FormValue
	Customers.Inv_Address.CurrentValue = Customers.Inv_Address.FormValue
	Customers.Inv_Address2.CurrentValue = Customers.Inv_Address2.FormValue
	Customers.inv_City.CurrentValue = Customers.inv_City.FormValue
	Customers.inv_Province.CurrentValue = Customers.inv_Province.FormValue
	Customers.inv_PostalCode.CurrentValue = Customers.inv_PostalCode.FormValue
	Customers.inv_Country.CurrentValue = Customers.inv_Country.FormValue
	Customers.inv_PhoneNumber.CurrentValue = Customers.inv_PhoneNumber.FormValue
	Customers.inv_EmailAddress.CurrentValue = Customers.inv_EmailAddress.FormValue
	Customers.inv_Fax.CurrentValue = Customers.inv_Fax.FormValue
	Customers.UserName.CurrentValue = Customers.UserName.FormValue
	Customers.passwrd.CurrentValue = Customers.passwrd.FormValue
End Function
%>
<%

' Render row values based on field settings
Sub RenderRow()

	' Call Row Rendering event
	Call Customers.Row_Rendering()

	' Common render codes for all row types
	' Inv_FirstName

	Customers.Inv_FirstName.CellCssStyle = ""
	Customers.Inv_FirstName.CellCssClass = ""

	' Inv_LastName
	Customers.Inv_LastName.CellCssStyle = ""
	Customers.Inv_LastName.CellCssClass = ""

	' Inv_Address
	Customers.Inv_Address.CellCssStyle = ""
	Customers.Inv_Address.CellCssClass = ""

	' Inv_Address2
	Customers.Inv_Address2.CellCssStyle = ""
	Customers.Inv_Address2.CellCssClass = ""

	' inv_City
	Customers.inv_City.CellCssStyle = ""
	Customers.inv_City.CellCssClass = ""

	' inv_Province
	Customers.inv_Province.CellCssStyle = ""
	Customers.inv_Province.CellCssClass = ""

	' inv_PostalCode
	Customers.inv_PostalCode.CellCssStyle = ""
	Customers.inv_PostalCode.CellCssClass = ""

	' inv_Country
	Customers.inv_Country.CellCssStyle = ""
	Customers.inv_Country.CellCssClass = ""

	' inv_PhoneNumber
	Customers.inv_PhoneNumber.CellCssStyle = ""
	Customers.inv_PhoneNumber.CellCssClass = ""

	' inv_EmailAddress
	Customers.inv_EmailAddress.CellCssStyle = ""
	Customers.inv_EmailAddress.CellCssClass = ""

	' inv_Fax
	Customers.inv_Fax.CellCssStyle = ""
	Customers.inv_Fax.CellCssClass = ""

	' UserName
	Customers.UserName.CellCssStyle = ""
	Customers.UserName.CellCssClass = ""

	' passwrd
	Customers.passwrd.CellCssStyle = ""
	Customers.passwrd.CellCssClass = ""
	If Customers.RowType = EW_ROWTYPE_VIEW Then ' View row
	ElseIf Customers.RowType = EW_ROWTYPE_ADD Then ' Add row

		' Inv_FirstName
		Customers.Inv_FirstName.EditCustomAttributes = ""
		Customers.Inv_FirstName.EditValue = ew_HtmlEncode(Customers.Inv_FirstName.CurrentValue)

		' Inv_LastName
		Customers.Inv_LastName.EditCustomAttributes = ""
		Customers.Inv_LastName.EditValue = ew_HtmlEncode(Customers.Inv_LastName.CurrentValue)

		' Inv_Address
		Customers.Inv_Address.EditCustomAttributes = ""
		Customers.Inv_Address.EditValue = ew_HtmlEncode(Customers.Inv_Address.CurrentValue)

		' Inv_Address2
		Customers.Inv_Address2.EditCustomAttributes = ""
		Customers.Inv_Address2.EditValue = ew_HtmlEncode(Customers.Inv_Address2.CurrentValue)

		' inv_City
		Customers.inv_City.EditCustomAttributes = ""
		Customers.inv_City.EditValue = ew_HtmlEncode(Customers.inv_City.CurrentValue)

		' inv_Province
		Customers.inv_Province.EditCustomAttributes = ""
		sSqlWrk = "SELECT [Prov], [Province] FROM [Province]"
		sSqlWrk = sSqlWrk & " ORDER BY [Province] Asc"
		Set rswrk = Server.CreateObject("ADODB.Recordset")
		rswrk.Open sSqlWrk, conn
		If Not rswrk.Eof Then
			arwrk = rswrk.GetRows
		Else
			arwrk = ""
		End If
		rswrk.Close
		Set rswrk = Nothing
		arwrk = ew_AddItemToArray(arwrk, 0, Array("", "Please Select"))
		Customers.inv_Province.EditValue = arwrk

		' inv_PostalCode
		Customers.inv_PostalCode.EditCustomAttributes = ""
		Customers.inv_PostalCode.EditValue = ew_HtmlEncode(Customers.inv_PostalCode.CurrentValue)

		' inv_Country
		Customers.inv_Country.EditCustomAttributes = ""
		Customers.inv_Country.EditValue = ew_HtmlEncode(Customers.inv_Country.CurrentValue)

		' inv_PhoneNumber
		Customers.inv_PhoneNumber.EditCustomAttributes = ""
		Customers.inv_PhoneNumber.EditValue = ew_HtmlEncode(Customers.inv_PhoneNumber.CurrentValue)

		' inv_EmailAddress
		Customers.inv_EmailAddress.EditCustomAttributes = ""
		Customers.inv_EmailAddress.EditValue = ew_HtmlEncode(Customers.inv_EmailAddress.CurrentValue)

		' inv_Fax
		Customers.inv_Fax.EditCustomAttributes = ""
		Customers.inv_Fax.EditValue = ew_HtmlEncode(Customers.inv_Fax.CurrentValue)

		' UserName
		Customers.UserName.EditCustomAttributes = ""
		Customers.UserName.EditValue = ew_HtmlEncode(Customers.UserName.CurrentValue)

		' passwrd
		Customers.passwrd.EditCustomAttributes = ""
		Customers.passwrd.EditValue = Customers.passwrd.CurrentValue
	ElseIf Customers.RowType = EW_ROWTYPE_EDIT Then ' Edit row
	ElseIf Customers.RowType = EW_ROWTYPE_SEARCH Then ' Search row
	End If

	' Call Row Rendered event
	Call Customers.Row_Rendered()
End Sub
%>
<%

' Add record
Function AddRow()
	On Error Resume Next
	Dim rs, sSql, sFilter
	Dim rsnew
	Dim bCheckKey, sSqlChk, sWhereChk, rsChk
	Dim bInsertRow

	' Check if valid user id
	Dim bValidUser
	bValidUser = False
	If Security.CurrentUserID <> "" And Not Security.IsAdmin Then ' Non system admin
		bValidUser = Security.IsValidUserID(Customers.CustomerID.CurrentValue)
		If Not bValidUser Then
			Session(EW_SESSION_MESSAGE) = "Unauthorized"
			AddRow = False
			Exit Function
		End If
	End If

	' Check for duplicate key
	bCheckKey = True
	sFilter = Customers.SqlKeyFilter
	If Customers.CustomerID.CurrentValue = "" Or IsNull(Customers.CustomerID.CurrentValue) Then
		bCheckKey = False
	Else
		sFilter = Replace(sFilter, "@CustomerID@", ew_AdjustSql(Customers.CustomerID.CurrentValue)) ' Replace key value
	End If
	If Not IsNumeric(Customers.CustomerID.CurrentValue) Then
		bCheckKey = False
	End If
	If bCheckKey Then
		Set rsChk = Customers.LoadRs(sFilter)
		If Not (rsChk Is Nothing) Then
			Session(EW_SESSION_MESSAGE) = "Duplicate value for primary key"
			rsChk.Close
			Set rsChk = Nothing
			AddRow = False
			Exit Function
		End If
	End If

	' Add new record
	sFilter = "(0 = 1)"
	Customers.CurrentFilter = sFilter
	sSql = Customers.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = EW_CURSORLOCATION
	rs.Open sSql, conn, 1, 2
	rs.AddNew
	If Err.Number <> 0 Then
		Session(EW_SESSION_MESSAGE) = Err.Description
		rs.Close
		Set rs = Nothing
		AddRow = False
		Exit Function
	End If

	' Field Inv_FirstName
	Call Customers.Inv_FirstName.SetDbValue(Customers.Inv_FirstName.CurrentValue, Null)
	rs("Inv_FirstName") = Customers.Inv_FirstName.DbValue

	' Field Inv_LastName
	Call Customers.Inv_LastName.SetDbValue(Customers.Inv_LastName.CurrentValue, Null)
	rs("Inv_LastName") = Customers.Inv_LastName.DbValue

	' Field Inv_Address
	Call Customers.Inv_Address.SetDbValue(Customers.Inv_Address.CurrentValue, Null)
	rs("Inv_Address") = Customers.Inv_Address.DbValue

	' Field Inv_Address2
	Call Customers.Inv_Address2.SetDbValue(Customers.Inv_Address2.CurrentValue, Null)
	rs("Inv_Address2") = Customers.Inv_Address2.DbValue

	' Field inv_City
	Call Customers.inv_City.SetDbValue(Customers.inv_City.CurrentValue, Null)
	rs("inv_City") = Customers.inv_City.DbValue

	' Field inv_Province
	Call Customers.inv_Province.SetDbValue(Customers.inv_Province.CurrentValue, Null)
	rs("inv_Province") = Customers.inv_Province.DbValue

	' Field inv_PostalCode
	Call Customers.inv_PostalCode.SetDbValue(Customers.inv_PostalCode.CurrentValue, Null)
	rs("inv_PostalCode") = Customers.inv_PostalCode.DbValue

	' Field inv_Country
	Call Customers.inv_Country.SetDbValue(Customers.inv_Country.CurrentValue, Null)
	rs("inv_Country") = Customers.inv_Country.DbValue

	' Field inv_PhoneNumber
	Call Customers.inv_PhoneNumber.SetDbValue(Customers.inv_PhoneNumber.CurrentValue, Null)
	rs("inv_PhoneNumber") = Customers.inv_PhoneNumber.DbValue

	' Field inv_EmailAddress
	Call Customers.inv_EmailAddress.SetDbValue(Customers.inv_EmailAddress.CurrentValue, Null)
	rs("inv_EmailAddress") = Customers.inv_EmailAddress.DbValue

	' Field inv_Fax
	Call Customers.inv_Fax.SetDbValue(Customers.inv_Fax.CurrentValue, Null)
	rs("inv_Fax") = Customers.inv_Fax.DbValue

	' Field UserName
	Call Customers.UserName.SetDbValue(Customers.UserName.CurrentValue, Null)
	rs("UserName") = Customers.UserName.DbValue

	' Field passwrd
	Call Customers.passwrd.SetDbValue(Customers.passwrd.CurrentValue, Null)
	rs("passwrd") = Customers.passwrd.DbValue

	' Check recordset update error
	If Err.Number <> 0 Then
		Session(EW_SESSION_MESSAGE) = Err.Description
		rs.Close
		Set rs = Nothing
		AddRow = False
		Exit Function
	End If

	' Call Row Inserting event
	bInsertRow = Customers.Row_Inserting(rs)
	If bInsertRow Then

		' Clone new rs object
		Set rsnew = ew_CloneRs(rs)
		rs.Update
		If Err.Number <> 0 Then
			Session(EW_SESSION_MESSAGE) = Err.Description
			AddRow = False
		Else
			AddRow = True
		End If
	Else
		rs.CancelUpdate
		If Customers.CancelMessage <> "" Then
			Session(EW_SESSION_MESSAGE) = Customers.CancelMessage
			Customers.CancelMessage = ""
		Else
			Session(EW_SESSION_MESSAGE) = "Insert cancelled"
		End If
		AddRow = False
	End If
	rs.Close
	Set rs = Nothing
	If AddRow Then
		Customers.CustomerID.DbValue = rsnew("CustomerID")

		' Call Row Inserted event
		Call Customers.Row_Inserted(rsnew)
	End If
	If IsObject(rsnew) Then
		rsnew.Close
		Set rsnew = Nothing
	End If
End Function
%>
<%

' Page Load event
Sub Page_Load()

'***Response.Write "Page Load"
End Sub

' Page Unload event
Sub Page_Unload()

'***Response.Write "Page Unload"
End Sub
%>
