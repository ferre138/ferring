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
orderid= request.QueryString("orderId")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING

if(request.form("emailTo")<>"") then

 sFrEmail = request.form("emailFrom")
 sBccEmail="skarkera@ravenshoegroup.com"

 sToEmail = request.form("emailTo")
if( request.form("selectEmail")="confirm") then
	 sSubject = "Confirmation of shipment from VSL3.ca order no # " & request.form("hOrderid")  
else
	 sSubject =  "Cancellation of order from VSL3.ca order no # " & request.form("hOrderid")   
end if

 sbody = request.form("EmailText")
 //sbody = replace(sbody, vbcrlf,"<br>" )
 //sbody = replace(sbody, "/r/n","<br>" )
 sTextBody = sBody 
 sHTMLBody = "<body style=""font-family:Verdana, Verdana, Geneva, sans-serif; font-size:14px; color:#666666;"">" & sBody & "</body>"
 sbody = replace(sbody,"                   " & vbcrlf ,"")
 sMail = sBody
 Call Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)
 conn.Execute("update orders set EmailSent='" &  request.form("selectEmail") & "', EmailDate=#"& FormatDateTime(now(), 2) & "# where orderid= " &  request.form("hOrderid"))
'response.write "email sent"
'response.end
orderid =request.form("hOrderid")
end if



strSQL = "SELECT payer_email,Ship_FirstName,Ship_LastName,Ship_Address,Ship_Address2,Ship_City,Ship_Province,Ship_Postal,Ship_Country,Ship_Phone, productId,Quantity,orderdetails.price as oprice, Amount, Tax, Shipping FROM orders,orderdetails "
strSQL = strSQL & "WHERE orders.orderId = orderdetails.orderId and orders.orderId =" & orderid & ";"

Set rst = conn.Execute(strSQL)
Set rs = Server.CreateObject("ADODB.Recordset")
totalAmount = CDbl(rst("Amount"))
Tx = CDbl(rst("Tax"))
ship = CDbl(rst("Shipping"))
strHtml="<table cellspacing=""0"" cellpadding=""0""  border='0' width=600px style='border: thin solid #999;'>"
counter=0
Do While not rst.EOF
	
	qProd = "SELECT ItemNo, description, Price FROM Products WHERE ItemId =" & rst("ProductId") & ";"
	rs.open qProd,conn					
	if counter = 0 Then
		strHtml = strHtml &  "<tr height=25px>"
		strHtml = strHtml & "<td align='left' valign='top'>Item #</td>"						
		strHtml = strHtml & "<td align='left' valign='top'>Description</td>"						
		strHtml = strHtml & "<td align='center' valign='top'>Quantity</td>"
		strHtml = strHtml & "<td align='right' valign='top'>Price</td>"						
		strHtml = strHtml & "<td align='right' valign='top'>Amount</td>"						
		strHtml = strHtml & "</tr>"
    	payer_email = rst("payer_email")
		pShipName = rst("Ship_FirstName") & " " & rst("Ship_LastName")
		pShipAddr = rst("Ship_Address")
		pShipAddr2 = rst("Ship_Address2")
		pShipCity = rst("Ship_City")
		pShipProv = rst("Ship_Province")
		pShipPostal = rst("Ship_Postal")
		pShipCountry = rst("Ship_Country")
		pShipPhone = rst("Ship_Phone")
	End if 
	strHtml = strHtml & "<tr height=25px>"
	strHtml = strHtml & "<td align='left' valign='top'>" & rs("ItemNo") & "</td>"
	strHtml = strHtml & "<td align='left' valign='top'>" & rs("description") & "</td>"
	strHtml = strHtml & "<td align='center' valign='top'>" & rst("Quantity") & "</td>"						
	strHtml = strHtml & "<td align='center' valign='top'>$" & rst("oprice")  & "</td>"
	strHtml = strHtml & "<td align='right' valign='top'>$" & FormatNumber(rst("oprice") * rst("Quantity"),2) & "</td>"
	strHtml = strHtml & "</tr>"
	counter = 1
	rs.close
	rst.MoveNext						
Loop

'strHtml = strHtml &  "<tr >"
'strHtml = strHtml & "<td colspan='4' align='right' valign='top'>Shipping and handling:</td><td align='right'  valign='top'>$" & FormatNumber(ship,2) & "</td>"
'strHtml = strHtml &  "</tr>"

strHtml = strHtml &  "<tr>"
strHtml = strHtml &  "<td colspan='4' align='right' valign='top'>Tax:</td><td align='right' valign='top'>$" & FormatNumber(tx,2) & "</td>"
strHtml = strHtml &  "</tr>"
				
strHtml = strHtml &  "<tr height='36'>"
strHtml = strHtml &  "<td colspan='4' align='right' valign='top'><b>Total Amount: </td><td align='right' valign='top'><b>$" & FormatNumber(totalAmount,2) & "</b></td>"
strHtml = strHtml &  "</tr>"
strHtml = strHtml &  "</table>"		
	
strAddress= ""
strAddress= strAddress & pShipAddr & "<br/>"
strAddress= strAddress & pShipAddr2 & "<br/>"

strAddress= strAddress & pShipCity & "," & pShipProv & "-" & pShipPostal & "<br/>"
strAddress= strAddress & pShipCountry & "<br/>"
strAddress= strAddress & "Tel:" &  pShipPhone & "<br/>"



sMail =getFileString("shipEmail.txt")
sMail=replace(sMail,"VSLORDER_DETAILS",strHtml)	
sMail=replace(sMail,"VSLORDER_SHIPPING",strAddress)	
sMail=replace(sMail,"VSLFIRST_NAME",pShipName)	
sMail=replace(sMail,"VSLORDER_NO",orderid)	
sMail=replace(sMail,"VSLBILLNUM","VSL" & mid(pShipPostal,1,3) )	


fMail =getFileString("failEmail.txt")
fMail=replace(fMail,"VSLORDER_DETAILS",strHtml)	
fMail=replace(fMail,"VSLFIRST_NAME",pShipName)	
fMail=replace(fMail,"VSLORDER_NO",orderid)	
			
rst.close  ' Close Recordset
conn.Close ' Close Connection
%>
<!-- Put your custom html here -->
<link href="css/vslpaypal.css" rel="stylesheet" type="text/css" />

<form id="formEmail" name="formEmail" method="post" action="sendEmail.asp">
  <table width="100%" border="0">
    <tr>
      <td>Email From</td>
      <td><label for="emailFrom"></label>
        <input name="emailFrom" type="text" id="emailFrom" value="info@vsl3.ca" /></td>
    </tr>
    <tr>
      <td>Email To</td>
      <td><input name="emailTo" type="text" id="emailTo" value="<%=payer_email%>" size="40"/></td>
    </tr>
    <tr>
      <td>Email Text</td>
      <td><label for="selectEmail"></label>
        <select name="selectEmail" id="selectEmail" onchange="reloadtxt(this);">
          <option value="confirm">Send Confirmation</option>
          <option value="cancel">Send Cancelled</option>
        </select></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><textarea name="EmailText" cols="120" rows="40" id="EmailText"><%=sMail%></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="button" id="button" value="Send Email" />
        <input name="hOrderid" type="hidden" id="hOrderid" value="<%=orderid%>" /></td>
    </tr>
  </table>
  <div id="ship_email" style="visibility:hidden"><%=sMail%><div> <div id="fail_email"  style="visibility:hidden"><%=fMail%><div>
</form>
<p>&nbsp;</p>
<p> 
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
%>
</p>
