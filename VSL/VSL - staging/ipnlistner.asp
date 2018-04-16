<!--#include file="ewcfg60.asp"-->
<% 
Dim Item_name, Item_number, Payment_status, Payment_amount
 Dim Txn_id, Receiver_email, Payer_email
 Dim objHttp, str 
 'read post from PayPal system and add 'cmd'
 str = Request.Form & "&cmd=_notify-validate"
  
 'post back to PayPal system to validate
 'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
 set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
 'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
 'set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
objHttp.setProxy 2, "ntproxy.1and1.com:3128"
 '"https://www.sandbox.paypal.com/cgi-bin/webscr"
 'https://www.paypal.com/cgi-bin/webscr
 
 objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
 ' objHttp.open "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false
 objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
 WriteToFile "log.txt" ,   str & now & vbcrlf,true 
 objHttp.Send str
WriteToFile "log.txt" ,   "xxxxx" & str & now & vbcrlf,true 

 

  
'assign posted variables to local variables
 first_name      = ForSQL(Request.Form("first_name"))
 last_name       = ForSQL(Request.Form("last_name"))
 address_name    = ForSQL(Request.Form("address_name"))
 address_street  = ForSQL(Request.Form("address_street"))
 address_city    = ForSQL(Request.Form("address_city"))
 address_state   = ForSQL(Request.Form("address_state"))
 address_country = ForSQL(Request.Form("address_country"))
 address_zip     = ForSQL(Request.Form("address_zip"))
 if IsNull(Request.Form("quantity")) Then
	quantity = 0 
 else
	quantity = ForSQL(Request.Form("quantity"))
 End if
 if IsNull(Request.Form("mc_fee")) Then
	mc_fee = 0 
 else
	mc_fee = ForSQL(Request.Form("mc_fee"))
 End if
 ' mc_gross = Request.Form("mc_gross_1")
 if IsNull(Request.Form("mc_gross")) Then
	mc_gross = 0 
 else
	mc_gross = ForSQL(Request.Form("mc_gross"))
 End if
 ' Item_name = Request.Form("item_name")
 Item_number      = ForSQL(Request.Form("item_number"))
 payment_type     = ForSQL(Request.Form("payment_type"))
 Payment_status   = ForSQL(Request.Form("payment_status"))
 ' payment_amount = Request.Form("payment_amount")
 Payment_currency = ForSQL(Request.Form("Payment_currency"))
 payment_date     = ForSQL(Request.Form("payment_date"))
 Txn_type         = ForSQL(Request.Form("txn_type"))
 Txn_id           = ForSQL(Request.Form("txn_id"))
 orderid          = ForSQL(Request.Form("invoice"))
 business         = ForSQL(Request.Form("business") )
 receiver_id      = ForSQL(Request.Form("receiver_id"))
 receiver_email   = ForSQL(Request.Form("receiver_email"))
 custom           = ForSQL(Decode(Request.Form("custom")))
 Payer_id         = ForSQL(Request.Form("payer_id"))
 Payer_email      = ForSQL(Request.Form("payer_email"))
 
 If IsEmpty(orderid) Then orderid = Session("orderid")
		
 Set conn = Server.CreateObject("ADODB.Connection")
 conn.Open EW_DB_CONNECTION_STRING
 
 'Check notification validation
 if (objHttp.status <> 200 ) then
	strSQL = "UPDATE Orders SET payment_status='Error: objHttp.status <> 200', payment_date=#" & Now() & "#" 
	strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	conn.Execute(strSQL)
	
 elseif (objHttp.responseText = "VERIFIED") then
	'check that Payment_status=Completed
	'check that Txn_id has not been previously processed
	'check that Receiver_email is your Primary PayPal email
	'check that Payment_amount/Payment_currency are correct
	'process payment
	
	if (Payment_status="Completed") Then
		strSQL = "UPDATE Orders SET payment_status='Completed', payer_email='"&Payer_email&"', pfirst_name='"&first_name&"',plast_name='"&last_name&"',txn_id='"&Txn_id&"',receiver_email='"&receiver_email&"',"
		strSQL = strSQL & "payment_gross=" & mc_gross & ",payment_type='"& payment_type &"',payment_fee="& mc_fee &",txn_type='"& Txn_type &"', payment_date=#" & Now() &"#," 
		strSQL = strSQL & "pShip_Name='" & address_name & "',pShip_Address='"& address_street &"',pShip_City='"& address_city &"',pShip_Province='"& address_state &"',pShip_Postal='"& address_zip &"',pShip_Country='"& address_country &"'" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
		
		if(custom & "x"<>"x") then
			strSQL1 = " UPDATE Discountcodes SET Discountcodes.used = Yes, Discountcodes.Use_date = Now(), OrderId=" & orderid 
			strSQL1 = strSql1 &  " WHERE (((Discountcodes.DiscountCode)='" & mid(custom,1,5) & "'));"
			conn.execute strSQL1
		end if

	End If
	
	if (Payment_status="Pending") Then
		strSQL = "UPDATE Orders SET payment_status='Pending', payer_email='"&Payer_email&"', pfirst_name='"&first_name&"',plast_name='"&last_name&"',txn_id='"&Txn_id&"',receiver_email='"&receiver_email&"',"
		strSQL = strSQL & "payment_gross=" & mc_gross & ",payment_type='"& payment_type &"',payment_fee="& mc_fee &",txn_type='"& Txn_type &"', payment_date=#" & Now() &"#," 
		strSQL = strSQL & "pShip_Name='" & address_name & "',pShip_Address='"& address_street &"',pShip_City='"& address_city &"',pShip_Province='"& address_state &"',pShip_Postal='"& address_zip &"',pShip_Country='"& address_country &"'" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"   
		
		if(custom & "x"<>"x") then
			strSQL1 = " UPDATE Discountcodes SET Discountcodes.used = Yes, Discountcodes.Use_date = Now(), OrderId=" & orderid 
			strSQL1 = strSql1 &  " WHERE (((Discountcodes.DiscountCode)='" & mid(custom,1,5) & "'));"
			conn.execute strSQL1
		end if
		
	End If	
	
	if (Payment_status="Denied") Then
		strSQL = "UPDATE Orders SET payment_status='Denied', payment_date=#" & Now() &"#" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	End If
	
	if (Payment_status="Refunded") Then
		strSQL = "UPDATE Orders SET payment_status='Refunded', payment_date=#" & Now() &"#" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	End If
	
	if (Payment_status="Reversed") Then
		strSQL = "UPDATE Orders SET payment_status='Reversed', payment_date=#" & Now() &"#" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	End If
	
	if (Payment_status="Voided") Then
		strSQL = "UPDATE Orders SET payment_status='Voided', payment_date=#" & Now() &"#" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	End If
	
	if (Payment_status="Failed") Then
		strSQL = "UPDATE Orders SET payment_status='Failed', payment_date=#" & Now() &"#" 
		strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
		conn.Execute(strSQL)
	End If
	conn.Execute(strSQL)
	
 elseif (objHttp.responseText = "INVALID") then
	'log for manual investigation
	strSQL = "UPDATE Orders SET payment_status='Error: objHttp.responseText - INVALID: discrepancy with what was originally sent', payment_date=#" & Now() & "#" 
	strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	conn.Execute(strSQL)
 else
	'error
	strSQL = "UPDATE Orders SET payment_status='Error: Can not process order', payment_date=#" & Now() & "#" 
	strSQL = strSQL & " WHERE orderid =" & orderid & ";"  
	conn.Execute(strSQL)
 end if
 
 ' WriteToFile "log.txt" , Request.Form  & now & vbcrlf,true 
  
 ' send notification email
 '********************************************************************************
 const EW_SMTPSERVER="smtp.1and1.com"
 const EW_SMTPSERVER_USERNAME="info@vsl3.ca"
 const EW_SMTPSERVER_PASSWORD="web4colon"
 const EW_SMTPSERVER_PORT=25

 sFrEmail = "info@ferring.ca"
 ' sToEmail="info@ferring.ca,systemdesign@ravenshoegroup.com"
 sToEmail = "Vicky.Ball@ferring.com"
 sToEmail = "Antonietta.Pozzebon@ferring.com"
 sToEmail = "CA0-VSL3@ferring.com"
 'sToEmail = "skarkera@ravenshoegroup.com"
 sSubject = "New order submission from VSL3.ca ordered by " & first_name & " " & last_name  
 sbody = getEmailText(orderid,conn)
 sbody = replace(sbody, vbcrlf,"<br>" )
 sbody = replace(sbody, "/r/n","<br>" )
 sTextBody = sBody 
 sHTMLBody = sBody
 sbody = replace(sbody,"                   " & vbcrlf ,"")
 sMail = sBody
 'Call Send_Email(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)
 Call Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)
 
 '********************************************************************************
 
set objHttp = nothing
conn.Close ' Close Connection
 
function WriteToFile(FileName, Contents, Append)
	on error resume next
	if Append = true then
		iMode = 8
	else 
		iMode = 2
	end if
	set oFs = server.createobject("Scripting.FileSystemObject")
	'Added by Ramy for testing 
	set oTextFile = oFs.OpenTextFile( server.mappath("/db/") & "\" & FileName, iMode, True)
	'set oTextFile = oFs.OpenTextFile( server.mappath("/infopaknews/test/") & "\" & FileName, iMode, True)
	oTextFile.Write Contents
	oTextFile.Close
	set oTextFile = nothing
	set oFS = nothing
end function

function getEmailText(InvId,conn)
	
	sBody  = "New Payment updated at VSL3.ca on " & Now &   vbcrlf
	sBody  =sBody & "<table cellspacing='0' cellpadding='3' class='invoice' style='border-style:dashed;border-width:1px;' border='0'>"
	sBody  =sBody & "<tr height='36'>"
	sBody  =sBody & "<th align='left' width='80px' style='padding-left:5px'>Item #</th>"						
	sBody  =sBody & "<th align='center' width='*'>Description</th>"						
	sBody  =sBody & "<th align='center' width='70px'>Quantity</th>"
	sBody  =sBody & "<th align='right' width='70px'>Price</th>"						
	sBody  =sBody & "<th align='right' width='80px' style='padding-right:3px'>Amount</th>"						
	sBody  =sBody & "</tr>"
	strSQL = "SELECT products.ItemNo AS ItemNo,products.Description AS Description,orderdetails.Price AS Price,orderdetails.Quantity AS Quantity "
	strSQL = strSQL & " FROM orders,orderdetails,products WHERE orders.orderId = orderdetails.orderId and products.itemid = orderdetails.productId and orders.orderid = " & InvId & ";"
	Set rst = conn.Execute(strSQL)
	Do While not rst.EOF
		sBody  =sBody & "<tr>"
		sBody  =sBody & "<td align='left'>" & rst("ItemNo") & "</td>"						
		sBody  =sBody & "<td >" & rst("Description") & "</td>"						
		sBody  =sBody & "<td align='center'>" & rst("Quantity") & "</td>"
		sBody  =sBody & "<td align='right'>$" & rst("Price") & "</td>"						
		sBody  =sBody & "<td align='right'>$" & rst("Price") * rst("Quantity") & "</td>"						
		sBody  =sBody & "</tr>"		
		Subtotal = Subtotal + CDbl(rst("Price") * rst("Quantity")) 
		rst.MoveNext						
	Loop
	rst.close
	
	sBody  =sBody & "<tr>"
	sBody  =sBody & "<td align='right' colspan='4'><b>Sub-Total:</b></td>"						
	sBody  =sBody & "<td align='right'><b>$" & Subtotal & "</b></td>"						
	sBody  =sBody & "</tr>"	
	sBody  =sBody & "</table>"
	
	sBody  =sBody & "<br /><div style='font-family: Verdana;font-size: 14px; padding-top:10px;' ><a href='http://www.vsl3.ca/beta/admin/'><i>Please login to see the payment.</i></a></div>"
	
	getEmailText = sBody
End Function
	
Sub Send_Email(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat,sRtn,aAtt)
	Dim objMail, objConfig, sServerVersion, i, sIISVer
	Dim sSmtpServer, iSmtpServerPort
	sServerVersion = Request.ServerVariables("SERVER_SOFTWARE")
	If InStr(sServerVersion, "Microsoft-IIS") > 0 Then
		i = InStr(sServerVersion, "/")
		If i > 0 Then
			sIISVer = Trim(Mid(sServerVersion, i+1))
		End If
	End If
	If sIISVer < "5.0" Then
		' NT using CDONTS
		Set objMail = Server.CreateObject("CDONTS.NewMail")
		objMail.From = sFrEmail
		objMail.To = Replace(sToEmail, ",", ";")
		If sCcEmail <> "" Then
			objMail.Cc = Replace(sCcEmail, ",", ";")
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = Replace(sBccEmail, ",", ";")
		End If
		If LCase(sFormat) = "html" Then
			objMail.BodyFormat = 0 ' 0 means HTML format, 1 means text
			objMail.MailFormat = 0 ' 0 means MIME, 1 means text
		End If
		objMail.Subject = sSubject
		objMail.Body = sMail
		objMail.Send
		Set objMail = Nothing
	Else
		' 2000 / XP / 2003 using CDO
		' Set up Configuration
		Set objConfig = Server.CreateObject("CDO.Configuration")
		objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EW_SMTPSERVER ' cdoSMTPServer
		objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = EW_SMTPSERVER_PORT ' cdoSMTPServerPort
		objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		If EW_SMTPSERVER_USERNAME <> "" And EW_SMTPSERVER_PASSWORD <> "" Then
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic (clear text)
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = EW_SMTPSERVER_USERNAME
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EW_SMTPSERVER_PASSWORD
		End If
		objConfig.Fields.Update
		' Set up Mail
		Set objMail = Server.CreateObject("CDO.Message")
		objMail.From = sFrEmail
		objMail.To = Replace(sToEmail, ",", ";")
		If sCcEmail <> "" Then
			objMail.Cc = Replace(sCcEmail, ",", ";")
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = Replace(sBccEmail, ",", ";")
		End If
		If LCase(sFormat) = "html" Then
			objMail.HtmlBody = sMail
		Else
			objMail.TextBody = sMail
		End If
		
		If (sAtt<>"") Then
			objMail.AddAttachment sAtt
		End If
		
		objMail.Subject = sSubject
		If EW_SMTPSERVER <> "" And LCase(EW_SMTPSERVER) <> "localhost" Then
			Set objMail.Configuration = objConfig ' Use Configuration
			objMail.Send
		Else
			On Error Resume Next
			objMail.Send ' Send without Configuration
			If Err.Number <> 0 Then
				If Hex(Err.Number) = "80040220" Then ' Requires Configuration
					Set objMail.Configuration = objConfig
					Err.Clear
					On Error GoTo 0
					objMail.Send
				Else
					Dim ErrNo, ErrSrc, ErrDesc
					ErrNo = Err.Number
					ErrSrc = Err.Source
					ErrDesc = Err.Description
					On Error GoTo 0
					Err.Raise ErrNo, ErrSrc, ErrDesc
				End If
			else
				' response.redirect(sRtn)
			End If
		End If
		Set objMail = Nothing
		Set objConfig = Nothing
		' response.redirect(sRtn)
	End If
End Sub

Sub Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat,sRtn,aAtt)
	Dim objMail
Set objMail = Server.CreateObject("CDO.Message")
Set objConfig = Server.CreateObject("CDO.Configuration")


objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.1and1.com"
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vsl3.ca" 'Enter YOUR E-MAIL ADDRESS here
objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "web4colon" 'Enter the PASSWORD for your email address
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

Function Decode(sIn)
dim x, y, abfrom, abto
Decode="": ABFrom = ""
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
If y = 0 then
Decode = Decode & Mid(sin, x, 1)
Else
Decode = Decode & Mid(abfrom, y, 1)
End If
Next
End Function


Function ForSQL(strString)  
  ForSQL = Replace(strString, "'", "''")  
End Function

 %>
 