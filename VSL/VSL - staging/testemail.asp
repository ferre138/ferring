<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%

const EW_SMTPSERVER="smtp.1and1.com"
 const EW_SMTPSERVER_USERNAME="info@vsl3.ca"
 const EW_SMTPSERVER_PASSWORD="web4colon"
 const EW_SMTPSERVER_PORT=25

 sFrEmail = "info@ferring.ca"
 ' sToEmail="info@ferring.ca,systemdesign@ravenshoegroup.com"
 'sToEmail = "Vicky.Ball@ferring.com"
 'sToEmail = "Antonietta.Pozzebon@ferring.com"
 'sToEmail = "CA0-VSL3@ferring.com"
 'sToEmail = "skarkera@ravenshoegroup.com"
 sToEmail = "bcostoff@ravenshoegroup.com"
 sSubject = "New order submission from VSL3.ca ordered by " 
 sbody = "test message from site"
 sbody = replace(sbody, vbcrlf,"<br>" )
 sbody = replace(sbody, "/r/n","<br>" )
 sTextBody = sBody 
 sHTMLBody = sBody
 sbody = replace(sbody,"                   " & vbcrlf ,"")
 sMail = sBody
 'Call Send_Email(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)
 Call Email_1and1(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, "html",rtnUrl,sAtt)




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
%>

