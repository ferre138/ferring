<!--#include file="ewcfg60.asp"-->
<%


    


const EW_SMTPSERVER="smtp.1and1.com"
const EW_SMTPSERVER_USERNAME="info@vsl3.ca"
const EW_SMTPSERVER_PASSWORD="web4colon"
const EW_SMTPSERVER_PORT=25
   
sUserName = "Beau"
sPassword = "Beau"
sFrEmail = "info@ferring.ca"
sToEmail = "bcostoff@ravenshoegroup.com"
sSubject = "Password Recovery From VSL3.ca " 
sbody = "Dear Sir/Madam, \n\n Please see below for the requested information: \n\n User Name: " & sUserName & "\n Password: " & sPassword & "\n \n Please feel free to contact us in case of further queries. \n \n Best Regards, \n Support"
sbody = replace(sbody, "\n", vbCrLf)
sbody = replace(sbody, vbcrlf,"<br>" )
sbody = replace(sbody, "/r/n","<br>" )
sTextBody = sBody 
sHTMLBody = sBody
sbody = replace(sbody,"                   " & vbcrlf ,"")
sMail = sBody
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
 objMail.Subject = sSubject 'Enter a SUBJECT
 
 objMail.TextBody = sMail 'Enter the BODY of the message
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

