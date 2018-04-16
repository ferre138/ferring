 <% 
'Declare variables 
Dim sMsg 
Dim sTo 
Dim sFrom 
Dim sBody 
Dim sTextBody 
Dim sHTMLBody 


 
'Get data from previous page

'sTo = request("recipient")

sPage=Request("RequestPage")
sFrom = Request("name")
response.write sfrom
'if(sFrom & "x"="x") then 
'	response.write "<h2>Please enable Javascript on your browser and Try again</h2>"
'	response.end
'end if

if(sPage="F") then
	sBody =  " French " 
else
	sBody =  " English " 
end if

sBody  =sBody &  "Testimonial submission from vsl3.ca at " & Now  &   vbcrlf
sBody  =sBody & "------------------------------------------------"  &   vbcrlf
sBody  =sBody & "Name: " & Request("name")  & vbcrlf
sBody  =sBody & "City: " & Request("city")  & vbcrlf
sBody  =sBody & "Publish Online: " & Request("publish")   & vbcrlf
sBody  =sBody & "Show Name Online: " & Request("showname")   & vbcrlf
sBody  =sBody & "Show Initial Online: " & Request("showinit")   & vbcrlf
sBody  =sBody & "Show City Online: " & Request("showcity")   & vbcrlf


sBody  =sBody & "------------------------------------------------"    & vbcrlf
sBody  =sBody & "Testimonial: " & Request("comments")& vbcrlf
sBody  =sBody & vbcrlf
sBody  =sBody & "----------------End of message------------------"    & vbcrlf

sTextBody = sBody 
sHTMLBody = sBody
sbody= replace(sbody,"                   " & vbcrlf ,"")

'response.write replace(sbody,vbcrlf ,"<br>")
'response.end


strTo ="nhp@ferring.com" 'Make sure the From field has no spaces. 
'strTo ="bcostoff@ravenshoegroup.com" 'Make sure the From field has no spaces. 
strFrom = "info@vsl3.ca" 
strSubject = "Testimonials submission from " &  Request("name")
strBody = sBody


' Create an instance of the NewMail object. 
'Set objCDOMail = Server.CreateObject("CDO.Message") 

'objCDOMail.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
'Name or IP of remote SMTP server
'objCDOMail.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
'="mail.pico-salax.ca"
'Server port
'objCDOMail.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpserverport") _
'=25 
'objCDOMail.Configuration.Fields.Update

Set Mail = Server.CreateObject("SMTPsvg.Mailer") 'create an Asp mail component.
Mail.FromName   = "VSL3 web info"
Mail.FromAddress= strFrom
Mail.RemoteHost = "mrelay.perfora.net" ' The mail server you have to use with Asp Mail
Mail.AddRecipient "Medical Information", strTo
Mail.Subject    = strSubject
Mail.BodyText   = sBody
    
' Set the properties of the object 
'objCDOMail.Sender = StrFrom 
'objCDOMail.To = strTo 
'objCDOMail.Subject = strSubject 
'objCDOMail.TextBody = strBody 
   strErr = ""
   bSuccess = False
   On Error Resume Next ' catch errors
if Mail.SendMail then

	Response.Redirect ("thankyou.html")

else
	strErr = Err.Description%>
	<h3>Error occurred: <% = strErr %></h3>

<% End If %>
