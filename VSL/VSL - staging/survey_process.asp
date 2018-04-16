<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->

<%
  Dim Email
  Set Email = New cEmail
  Email.Load("txt/survey.txt")
  Email.ReplaceSender(EW_SENDER_EMAIL) ' Replace Sender
  Email.ReplaceRecipient("bcostoff@ravenshoegroup.com") ' Replace Recipient
  Email.ReplaceSubject("New Survey submission on " & now())
  Email.ReplaceContent "<!--$date-->", Now()
  Email.ReplaceContent "<!--$customer-->", Request.Form("customer")          
  Email.ReplaceContent "<!--$hearus-->", Request.Form("hearus")
  Email.ReplaceContent "<!--$specify-->", Request.Form("specify")
  Email.ReplaceContent "<!--$recommend-->", Request.Form("recommend")
  Email.ReplaceContent "<!--$what_do_you_like-->", Request.Form("what_do_you_like")
  Email.ReplaceContent "<!--$for_even_higher_rate-->", Request.Form("for_even_higher_rate")
  Email.ReplaceContent "<!--$for_higher_rate-->", Request.Form("for_higher_rate")
  Email.ReplaceContent "<br/>", vbcrlf
  Email.Send()
  Set Email = Nothing
  Response.Redirect "http://www.vsl3.ca/survey.asp"
%>  