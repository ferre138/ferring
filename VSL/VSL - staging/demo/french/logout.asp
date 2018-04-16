<!--#include file="ewcfg60.asp"-->
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
Dim bValidate
bValidate = True
Dim sLastUrl, sUsername
sUsername = Security.CurrentUserName

' Call User LoggingOut event
bValidate = User_LoggingOut(sUsername)
If Not bValidate Then
	sLastUrl = Security.LastUrl
	If sLastUrl = "" Then sLastUrl = "default.asp"
	Call Page_Terminate(sLastUrl) ' Go to last accessed url
Else
	If Request.Cookies(EW_PROJECT_NAME)("autologin") = "" Then  ' Not autologin
		Response.Cookies(EW_PROJECT_NAME)("username") = "" ' clear user name cookie
	End If
	Response.Cookies(EW_PROJECT_NAME)("password") = "" ' clear password cookie
	Response.Cookies(EW_PROJECT_NAME)("lasturl") = "" ' clear last url

	' Clear session
	Session.Abandon

	' Call User LoggedOut event
	Call User_LoggedOut(sUsername)
	Call Page_Terminate("login.asp") ' Go to login page
End If
%>
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

' Page Load event
Sub Page_Load()

'***Response.Write "Page Load"
End Sub

' Page Unload event
Sub Page_Unload()

'***Response.Write "Page Unload"
End Sub
%>
<%

' User Logging Out event
Function User_LoggingOut(usr)
	On Error Resume Next

	' Enter your code here
	' To cancel, set return value to False

	User_LoggingOut = True
End Function

' User Logged Out event
Sub User_LoggedOut(usr)

	' Response.Write "User Logged Out"
End Sub
%>
