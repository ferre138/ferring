<%
Const EW_PAGE_ID = "default"
%>
<!--#include file="..\ewcfg60.asp"-->
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
If Security.IsLoggedIn() Then
	Call Page_Terminate("vslCart.asp") ' Exit and go to default page
End If
If Security.IsLoggedIn() Then
	Call Page_Terminate("Shippinglist.asp")
End If
If Security.IsLoggedIn() Then
	Call Page_Terminate("Productslist.asp")
End If
If Security.IsLoggedIn() Then
	Response.Write "You do not have the right permission to view the page"
%>
<br>
<a href="logout.asp">Back to Login Page</a>
<%
Else
	Call Page_Terminate("login.asp") ' Exit and go to login page
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
%>
