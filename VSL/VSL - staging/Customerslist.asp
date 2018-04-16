<%
Const EW_PAGE_ID = "list"
Const EW_TABLE_NAME = "Customers"
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
If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
If Not Security.IsLoggedIn() Then
	Call Security.SaveLastUrl()
	Call Page_Terminate("login.asp")
End If
If Security.IsLoggedIn() And Security.CurrentUserID = "" Then
	Session(EW_SESSION_MESSAGE) = "You do not have the right permission to view the page"
	Call Page_Terminate("login.asp")
End If
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
Customers.Export = Request.QueryString("export") ' Get export parameter
sExport = Customers.Export ' Get export parameter, used in header
sExportFile = Customers.TableVar ' Get export file, used in header
%>
<%

' Paging variables
Dim Pager, PagerItem ' Pager
Dim nDisplayRecs ' Number of display records
Dim nRecRange ' Record display range
Dim nStartRec, nStopRec, nTotalRecs
nStartRec = 0 ' Start record index
nStopRec = 0 ' Stop record index
nTotalRecs = 0 ' Total number of records
nDisplayRecs = 20
nRecRange = 10
Dim i
Dim nRecCount
nRecCount = 0 ' Record count
Dim RowCnt, RowIndex, OptionCnt

' Sort
Dim sSortOrder

' Search filters
Dim sSrchAdvanced, sSrchBasic, sSrchWhere, sFilter
sSrchAdvanced = "" ' Advanced search filter
sSrchBasic = "" ' Basic search filter
sSrchWhere = "" ' Search where clause
sFilter = ""
Dim bEditRow, nEditRowCnt ' Edit row

' Master/Detail
Dim sDbMasterFilter, sDbDetailFilter
sDbMasterFilter = "" ' Master filter
sDbDetailFilter = "" ' Detail filter
Dim sSqlMaster
sSqlMaster = "" ' Sql for master record

' Handle reset command
ResetCmd()

' Get basic search criteria
sSrchBasic = BasicSearchWhere()

' Build search criteria
If sSrchAdvanced <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchAdvanced & ")"
End If
If sSrchBasic <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchBasic & ")"
End If

' Save search criteria
If sSrchWhere <> "" Then
	If sSrchBasic = "" Then Call ResetBasicSearchParms()
	Customers.SearchWhere = sSrchWhere ' Save to Session
	nStartRec = 1 ' Reset start record counter
	Customers.StartRecordNumber = nStartRec
Else
	Call RestoreSearchParms()
End If

' Build filter
sFilter = ""
If Security.CurrentUserID <> "" And Not Security.IsAdmin Then ' Non system admin
	sFilter = Customers.AddUserIDFilter(sFilter, Security.CurrentUserID) ' Add user id filter
End If
If sDbDetailFilter <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sDbDetailFilter & ")"
End If
If sSrchWhere <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sSrchWhere & ")"
End If

' Set up filter in Session
Customers.SessionWhere = sFilter
Customers.CurrentFilter = ""

' Set Up Sorting Order
SetUpSortOrder()

' Set Return Url
Customers.ReturnUrl = "Customerslist.asp"
%>
<!--#include file="header.asp"-->
<% If Customers.Export = "" Then %>
<script type="text/javascript">
<!--
var EW_PAGE_ID = "list"; // Page id
//-->
</script>
<script type="text/javascript">
<!--
var firstrowoffset = 1; // First data row start at
var lastrowoffset = 0; // Last data row end at
var EW_LIST_TABLE_NAME = 'ewlistmain'; // Table name for list page
var rowclass = 'ewTableRow'; // Row class
var rowaltclass = 'ewTableAltRow'; // Row alternate class
var rowmoverclass = 'ewTableHighlightRow'; // Row mouse over class
var rowselectedclass = 'ewTableSelectRow'; // Row selected class
var roweditclass = 'ewTableEditRow'; // Row edit class
//-->
</script>
<script type="text/javascript">
<!--
var ew_DHTMLEditors = [];
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
// To include another .js script, use:
// ew_ClientScriptInclude("my_javascript.js"); 
//-->
</script>
<% End If %>
<% If Customers.Export = "" Then %>
<% End If %>
<%

' Load recordset
Dim rs
Set rs = LoadRecordset()
nTotalRecs = rs.RecordCount
nStartRec = 1
If nDisplayRecs <= 0 Then ' Display all records
	nDisplayRecs = nTotalRecs
End If
If Not (EW_EXPORT_ALL And Customers.Export <> "") Then
	SetUpStartRec() ' Set up start record position
End If
%>
<p><span class="aspmaker" style="white-space: nowrap;">TABLE: Customers
</span></p>
<% If Customers.Export = "" Then %>
<% If Security.IsLoggedIn() Then %>
<form name="fCustomerslistsrch" id="fCustomerslistsrch" action="Customerslist.asp" >
<table class="ewBasicSearch">
	<tr>
		<td><span class="aspmaker">
			<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" size="20" value="<%= Customers.BasicSearchKeyword %>">
			<input type="Submit" name="Submit" id="Submit" value="Search (*)">&nbsp;
			<a href="Customerslist.asp?cmd=reset">Show all</a>&nbsp;
		</span></td>
	</tr>
	<tr>
	<td><span class="aspmaker"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="" <% If Customers.BasicSearchType = "" Then %>checked<% End If %>>Exact phrase&nbsp;&nbsp;<input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND" <% If Customers.BasicSearchType = "AND" Then %>checked<% End If %>>All words&nbsp;&nbsp;<input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" id="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR" <% If Customers.BasicSearchType = "OR" Then %>checked<% End If %>>Any word</span></td>
	</tr>
</table>
</form>
<% End If %>
<% End If %>
<%
If Session(EW_SESSION_MESSAGE) <> "" Then
%>
<p><span class="ewmsg"><%= Session(EW_SESSION_MESSAGE) %></span></p>
<%
	Session(EW_SESSION_MESSAGE) = "" ' Clear message
End If
%>
<% If Customers.Export = "" Then %>
<form action="Customerslist.asp" name="ewpagerform" id="ewpagerform">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td nowrap>
<span class="aspmaker">
<% If Not IsObject(Pager) Then Set Pager = ew_NewNumericPager(nStartRec, nDisplayRecs, nTotalRecs, nRecRange) %>
<% If Pager.RecordCount > 0 Then %>
	<% If Pager.FirstButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.FirstButton.Start %>"><b>First</b></a>&nbsp;
	<% End If %>
	<% If Pager.PrevButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.PrevButton.Start %>"><b>Previous</b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="Customerslist.asp?start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Pager.NextButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.NextButton.Start %>"><b>Next</b></a>&nbsp;
	<% End If %>
	<% If Pager.LastButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.LastButton.Start %>"><b>Last</b></a>&nbsp;
	<% End If %>
	<% If Pager.ButtonCount > 0 Then %><br><%	End If %>
	Records <%= Pager.FromIndex %> to <%= Pager.ToIndex %> of <%= Pager.RecordCount %>
<% Else %>	
	<% If sSrchWhere = "0=101" Then %>
	Please enter search criteria
	<% Else %>
	No records found
	<% End If %>
<% End If %>
</span>
		</td>
	</tr>
</table>
</form>
<% End If %>
<form method="post" name="fCustomerslist" id="fCustomerslist">
<% If Customers.Export = "" Then %>
<table>
	<tr><td><span class="aspmaker">
<% If Security.IsLoggedIn() Then %>
<a href="Customersadd.asp">Add</a>&nbsp;&nbsp;
<% End If %>
	</span></td></tr>
</table>
<% End If %>
<% If nTotalRecs > 0 Then %>
<table id="ewlistmain" class="ewTable">
<%
	OptionCnt = 0
If Security.IsLoggedIn() Then
	OptionCnt = OptionCnt + 1 ' view
End If
	OptionCnt = OptionCnt + 1 ' edit
If Security.IsLoggedIn() Then
	OptionCnt = OptionCnt + 1 ' delete
End If
If Security.IsLoggedIn() Then
	OptionCnt = OptionCnt + 1 '  detail
End If
%>
	<!-- Table header -->
	<tr class="ewTableHeader">
		<td valign="top">
<% If Customers.Export <> "" Then %>
First Name
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("Inv_FirstName") %>&ordertype=<%= Customers.Inv_FirstName.ReverseSort %>">First Name&nbsp;(*)<% If Customers.Inv_FirstName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_FirstName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Last Name
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("Inv_LastName") %>&ordertype=<%= Customers.Inv_LastName.ReverseSort %>">Last Name&nbsp;(*)<% If Customers.Inv_LastName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_LastName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Billing Address
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("Inv_Address") %>&ordertype=<%= Customers.Inv_Address.ReverseSort %>">Billing Address&nbsp;(*)<% If Customers.Inv_Address.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_Address.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Address 2
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("Inv_Address2") %>&ordertype=<%= Customers.Inv_Address2.ReverseSort %>">Address 2&nbsp;(*)<% If Customers.Inv_Address2.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.Inv_Address2.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
City
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("inv_City") %>&ordertype=<%= Customers.inv_City.ReverseSort %>">City&nbsp;(*)<% If Customers.inv_City.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_City.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Province
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("inv_Province") %>&ordertype=<%= Customers.inv_Province.ReverseSort %>">Province<% If Customers.inv_Province.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_Province.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Email Address
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("inv_EmailAddress") %>&ordertype=<%= Customers.inv_EmailAddress.ReverseSort %>">Email Address&nbsp;(*)<% If Customers.inv_EmailAddress.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_EmailAddress.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
Fax
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("inv_Fax") %>&ordertype=<%= Customers.inv_Fax.ReverseSort %>">Fax&nbsp;(*)<% If Customers.inv_Fax.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.inv_Fax.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
		<td valign="top">
<% If Customers.Export <> "" Then %>
User Name
<% Else %>
	<a href="Customerslist.asp?order=<%= Server.URLEncode("UserName") %>&ordertype=<%= Customers.UserName.ReverseSort %>">User Name&nbsp;(*)<% If Customers.UserName.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Customers.UserName.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</td>
<% If Customers.Export = "" Then %>
<% If Security.IsLoggedIn() Then %>
<td nowrap>&nbsp;</td>
<% End If %>
<td nowrap>&nbsp;</td>
<% If Security.IsLoggedIn() Then %>
<td nowrap>&nbsp;</td>
<% End If %>
<% If Security.IsLoggedIn() Then %>
<td nowrap>&nbsp;</td>
<% End If %>
<% End If %>
	</tr>
<%
If (EW_EXPORT_ALL And Customers.Export <> "") Then
	nStopRec = nTotalRecs
Else
	nStopRec = nStartRec + nDisplayRecs - 1 ' Set the last record to display
End If

' Move to first record directly for performance reason
nRecCount = nStartRec - 1
If Not rs.Eof Then
	rs.MoveFirst
	rs.Move nStartRec - 1
End If
RowCnt = 0
Do While (Not rs.Eof) And (nRecCount < nStopRec)
	nRecCount = nRecCount + 1
	If CLng(nRecCount) >= CLng(nStartRec) Then
		RowCnt = RowCnt + 1

	' Init row class and style
	Customers.CssClass = "ewTableRow"
	Customers.CssStyle = ""

	' Init row event
	Customers.RowClientEvents = "onmouseover='ew_MouseOver(this);' onmouseout='ew_MouseOut(this);' onclick='ew_Click(this);'"

	' Display alternate color for rows
	If RowCnt Mod 2 = 0 Then
		Customers.CssClass = "ewTableAltRow"
	End If
	Call LoadRowValues(rs) ' Load row values
	Customers.RowType = EW_ROWTYPE_VIEW ' Render view
	Call RenderRow()
%>
	<!-- Table body -->
	<tr<%= Customers.DisplayAttributes %>>
		<!-- Inv_FirstName -->
		<td<%= Customers.Inv_FirstName.CellAttributes %>>
<div<%= Customers.Inv_FirstName.ViewAttributes %>><%= Customers.Inv_FirstName.ViewValue %></div>
</td>
		<!-- Inv_LastName -->
		<td<%= Customers.Inv_LastName.CellAttributes %>>
<div<%= Customers.Inv_LastName.ViewAttributes %>><%= Customers.Inv_LastName.ViewValue %></div>
</td>
		<!-- Inv_Address -->
		<td<%= Customers.Inv_Address.CellAttributes %>>
<div<%= Customers.Inv_Address.ViewAttributes %>><%= Customers.Inv_Address.ViewValue %></div>
</td>
		<!-- Inv_Address2 -->
		<td<%= Customers.Inv_Address2.CellAttributes %>>
<div<%= Customers.Inv_Address2.ViewAttributes %>><%= Customers.Inv_Address2.ViewValue %></div>
</td>
		<!-- inv_City -->
		<td<%= Customers.inv_City.CellAttributes %>>
<div<%= Customers.inv_City.ViewAttributes %>><%= Customers.inv_City.ViewValue %></div>
</td>
		<!-- inv_Province -->
		<td<%= Customers.inv_Province.CellAttributes %>>
<div<%= Customers.inv_Province.ViewAttributes %>><%= Customers.inv_Province.ViewValue %></div>
</td>
		<!-- inv_EmailAddress -->
		<td<%= Customers.inv_EmailAddress.CellAttributes %>>
<div<%= Customers.inv_EmailAddress.ViewAttributes %>><%= Customers.inv_EmailAddress.ViewValue %></div>
</td>
		<!-- inv_Fax -->
		<td<%= Customers.inv_Fax.CellAttributes %>>
<div<%= Customers.inv_Fax.ViewAttributes %>><%= Customers.inv_Fax.ViewValue %></div>
</td>
		<!-- UserName -->
		<td<%= Customers.UserName.CellAttributes %>>
<div<%= Customers.UserName.ViewAttributes %>><%= Customers.UserName.ViewValue %></div>
</td>
<% If Customers.Export = "" Then %>
<% If Security.IsLoggedIn() Then %>
<td nowrap><span class="aspmaker"><% If ShowOptionLink() Then %>
<a href="<%= Customers.ViewUrl %>"><img src='images/view.gif' alt='View' title='View' width='16' height='16' border='0'></a>
<% End If %></span></td>
<% End If %>
<td nowrap><span class="aspmaker">
<a href="<%= Customers.EditUrl %>"><img src='images/edit.gif' alt='Edit' title='Edit' width='16' height='16' border='0'></a>
</span></td>
<% If Security.IsLoggedIn() Then %>
<td nowrap><span class="aspmaker"><% If ShowOptionLink() Then %>
<a href="<%= Customers.DeleteUrl %>"><img src='images/delete.gif' alt='Delete' title='Delete' width='16' height='16' border='0'></a>
<% End If %></span></td>
<% End If %>
<% If Security.IsLoggedIn() Then %>
<td nowrap><span class="aspmaker"><% If ShowOptionLink() Then %>
<a href="Shippinglist.asp?<%= EW_TABLE_SHOW_MASTER %>=Customers&CustomerID=<%= Server.URLEncode(Customers.CustomerID.CurrentValue&"") %>">Shipping<img src='images/detail.gif' alt='Details' title='Details' width='16' height='16' border='0'></a>
<% End If %></span></td>
<% End If %>
<% End If %>
	</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If Customers.Export = "" Then %>
<table>
	<tr><td><span class="aspmaker">
<% If Security.IsLoggedIn() Then %>
<a href="Customersadd.asp">Add</a>&nbsp;&nbsp;
<% End If %>
	</span></td></tr>
</table>
<% End If %>
<% End If %>
</form>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
%>
<% If Customers.Export = "" Then %>
<form action="Customerslist.asp" name="ewpagerform" id="ewpagerform">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td nowrap>
<span class="aspmaker">
<% If Not IsObject(Pager) Then Set Pager = ew_NewNumericPager(nStartRec, nDisplayRecs, nTotalRecs, nRecRange) %>
<% If Pager.RecordCount > 0 Then %>
	<% If Pager.FirstButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.FirstButton.Start %>"><b>First</b></a>&nbsp;
	<% End If %>
	<% If Pager.PrevButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.PrevButton.Start %>"><b>Previous</b></a>&nbsp;
	<% End If %>
	<% For Each PagerItem In Pager.Items %>
		<% If PagerItem.Enabled Then %><a href="Customerslist.asp?start=<%= PagerItem.Start %>"><% End If %><b><%= PagerItem.Text %></b><% If PagerItem.Enabled Then %></a><% End If %>&nbsp;
	<% Next %>
	<% If Pager.NextButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.NextButton.Start %>"><b>Next</b></a>&nbsp;
	<% End If %>
	<% If Pager.LastButton.Enabled Then %>
	<a href="Customerslist.asp?start=<%= Pager.LastButton.Start %>"><b>Last</b></a>&nbsp;
	<% End If %>
	<% If Pager.ButtonCount > 0 Then %><br><%	End If %>
	Records <%= Pager.FromIndex %> to <%= Pager.ToIndex %> of <%= Pager.RecordCount %>
<% Else %>	
	<% If sSrchWhere = "0=101" Then %>
	Please enter search criteria
	<% Else %>
	No records found
	<% End If %>
<% End If %>
</span>
		</td>
	</tr>
</table>
</form>
<% End If %>
<% If Customers.Export = "" Then %>
<% End If %>
<!--#include file="footer.asp"-->
<% If Customers.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
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
	Set Customers = Nothing

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

' Return Basic Search sql
Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = ew_AdjustSql(Keyword)
	BasicSearchSQL = ""
	BasicSearchSQL = BasicSearchSQL & "[Inv_FirstName] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Inv_LastName] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Inv_Address] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Inv_Address2] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_City] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_Province] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_PostalCode] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_Country] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_PhoneNumber] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_EmailAddress] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[inv_Fax] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Notes] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[UserName] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[passwrd] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

' Return Basic Search Where based on search keyword and type
Function BasicSearchWhere()
	Dim sSearchStr, sSearchKeyword, sSearchType
	Dim sSearch, arKeyword, sKeyword
	sSearchStr = ""
	sSearchKeyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)
	sSearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)
	If sSearchKeyword <> "" Then
		sSearch = Trim(sSearchKeyword)
		If sSearchType <> "" Then
			While InStr(sSearch, "  ") > 0
				sSearch = Replace(sSearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(sSearch), " ")
			For Each sKeyword In arKeyword
				If sSearchStr <> "" Then sSearchStr = sSearchStr & " " & sSearchType & " "
				sSearchStr = sSearchStr & "(" & BasicSearchSQL(sKeyword) & ")"
			Next
		Else
			sSearchStr = BasicSearchSQL(sSearch)
		End If
	End If
	If sSearchKeyword <> "" then
		Customers.BasicSearchKeyword = sSearchKeyword
		Customers.BasicSearchType = sSearchType
	End If
	BasicSearchWhere = sSearchStr
End Function

' Clear all search parameters
Sub ResetSearchParms()

	' Clear search where
	sSrchWhere = ""
	Customers.SearchWhere = sSrchWhere

	' Clear basic search parameters
	Call ResetBasicSearchParms()
End Sub

' Clear all basic search parameters
Sub ResetBasicSearchParms()

	' Clear basic search parameters
	Customers.BasicSearchKeyword = ""
	Customers.BasicSearchType = ""
End Sub

' Restore all search parameters
Sub RestoreSearchParms()
	sSrchWhere = Customers.SearchWhere
End Sub

' Set up Sort parameters based on Sort Links clicked
Sub SetUpSortOrder()
	Dim sOrderBy
	Dim sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		Customers.CurrentOrder = Request.QueryString("order")
		Customers.CurrentOrderType = Request.QueryString("ordertype")

		' Field Inv_FirstName
		Call Customers.UpdateSort(Customers.Inv_FirstName)

		' Field Inv_LastName
		Call Customers.UpdateSort(Customers.Inv_LastName)

		' Field Inv_Address
		Call Customers.UpdateSort(Customers.Inv_Address)

		' Field Inv_Address2
		Call Customers.UpdateSort(Customers.Inv_Address2)

		' Field inv_City
		Call Customers.UpdateSort(Customers.inv_City)

		' Field inv_Province
		Call Customers.UpdateSort(Customers.inv_Province)

		' Field inv_EmailAddress
		Call Customers.UpdateSort(Customers.inv_EmailAddress)

		' Field inv_Fax
		Call Customers.UpdateSort(Customers.inv_Fax)

		' Field UserName
		Call Customers.UpdateSort(Customers.UserName)
		Customers.StartRecordNumber = 1 ' Reset start position
	End If
	sOrderBy = Customers.SessionOrderBy ' Get order by from Session
	If sOrderBy = "" Then
		If Customers.SqlOrderBy <> "" Then
			sOrderBy = Customers.SqlOrderBy
			Customers.SessionOrderBy = sOrderBy
		End If
	End If
End Sub

' Reset command based on querystring parameter cmd=
' - RESET: reset search parameters
' - RESETALL: reset search & master/detail parameters
' - RESETSORT: reset sort parameters
Sub ResetCmd()
	Dim sCmd

	' Get reset cmd
	If Request.QueryString("cmd").Count > 0 Then
		sCmd = Request.QueryString("cmd")

		' Reset search criteria
		If LCase(sCmd) = "reset" Or LCase(sCmd) = "resetall" Then
			Call ResetSearchParms()
		End If

		' Reset Sort Criteria
		If LCase(sCmd) = "resetsort" Then
			Dim sOrderBy
			sOrderBy = ""
			Customers.SessionOrderBy = sOrderBy
			Customers.Inv_FirstName.Sort = ""
			Customers.Inv_LastName.Sort = ""
			Customers.Inv_Address.Sort = ""
			Customers.Inv_Address2.Sort = ""
			Customers.inv_City.Sort = ""
			Customers.inv_Province.Sort = ""
			Customers.inv_EmailAddress.Sort = ""
			Customers.inv_Fax.Sort = ""
			Customers.UserName.Sort = ""
		End If

		' Reset start position
		nStartRec = 1
		Customers.StartRecordNumber = nStartRec
	End If
End Sub
%>
<%

' Set up Starting Record parameters based on Pager Navigation
Sub SetUpStartRec()
	Dim nPageNo

	' Exit if nDisplayRecs = 0
	If nDisplayRecs = 0 Then Exit Sub

	' Check for a START parameter
	If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
		nStartRec = Request.QueryString(EW_TABLE_START_REC)
		Customers.StartRecordNumber = nStartRec
	ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
		nPageNo = Request.QueryString(EW_TABLE_PAGE_NO)
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Customers.StartRecordNumber = nStartRec
		Else
			nStartRec = Customers.StartRecordNumber
		End If
	Else
		nStartRec = Customers.StartRecordNumber
	End If

	' Check if correct start record counter
	If Not IsNumeric(nStartRec) Or nStartRec = "" Then ' Avoid invalid start record counter
		nStartRec = 1 ' Reset start record counter
		Customers.StartRecordNumber = nStartRec
	ElseIf CLng(nStartRec) > CLng(nTotalRecs) Then ' Avoid starting record > total records
		nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to last page first record
		Customers.StartRecordNumber = nStartRec
	ElseIf (nStartRec-1) Mod nDisplayRecs <> 0 Then
		nStartRec = ((nStartRec-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to page boundary
		Customers.StartRecordNumber = nStartRec
	End If
End Sub
%>
<%

' Load recordset
Function LoadRecordset()

	' Call Recordset Selecting event
	Call Customers.Recordset_Selecting(Customers.CurrentFilter)

	' Load list page sql
	Dim sSql
	sSql = Customers.ListSQL

	' Response.Write sSql ' Uncomment to show SQL for debugging
	' Load recordset

	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = EW_CURSORLOCATION
	rs.Open sSql, conn, 1, 2

	' Call Recordset Selected event
	Call Customers.Recordset_Selected(rs)
	Set LoadRecordset = rs
End Function
%>
<%

' Load row based on key values
Function LoadRow()
	Dim rs, sSql, sFilter
	sFilter = Customers.SqlKeyFilter
	If Not IsNumeric(Customers.CustomerID.CurrentValue) Then
		LoadRow = False ' Invalid key, exit
		Exit Function
	End If
	sFilter = Replace(sFilter, "@CustomerID@", ew_AdjustSql(Customers.CustomerID.CurrentValue)) ' Replace key value
	If Security.CurrentUserID <> "" And Not Security.IsAdmin Then ' Non system admin
		sFilter = Customers.AddUserIDFilter(sFilter, Security.CurrentUserID) ' Add user id filter
	End If

	' Call Row Selecting event
	Call Customers.Row_Selecting(sFilter)

	' Load sql based on filter
	Customers.CurrentFilter = sFilter
	sSql = Customers.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadRow = False
	Else
		LoadRow = True
		rs.MoveFirst
		Call LoadRowValues(rs) ' Load row values

		' Call Row Selected event
		Call Customers.Row_Selected(rs)
	End If
	rs.Close
	Set rs = Nothing
End Function

' Load row values from recordset
Sub LoadRowValues(rs)
	Customers.CustomerID.DbValue = rs("CustomerID")
	Customers.Inv_FirstName.DbValue = rs("Inv_FirstName")
	Customers.Inv_LastName.DbValue = rs("Inv_LastName")
	Customers.Inv_Address.DbValue = rs("Inv_Address")
	Customers.Inv_Address2.DbValue = rs("Inv_Address2")
	Customers.inv_City.DbValue = rs("inv_City")
	Customers.inv_Province.DbValue = rs("inv_Province")
	Customers.inv_PostalCode.DbValue = rs("inv_PostalCode")
	Customers.inv_Country.DbValue = rs("inv_Country")
	Customers.inv_PhoneNumber.DbValue = rs("inv_PhoneNumber")
	Customers.inv_EmailAddress.DbValue = rs("inv_EmailAddress")
	Customers.inv_Fax.DbValue = rs("inv_Fax")
	Customers.Notes.DbValue = rs("Notes")
	Customers.UserName.DbValue = rs("UserName")
	Customers.passwrd.DbValue = rs("passwrd")
End Sub
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

	' inv_EmailAddress
	Customers.inv_EmailAddress.CellCssStyle = ""
	Customers.inv_EmailAddress.CellCssClass = ""

	' inv_Fax
	Customers.inv_Fax.CellCssStyle = ""
	Customers.inv_Fax.CellCssClass = ""

	' UserName
	Customers.UserName.CellCssStyle = ""
	Customers.UserName.CellCssClass = ""
	If Customers.RowType = EW_ROWTYPE_VIEW Then ' View row

		' Inv_FirstName
		Customers.Inv_FirstName.ViewValue = Customers.Inv_FirstName.CurrentValue
		Customers.Inv_FirstName.CssStyle = ""
		Customers.Inv_FirstName.CssClass = ""
		Customers.Inv_FirstName.ViewCustomAttributes = ""

		' Inv_LastName
		Customers.Inv_LastName.ViewValue = Customers.Inv_LastName.CurrentValue
		Customers.Inv_LastName.CssStyle = ""
		Customers.Inv_LastName.CssClass = ""
		Customers.Inv_LastName.ViewCustomAttributes = ""

		' Inv_Address
		Customers.Inv_Address.ViewValue = Customers.Inv_Address.CurrentValue
		Customers.Inv_Address.CssStyle = ""
		Customers.Inv_Address.CssClass = ""
		Customers.Inv_Address.ViewCustomAttributes = ""

		' Inv_Address2
		Customers.Inv_Address2.ViewValue = Customers.Inv_Address2.CurrentValue
		Customers.Inv_Address2.CssStyle = ""
		Customers.Inv_Address2.CssClass = ""
		Customers.Inv_Address2.ViewCustomAttributes = ""

		' inv_City
		Customers.inv_City.ViewValue = Customers.inv_City.CurrentValue
		Customers.inv_City.CssStyle = ""
		Customers.inv_City.CssClass = ""
		Customers.inv_City.ViewCustomAttributes = ""

		' inv_Province
		If Not IsNull(Customers.inv_Province.CurrentValue) And Customers.inv_Province.CurrentValue <> "" Then
			sSqlWrk = "SELECT [Province] FROM [Province] WHERE [Prov] = '" & ew_AdjustSql(Customers.inv_Province.CurrentValue) & "'"
			sSqlWrk = sSqlWrk & " ORDER BY [Province] Asc"
			Set rswrk = conn.Execute(sSqlWrk)
			If Not rswrk.Eof Then
				Customers.inv_Province.ViewValue = rswrk("Province")
			Else
				Customers.inv_Province.ViewValue = Customers.inv_Province.CurrentValue
			End If
			rswrk.Close
			Set rswrk = Nothing
		Else
			Customers.inv_Province.ViewValue = Null
		End If
		Customers.inv_Province.CssStyle = ""
		Customers.inv_Province.CssClass = ""
		Customers.inv_Province.ViewCustomAttributes = ""

		' inv_EmailAddress
		Customers.inv_EmailAddress.ViewValue = Customers.inv_EmailAddress.CurrentValue
		Customers.inv_EmailAddress.CssStyle = ""
		Customers.inv_EmailAddress.CssClass = ""
		Customers.inv_EmailAddress.ViewCustomAttributes = ""

		' inv_Fax
		Customers.inv_Fax.ViewValue = Customers.inv_Fax.CurrentValue
		Customers.inv_Fax.CssStyle = ""
		Customers.inv_Fax.CssClass = ""
		Customers.inv_Fax.ViewCustomAttributes = ""

		' UserName
		Customers.UserName.ViewValue = Customers.UserName.CurrentValue
		Customers.UserName.CssStyle = ""
		Customers.UserName.CssClass = ""
		Customers.UserName.ViewCustomAttributes = ""

		' Inv_FirstName
		Customers.Inv_FirstName.HrefValue = ""

		' Inv_LastName
		Customers.Inv_LastName.HrefValue = ""

		' Inv_Address
		Customers.Inv_Address.HrefValue = ""

		' Inv_Address2
		Customers.Inv_Address2.HrefValue = ""

		' inv_City
		Customers.inv_City.HrefValue = ""

		' inv_Province
		Customers.inv_Province.HrefValue = ""

		' inv_EmailAddress
		Customers.inv_EmailAddress.HrefValue = ""

		' inv_Fax
		Customers.inv_Fax.HrefValue = ""

		' UserName
		Customers.UserName.HrefValue = ""
	ElseIf Customers.RowType = EW_ROWTYPE_ADD Then ' Add row
	ElseIf Customers.RowType = EW_ROWTYPE_EDIT Then ' Edit row
	ElseIf Customers.RowType = EW_ROWTYPE_SEARCH Then ' Search row
	End If

	' Call Row Rendered event
	Call Customers.Row_Rendered()
End Sub
%>
<%

' Show link optionally based on user id
Function ShowOptionLink()
	ShowOptionLink = True
	If Security.IsLoggedIn() Then
		If Not Security.IsAdmin() Then
			ShowOptionLink = Security.IsValidUserID(Customers.CustomerID.CurrentValue)
		End If
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
