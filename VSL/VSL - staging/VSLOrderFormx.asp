<%
Const EW_PAGE_ID = "list"
Const EW_TABLE_NAME = "Products"
%>
<!--#include file="ewcfg60.asp"--> 
<!--#include file="Productsinfo.asp"-->
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

' Common page loading event (in userfn60.asp)
Call Page_Loading()
%>
<%

' Page load event, used in current page
Call Page_Load()
%>
<%
Products.Export = Request.QueryString("export") ' Get export parameter
sExport = Products.Export ' Get export parameter, used in header
sExportFile = Products.TableVar ' Get export file, used in header
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
	Products.SearchWhere = sSrchWhere ' Save to Session
	nStartRec = 1 ' Reset start record counter
	Products.StartRecordNumber = nStartRec
Else
	Call RestoreSearchParms()
End If

' Build filter
sFilter = "active=yes"
If sDbDetailFilter <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sDbDetailFilter & ")"
End If
If sSrchWhere <> "" Then
	If sFilter <> "" Then sFilter = sFilter & " AND "
	sFilter = sFilter & "(" & sSrchWhere & ")"
End If

' Set up filter in Session
Products.SessionWhere = sFilter
Products.CurrentFilter = ""

' Set Up Sorting Order
SetUpSortOrder()

' Set Return Url
Products.ReturnUrl = "vslOrderFormx.asp"
%>
<!--#include file="header.asp"-->


<table  width="820" border="0" cellpadding="0" cellspacing="0" id="Table_01">
            <tr>
            <td  ><h1 class="h1-header" style="padding:20px 0px 20px 0px; color:#0067A7;">Order VSL#3 Here</h1></td></tr><tr>
              <td  ><div align="left"><span style="color:#333333;font-weight: bold;">Please note:   VSL#3 orders are shipped Monday, Tuesday and Wednesday's only.<br/>  
Any orders received Wednesday after 10am EST; up to and including Friday, will ship the following Monday.</span><br/>
</div></td>
           <!--   <td width="28" valign="top"><img src="images/FontSize.png" border="0" alt=""> </td>
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
                <p><a href="french/VSLOrderForm.asp" class="bodycopy_small">en fran&ccedil;ais &gt;</a></p>
              </div></td>-->
              </tr>
        </table>
        


 
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


<%

' Load recordset
Dim rs
Set rs = LoadRecordset()
nTotalRecs = rs.RecordCount

nStartRec = 1
If nDisplayRecs <= 0 Then ' Display all records
	nDisplayRecs = nTotalRecs
End If
If Not (EW_EXPORT_ALL And Products.Export <> "") Then
	SetUpStartRec() ' Set up start record position
End If
%>




<% If nTotalRecs > 0 Then %>

<form action="vslCartx.asp" method="post" name="fQProductslist" id="fQProductslist" >
         <% if (not isNewCustomer()) then %>
           <strong class="ewHighlightSearch"><em style="color:#000000;">Minimum Qty:2</em></strong>
            <%end if%> 
<div class="t"  align="center">
  <div align="right" class="vslcss"><a href="vslCartx.asp">View Cart</a> :
		    <%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
	
            <a href="login.asp">login</a>
            <%else%>
			<a href="Customersedit.asp">Edit account</a> : <a href="changepwd.asp">Change Password</a> :
            <a href="logout.asp">logout</a>
            <%end if
	set Security =nothing %> 
  </div><table width="785" height="179" border="0" align="center" cellpadding="0" cellspacing="0" class="ewTableNoBorder" id="ewlistmain">
  <tr><td><div align="right">        </div>      
  
      <div align="left">		      </div></td>		
		<td>&nbsp;</td>
		<td>&nbsp;		  </td>
		</tr>
  <tr>
    <td colspan="3" bgcolor="#0A6FA5" height="5px"></td>
    </tr>
<tr><td colspan="3"> <%
If Session(EW_SESSION_MESSAGE) <> "" Then
%>
<p align="center" class="ewmsg"><%= Session(EW_SESSION_MESSAGE) %></p>
<%
	Session(EW_SESSION_MESSAGE) = "" ' Clear message
End If
%>
					<%if((Month(now())=12) and (day(now())<16) and (year(now())=2016)) then %>
						<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" > Annual Sale </font></p>
						<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >Until December 15th, 2016  or  while supplies last</font></p>
						<p style="background-color: #FFFFCC;"><font color="#FF0000" size="+1" >To our valued customers,  Ferring is  pleased to announce our Annual Sale on VSL#3.  </font></p>
						<p style="background-color: #FFFFCC;"><font color="#FF0000"  >Until December 15th, 2016, with every 3 cartons of VSL#3  you purchase, you will receive a 4th carton at no charge.  Free goods will be calculated at time of shipment

</font></p>
						
					<%end if%>
					
                   
                 <!--<span class="ewmsg"> ATTENTION! <br>The final day for ordering VSL#3 in 2016 for next-day delivery will be Wednesday December 21, 2016. All orders placed after this date will be delivered once next-day deliveries are resumed in January 3, 2017.

                 <br>
</span>-->
			
                      </td></tr>
<%
If (EW_EXPORT_ALL And Products.Export <> "") Then
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
	Products.CssClass = "ewTableRow"
	Products.CssStyle = ""

	' Init row event
	Products.RowClientEvents = "onmouseover='ew_MouseOver(this);' onmouseout='ew_MouseOut(this);' onclick='ew_Click(this);'"

	' Display alternate color for rows
	If RowCnt Mod 2 = 0 Then
		Products.CssClass = "ewTableAltRow"
	End If
	Call LoadRowValues(rs) ' Load row values
	Products.RowType = EW_ROWTYPE_VIEW ' Render view
	Call RenderRow()
%>

	<tr>

 
		<!-- Price -->
		<td colspan="4" class="subheading"><b><%= Products.ProductName.ViewValue %></b></td>
		<!-- Active -->
		</tr>
	<tr>
	  <td height="150" align="center" valign="middle"><% If Products.Image_Thumb.HrefValue <> "" Then %>
<% If Not IsNull(Products.Image_Thumb.Upload.DbValue) Then %>
<a href="<%= Products.Image_Thumb.HrefValue %>"><img src="products/thumbs/<%= Products.Image_Thumb.ViewValue %>" border="0"></a>
<% End If %>
<% Else %>
<% If Not IsNull(Products.Image_Thumb.Upload.DbValue) Then %>
<img src="products/thumbs/<%= Products.Image_Thumb.ViewValue %>" border="0">
<% End If %>
<% End If %></td>
	  <td width="288" align="left"><p class="bodybold"><%= Products.Description.ViewValue %></p>
	    <p class="vslcss"><em>UPC: <%= Products.UPC.ViewValue %><br>
	      Item #: <%= Products.ItemNo.ViewValue %></em> </p></td>
	  <td align="right"<%= Products.Active.CellAttributes %>>
	  <p align="right" class="vslcss"><strong>Price:</strong> $<%= Products.Price.ViewValue %> /pack<br> 
<%If(Products.ItemId.ViewValue=99) then%>
<div style="width: 100px; background-color: rgb(0, 101, 166); color: white; height: 80px; font-size: 12px; text-align: center; vertical-align: middle; line-height: 80px;"> Out of stock</div>

<%else%>
	  
	    <br>
	       <strong class="ewHighlightSearch"><em>Free Shipping!</em></strong> </p>
	    <p align="right" class="vslcss"><br>
            <strong>Qty :</strong>
            <input name="<%=right("000" & RowCnt,3)%>_Qty" type="text" value="0" size="5">
            <br>
   
            <input name="ItemId<%=right("000" & RowCnt,3)%>" type="hidden" value="<%=Products.ItemId.ViewValue%>">
  <input name="<%=right("000" & RowCnt,3)%>_Desc" type="hidden" value="<%=Products.ItemNo.ViewValue%>:<%=Products.ProductName.ViewValue%> / <%=Products.UPC.ViewValue%>">
  <br>
  <table cellpadding="0" cellspacing="0"><tr><td width="249"></td><td>
  <input name="Add to cart22" type="image" class="InputNoBorder" value="Add to cart" src="images/addtocart.gif" align="right"></td></tr></table>
	    
<%end if%>
	       </p> </td>
	</tr>
	<tr>
	  <td height="5" colspan="3" bgcolor="#666666"></td>
  </tr>
<%
	End If
	rs.MoveNext
Loop
%>
<tr>


		<!-- Price -->
		<td colspan="3"><div align="right">
        </div></td>		
		</tr>
</table>
  <div >
 
    <p align="left" class="vslcss">&nbsp; </p>
    <p align="left" class="vslcss">Single or multiple boxes of VSL #3 can also be ordered from your local pharmacy. <br>
        <br>
        If you have questions on how to order via the website or about VSL#3, please call Ferring Pharmaceuticals at 1-416-642-0075 or toll free at 1-877-681-7464 during business hours  (Mon - Fri 8:30 am to 4:00 pm EST) or send us an email at VSL3@ferring.com. </p>
    <p align="left" class="vslcss">Please note that there is voice mail availability after hours and the call will be returned the next business day. </p>
    <p align="left" class="vslcss">&nbsp;</p>

    <p align="left" class="vslcss">&nbsp;</p>
    <p align="left" class="vslcss">All prices are in Canadian Dollars </p>
    <p align="left" class="vslcss">&nbsp;</p>
    </div>
</div>
	  </form>
<% End If %>

<%

' Close recordset and connection
rs.Close
Set rs = Nothing
%>


<!--#include file="footer.asp"-->

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
	Set Products = Nothing

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
	BasicSearchSQL = BasicSearchSQL & "[ItemNo] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[UPC] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Image_Thumb] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[ProductName] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Description] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Price] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Image] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Sizes] LIKE '%" & sKeyword & "%' OR "
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
		Products.BasicSearchKeyword = sSearchKeyword
		Products.BasicSearchType = sSearchType
	End If
	BasicSearchWhere = sSearchStr
End Function

' Clear all search parameters
Sub ResetSearchParms()

	' Clear search where
	sSrchWhere = ""
	Products.SearchWhere = sSrchWhere

	' Clear basic search parameters
	Call ResetBasicSearchParms()
End Sub

' Clear all basic search parameters
Sub ResetBasicSearchParms()

	' Clear basic search parameters
	Products.BasicSearchKeyword = ""
	Products.BasicSearchType = ""
End Sub

' Restore all search parameters
Sub RestoreSearchParms()
	sSrchWhere = Products.SearchWhere
End Sub

' Set up Sort parameters based on Sort Links clicked
Sub SetUpSortOrder()
	Dim sOrderBy
	Dim sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		Products.CurrentOrder = Request.QueryString("order")
		Products.CurrentOrderType = Request.QueryString("ordertype")

		' Field ItemId
		Call Products.UpdateSort(Products.ItemId)

		' Field ItemNo
		Call Products.UpdateSort(Products.ItemNo)

		' Field UPC
		Call Products.UpdateSort(Products.UPC)

		' Field Image_Thumb
		Call Products.UpdateSort(Products.Image_Thumb)

		' Field ProductName
		Call Products.UpdateSort(Products.ProductName)

		' Field Description
		Call Products.UpdateSort(Products.Description)

		' Field Price
		Call Products.UpdateSort(Products.Price)

		' Field Active
		Call Products.UpdateSort(Products.Active)

		' Field Image
		Call Products.UpdateSort(Products.Image)

		' Field Sizes
		Call Products.UpdateSort(Products.Sizes)
		Products.StartRecordNumber = 1 ' Reset start position
	End If
	sOrderBy = Products.SessionOrderBy ' Get order by from Session
	If sOrderBy = "" Then
		If Products.SqlOrderBy <> "" Then
			sOrderBy = Products.SqlOrderBy
			Products.SessionOrderBy = sOrderBy
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
			Products.SessionOrderBy = sOrderBy
			Products.ItemId.Sort = ""
			Products.ItemNo.Sort = ""
			Products.UPC.Sort = ""
			Products.Image_Thumb.Sort = ""
			Products.ProductName.Sort = ""
			Products.Description.Sort = ""
			Products.Price.Sort = ""
			Products.Active.Sort = ""
			Products.Image.Sort = ""
			Products.Sizes.Sort = ""
		End If

		' Reset start position
		nStartRec = 1
		Products.StartRecordNumber = nStartRec
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
		Products.StartRecordNumber = nStartRec
	ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
		nPageNo = Request.QueryString(EW_TABLE_PAGE_NO)
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Products.StartRecordNumber = nStartRec
		Else
			nStartRec = Products.StartRecordNumber
		End If
	Else
		nStartRec = Products.StartRecordNumber
	End If

	' Check if correct start record counter
	If Not IsNumeric(nStartRec) Or nStartRec = "" Then ' Avoid invalid start record counter
		nStartRec = 1 ' Reset start record counter
		Products.StartRecordNumber = nStartRec
	ElseIf CLng(nStartRec) > CLng(nTotalRecs) Then ' Avoid starting record > total records
		nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to last page first record
		Products.StartRecordNumber = nStartRec
	ElseIf (nStartRec-1) Mod nDisplayRecs <> 0 Then
		nStartRec = ((nStartRec-1)\nDisplayRecs)*nDisplayRecs+1 ' Point to page boundary
		Products.StartRecordNumber = nStartRec
	End If
End Sub
%>
<%

' Load recordset
Function LoadRecordset()

	' Call Recordset Selecting event
	Call Products.Recordset_Selecting(Products.CurrentFilter)

	' Load list page sql
	Dim sSql
	sSql = Products.ListSQL & "Order by Description desc"

	 'Response.Write sSql ' Uncomment to show SQL for debugging
	' Load recordset

	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = EW_CURSORLOCATION
	rs.Open sSql, conn, 1, 2

	' Call Recordset Selected event
	Call Products.Recordset_Selected(rs)
	Set LoadRecordset = rs
End Function
%>
<%

' Load row based on key values
Function LoadRow()
	Dim rs, sSql, sFilter
	sFilter = Products.SqlKeyFilter
	If Not IsNumeric(Products.ItemId.CurrentValue) Then
		LoadRow = False ' Invalid key, exit
		Exit Function
	End If
	sFilter = Replace(sFilter, "@ItemId@", ew_AdjustSql(Products.ItemId.CurrentValue)) ' Replace key value

	' Call Row Selecting event
	Call Products.Row_Selecting(sFilter)

	' Load sql based on filter
	Products.CurrentFilter = sFilter
	sSql = Products.SQL
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadRow = False
	Else
		LoadRow = True
		rs.MoveFirst
		Call LoadRowValues(rs) ' Load row values

		' Call Row Selected event
		Call Products.Row_Selected(rs)
	End If
	rs.Close
	Set rs = Nothing
End Function

' Load row values from recordset
Sub LoadRowValues(rs)
	Products.ItemId.DbValue = rs("ItemId")
	Products.ItemNo.DbValue = rs("ItemNo")
	Products.UPC.DbValue = rs("UPC")
	Products.Image_Thumb.Upload.DbValue = rs("Image_Thumb")
	Products.ProductName.DbValue = rs("ProductName")
	Products.Description.DbValue = rs("Description")
	Products.Price.DbValue = rs("Price")
	Products.Active.DbValue = ew_IIf(rs("Active"), "1", "0")
	Products.Image.Upload.DbValue = rs("Image")
	Products.Sizes.DbValue = rs("Sizes")
End Sub
%>
<%

' Render row values based on field settings
Sub RenderRow()

	' Call Row Rendering event
	Call Products.Row_Rendering()

	' Common render codes for all row types
	' ItemId

	Products.ItemId.CellCssStyle = ""
	Products.ItemId.CellCssClass = ""

	' ItemNo
	Products.ItemNo.CellCssStyle = ""
	Products.ItemNo.CellCssClass = ""

	' UPC
	Products.UPC.CellCssStyle = ""
	Products.UPC.CellCssClass = ""

	' Image_Thumb
	Products.Image_Thumb.CellCssStyle = ""
	Products.Image_Thumb.CellCssClass = ""

	' ProductName
	Products.ProductName.CellCssStyle = ""
	Products.ProductName.CellCssClass = ""

	' Description
	Products.Description.CellCssStyle = ""
	Products.Description.CellCssClass = ""

	' Price
	Products.Price.CellCssStyle = ""
	Products.Price.CellCssClass = ""

	' Active
	Products.Active.CellCssStyle = ""
	Products.Active.CellCssClass = ""

	' Image
	Products.Image.CellCssStyle = ""
	Products.Image.CellCssClass = ""

	' Sizes
	Products.Sizes.CellCssStyle = ""
	Products.Sizes.CellCssClass = ""
	If Products.RowType = EW_ROWTYPE_VIEW Then ' View row

		' ItemId
		Products.ItemId.ViewValue = Products.ItemId.CurrentValue
		Products.ItemId.CssStyle = ""
		Products.ItemId.CssClass = ""
		Products.ItemId.ViewCustomAttributes = ""

		' ItemNo
		Products.ItemNo.ViewValue = Products.ItemNo.CurrentValue
		Products.ItemNo.CssStyle = ""
		Products.ItemNo.CssClass = ""
		Products.ItemNo.ViewCustomAttributes = ""

		' UPC
		Products.UPC.ViewValue = Products.UPC.CurrentValue
		Products.UPC.CssStyle = ""
		Products.UPC.CssClass = ""
		Products.UPC.ViewCustomAttributes = ""

		' Image_Thumb
		If Not IsNull(Products.Image_Thumb.Upload.DbValue) Then
			Products.Image_Thumb.ViewValue = Products.Image_Thumb.Upload.DbValue
			Products.Image_Thumb.ImageAlt = ""
		Else
			Products.Image_Thumb.ViewValue = ""
		End If
		Products.Image_Thumb.CssStyle = ""
		Products.Image_Thumb.CssClass = ""
		Products.Image_Thumb.ViewCustomAttributes = ""

		' ProductName
		Products.ProductName.ViewValue = Products.ProductName.CurrentValue
		Products.ProductName.CssStyle = ""
		Products.ProductName.CssClass = ""
		Products.ProductName.ViewCustomAttributes = ""

		' Description
		Products.Description.ViewValue = Products.Description.CurrentValue
		If Not IsNull(Products.Description.ViewValue) Then
			Products.Description.ViewValue = Replace(Products.Description.ViewValue, vbLf, "<br>")
		End If
		Products.Description.CssStyle = ""
		Products.Description.CssClass = ""
		Products.Description.ViewCustomAttributes = ""

		' Price
		Products.Price.ViewValue = Products.Price.CurrentValue
		Products.Price.CssStyle = ""
		Products.Price.CssClass = ""
		Products.Price.ViewCustomAttributes = ""

		' Active
		If Products.Active.CurrentValue = "1" Then
			Products.Active.ViewValue = "Yes"
		Else
			Products.Active.ViewValue = "No"
		End If
		Products.Active.CssStyle = ""
		Products.Active.CssClass = ""
		Products.Active.ViewCustomAttributes = ""

		' Image
		If Not IsNull(Products.Image.Upload.DbValue) Then
			Products.Image.ViewValue = Products.Image.Upload.DbValue
			Products.Image.ImageAlt = ""
		Else
			Products.Image.ViewValue = ""
		End If
		Products.Image.CssStyle = ""
		Products.Image.CssClass = ""
		Products.Image.ViewCustomAttributes = ""

		' Sizes
		Products.Sizes.ViewValue = Products.Sizes.CurrentValue
		Products.Sizes.CssStyle = ""
		Products.Sizes.CssClass = ""
		Products.Sizes.ViewCustomAttributes = ""

		' ItemId
		Products.ItemId.HrefValue = ""

		' ItemNo
		Products.ItemNo.HrefValue = ""

		' UPC
		Products.UPC.HrefValue = ""

		' Image_Thumb
		Products.Image_Thumb.HrefValue = ""

		' ProductName
		Products.ProductName.HrefValue = ""

		' Description
		Products.Description.HrefValue = ""

		' Price
		Products.Price.HrefValue = ""

		' Active
		Products.Active.HrefValue = ""

		' Image
		Products.Image.HrefValue = ""

		' Sizes
		Products.Sizes.HrefValue = ""
	ElseIf Products.RowType = EW_ROWTYPE_ADD Then ' Add row
	ElseIf Products.RowType = EW_ROWTYPE_EDIT Then ' Edit row
	ElseIf Products.RowType = EW_ROWTYPE_SEARCH Then ' Search row
	End If

	' Call Row Rendered event
	Call Products.Row_Rendered()
End Sub
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
