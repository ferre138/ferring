<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="ewcfg9.asp"-->

<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-cold-1" />
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<%
datef= request.querystring("datef")
datet= request.querystring("datet")
pr= request.querystring("pr")
sqlfilter=" 1=1 "
if(datef<>"") then sqlfilter= sqlfilter & " and (Paymentdate>= #"  & datef & "#)"
if(datet<>"") then sqlfilter= sqlfilter &  " and (Paymentdate<= #"  & datet & "#)"
if(pr<>"") then sqlfilter= sqlfilter &  " and ((Ship_Province)=""" & pr & """)"
reportheading = "Daily Report "
if(pr<>"") then reportheading = reportheading & " Province:" & pr
if(datef<>"") then reportheading = reportheading & "   Period (" & datef & " : " & datet & ")"
'response.write sqlfilter
%>
<%

' ASPMaker configuration for Table Daily report
Dim Daily_report

' Define table class
Class cDaily_report

	' Class Initialize
	Private Sub Class_Initialize()
		UseTokenInUrl = EW_USE_TOKEN_IN_URL
		ExportOriginalValue = EW_EXPORT_ORIGINAL_VALUE
		ExportAll = True
		Set RowAttrs = New cAttributes ' Row attributes
		Call ew_SetArObj(Fields, "InvoiceId", InvoiceId)
		Call ew_SetArObj(Fields, "CustomerId", CustomerId)
		Call ew_SetArObj(Fields, "Inv_FirstName", Inv_FirstName)
		Call ew_SetArObj(Fields, "Inv_LastName", Inv_LastName)
		Call ew_SetArObj(Fields, "inv_EmailAddress", inv_EmailAddress)
		Call ew_SetArObj(Fields, "payment_gross", payment_gross)
		Call ew_SetArObj(Fields, "payment_fee", payment_fee)
		Call ew_SetArObj(Fields, "Tax", Tax)
		Call ew_SetArObj(Fields, "Shipping", Shipping)
		Call ew_SetArObj(Fields, "Ship_Province", Ship_Province)
		Call ew_SetArObj(Fields, "Paymentdate", Paymentdate)
		Call ew_SetArObj(Fields, "payment_status", payment_status)
	End Sub

	' Reset attributes for table object
	Public Sub ResetAttrs()
		CssClass = ""
		CssStyle = ""
		RowAttrs.Clear()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				Call fld.ResetAttrs()
			Next
		End If
	End Sub

	' Setup field titles
	Public Sub SetupFieldTitles()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				If fld.FldTitle <> "" Then
					fld.EditAttrs.AddAttribute "onmouseover", "ew_ShowTitle(this, '" & ew_JsEncode3(fld.FldTitle) & "');", True
					fld.EditAttrs.AddAttribute "onmouseout", "ew_HideTooltip();", True
				End If
			Next
		End If
	End Sub

	' Define table level constants
	' Use table token in Url

	Dim UseTokenInUrl

	' Table variable
	Public Property Get TableVar()
		TableVar = "Daily_report"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Daily report"
	End Property

	' Table type
	Public Property Get TableType()
		TableType = "REPORT"
	End Property

	' Table caption
	Public Property Get TableCaption()
		TableCaption = Language.TablePhrase(TableVar, "TblCaption")
	End Property

	' Page caption
	Public Property Get PageCaption(Page)
		PageCaption = Language.TablePhrase(TableVar, "TblPageCaption" & Page)
		If PageCaption = "" Then PageCaption = "Page " & Page
	End Property

	' Report Group Level SQL
	Public Property Get SqlGroupSelect() ' Select
		SqlGroupSelect = "SELECT DISTINCT [Ship_Province],[Paymentdate] FROM [ReportOrders]"
	End Property

	Public Property Get SqlGroupWhere() ' Where
		SqlGroupWhere = sqlfilter
	End Property

	Public Property Get SqlGroupGroupBy() ' Group By
		SqlGroupGroupBy = ""
	End Property

	Public Property Get SqlGroupHaving() ' Having
		SqlGroupHaving = ""
	End Property

	Public Property Get SqlGroupOrderBy() ' Order By
		SqlGroupOrderBy = "[Ship_Province] ASC,[Paymentdate] ASC"
	End Property

	' Report Detail Level SQL
	Public Property Get SqlDetailSelect() ' Select
		SqlDetailSelect = "SELECT * FROM [ReportOrders]"
	End Property

	Public Property Get SqlDetailWhere() ' Where
		SqlDetailWhere = ""
	End Property

	Public Property Get SqlDetailGroupBy() ' Group By
		SqlDetailGroupBy = ""
	End Property

	Public Property Get SqlDetailHaving() ' Having
		SqlDetailHaving = ""
	End Property

	Public Property Get SqlDetailOrderBy() ' Order By
		SqlDetailOrderBy = "[InvoiceId] ASC"
	End Property

	' SQL variables
	Dim CurrentFilter ' Current filter
	Dim CurrentOrder ' Current order
	Dim CurrentOrderType ' Current order type

	' Return report group sql
	Public Property Get GroupSQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = ""
		GroupSQL = ew_BuildSelectSql(SqlGroupSelect, SqlGroupWhere, SqlGroupGroupBy, SqlGroupHaving, SqlGroupOrderBy, sFilter, sSort)
	End Property

	' Return report detail sql
	Public Property Get DetailSQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = ""
		DetailSQL = ew_BuildSelectSql(SqlDetailSelect, SqlDetailWhere, SqlDetailGroupBy, SqlDetailHaving, SqlDetailOrderBy, sFilter, sSort)
	End Property

	' Return url
	Public Property Get ReturnUrl()

		' Get referer url automatically
		If Request.ServerVariables("HTTP_REFERER") <> "" Then
			If ew_ReferPage <> ew_CurrentPage And ew_ReferPage <> "login.asp" Then ' Referer not same page or login page
				Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = Request.ServerVariables("HTTP_REFERER") ' Save to Session
			End If
		End If
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "Daily_reportreport.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "Daily_reportreport.asp"
	End Function

	' View url
	Public Function ViewUrl()
		ViewUrl = KeyUrl("", UrlParm(""))
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = ""

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		KeyUrl = sUrl
	End Function

	' Sort Url
	Public Property Get SortUrl(fld)
		If CurrentAction <> "" Or Export <> "" Or (fld.FldType = 201 Or fld.FldType = 203 Or fld.FldType = 205 Or fld.FldType = 141) Then
			SortUrl = ""
		ElseIf fld.Sortable Then
			SortUrl = ew_CurrentPage
			Dim sUrlParm
			sUrlParm = UrlParm("order=" & Server.URLEncode(fld.FldName) & "&amp;ordertype=" & fld.ReverseSort)
			SortUrl = SortUrl & "?" & sUrlParm
		Else
			SortUrl = ""
		End If
	End Property

	' Url parm
	Function UrlParm(parm)
		If UseTokenInUrl Then
			UrlParm = "t=Daily_report"
		Else
			UrlParm = ""
		End If
		If parm <> "" Then
			If UrlParm <> "" Then UrlParm = UrlParm & "&"
			UrlParm = UrlParm & parm
		End If
	End Function

	' Get record keys from Form/QueryString/Session
	Public Function GetRecordKeys()
		Dim arKeys, arKey, cnt, i, bHasKey
		bHasKey = False

		' Check ObjForm first
		If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
			ObjForm.Index = 0
			If ObjForm.HasValue("key_m") Then
				arKeys = ObjForm.GetValue("key_m")
				If Not IsArray(arKeys) Then
					arKeys = Array(arKeys)
				End If
				bHasKey = True
			End If
		End If

		' Check Form/QueryString
		If Not bHasKey Then
			If Request.Form("key_m").Count > 0 Then
				cnt = Request.Form("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.Form("key_m")(i)
				Next
			ElseIf Request.QueryString("key_m").Count > 0 Then
				cnt = Request.QueryString("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.QueryString("key_m")(i)
				Next
			ElseIf Request.QueryString <> "" Then
				ReDim arKeys(0)

				'GetRecordKeys = arKeys ' do not return yet, so the values will also be checked by the following code
			End If
		End If

		' Check keys
		Dim ar, key
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
						Dim skip
						skip = False
						If Not skip Then
							If IsArray(ar) Then
								ReDim Preserve ar(UBound(ar)+1)
							Else
								ReDim ar(0)
							End If
							ar(UBound(ar)) = key
						End If
			Next
		End If
		GetRecordKeys = ar
	End Function

	' Get key filter
	Public Function GetKeyFilter()
		Dim arKeys, sKeyFilter, i, key
		arKeys = GetRecordKeys()
		sKeyFilter = ""
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
				If sKeyFilter <> "" Then sKeyFilter = sKeyFilter & " OR "
				sKeyFilter = sKeyFilter & "(" & KeyFilter & ")"
			Next
		End If
		GetKeyFilter = sKeyFilter
	End Function

	' Function LoadRecordCount
	' - Load record count based on filter
	Public Function LoadRecordCount(sFilter)
		Dim wrkrs
		Set wrkrs = LoadRs(sFilter)
		If Not wrkrs Is Nothing Then
			LoadRecordCount = wrkrs.RecordCount
		Else
			LoadRecordCount = 0
		End If
		Set wrkrs = Nothing
	End Function

	' Function LoadRs
	' - Load Rows based on filter
	Public Function LoadRs(sFilter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim RsRows, sSql

		' Set up filter (Sql Where Clause) and get Return Sql
		CurrentFilter = sFilter
		sSql = SQL
		Err.Clear
		Set RsRows = Server.CreateObject("ADODB.Recordset")
		RsRows.CursorLocation = EW_CURSORLOCATION
		RsRows.Open sSql, Conn, 3, 1, 1 ' adOpenStatic, adLockReadOnly, adCmdText
		If Err.Number <> 0 Then
			Err.Clear
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		ElseIf RsRows.Eof Then
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		Else
			Set LoadRs = RsRows
		End If
	End Function

	' Row Type
	Private m_RowType

	Public Property Get RowType()
		RowType = m_RowType
	End Property

	Public Property Let RowType(v)
		m_RowType = v
	End Property
	Dim CssClass ' Css class
	Dim CssStyle' Css style

'	Dim RowClientEvents ' Row client events
	Dim RowAttrs ' Row attributes

	' Row Styles
	Public Property Get RowStyles()
		Dim sAtt, Value
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If RowAttrs.Exists("style") Then
			Value = RowAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If RowAttrs.Exists("class") Then
			Value = RowAttrs.Item("class")
			If Trim(Value) <> "" Then
				sClass = sClass & " " & Value
			End If
		End If
		If Trim(sStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(sStyle) & """" 
		End If
		If Trim(sClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(sClass) & """" 
		End If
		RowStyles = sAtt
	End Property

	' Row Attribute
	Public Property Get RowAttributes()
		Dim sAtt, Attr, Value, i
		sAtt = RowStyles
		If m_Export = "" Then

'			If Trim(RowClientEvents) <> "" Then
'				sAtt = sAtt & " " & Trim(RowClientEvents)
'			End If

			For i = 0 to UBound(RowAttrs.Attributes)
				Attr = RowAttrs.Attributes(i)(0)
				Value = RowAttrs.Attributes(i)(1)
				If Attr <> "style" And Attr <> "class" And Attr <> "" And Value <> "" Then
					sAtt = sAtt & " " & Attr & "=""" & Value & """"
				End If
			Next
		End If
		RowAttributes = sAtt
	End Property

	' Export
	Private m_Export

	Public Property Get Export()
		Export = m_Export
	End Property

	Public Property Let Export(v)
		m_Export = v
	End Property

	' Export Original Value
	Dim ExportOriginalValue

	' Export All
	Dim ExportAll

	' Send Email
	Dim SendEmail

	' Custom Inner Html
	Dim TableCustomInnerHtml

	' ----------------
	'  Field objects
	' ----------------
	' Field InvoiceId
	Private m_InvoiceId

	Public Property Get InvoiceId()
		If Not IsObject(m_InvoiceId) Then
			Set m_InvoiceId = NewFldObj("Daily_report", "Daily report", "x_InvoiceId", "InvoiceId", "[InvoiceId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_InvoiceId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set InvoiceId = m_InvoiceId
	End Property

	' Field CustomerId
	Private m_CustomerId

	Public Property Get CustomerId()
		If Not IsObject(m_CustomerId) Then
			Set m_CustomerId = NewFldObj("Daily_report", "Daily report", "x_CustomerId", "CustomerId", "[CustomerId]", 3, 8, "", False, False, "FORMATTED TEXT")
			m_CustomerId.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set CustomerId = m_CustomerId
	End Property

	' Field Inv_FirstName
	Private m_Inv_FirstName

	Public Property Get Inv_FirstName()
		If Not IsObject(m_Inv_FirstName) Then
			Set m_Inv_FirstName = NewFldObj("Daily_report", "Daily report", "x_Inv_FirstName", "Inv_FirstName", "[Inv_FirstName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_FirstName = m_Inv_FirstName
	End Property

	' Field Inv_LastName
	Private m_Inv_LastName

	Public Property Get Inv_LastName()
		If Not IsObject(m_Inv_LastName) Then
			Set m_Inv_LastName = NewFldObj("Daily_report", "Daily report", "x_Inv_LastName", "Inv_LastName", "[Inv_LastName]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Inv_LastName = m_Inv_LastName
	End Property

	' Field inv_EmailAddress
	Private m_inv_EmailAddress

	Public Property Get inv_EmailAddress()
		If Not IsObject(m_inv_EmailAddress) Then
			Set m_inv_EmailAddress = NewFldObj("Daily_report", "Daily report", "x_inv_EmailAddress", "inv_EmailAddress", "[inv_EmailAddress]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set inv_EmailAddress = m_inv_EmailAddress
	End Property

	' Field payment_gross
	Private m_payment_gross

	Public Property Get payment_gross()
		If Not IsObject(m_payment_gross) Then
			Set m_payment_gross = NewFldObj("Daily_report", "Daily report", "x_payment_gross", "payment_gross", "[payment_gross]", 5, 8, "", False, False, "FORMATTED TEXT")
			m_payment_gross.FldDefaultErrMsg = Language.Phrase("IncorrectFloat")
		End If
		Set payment_gross = m_payment_gross
	End Property

	' Field payment_fee
	Private m_payment_fee

	Public Property Get payment_fee()
		If Not IsObject(m_payment_fee) Then
			Set m_payment_fee = NewFldObj("Daily_report", "Daily report", "x_payment_fee", "payment_fee", "[payment_fee]", 5, 8, "", False, False, "FORMATTED TEXT")
			m_payment_fee.FldDefaultErrMsg = Language.Phrase("IncorrectFloat")
		End If
		Set payment_fee = m_payment_fee
	End Property

	' Field Tax
	Private m_Tax

	Public Property Get Tax()
		If Not IsObject(m_Tax) Then
			Set m_Tax = NewFldObj("Daily_report", "Daily report", "x_Tax", "Tax", "[Tax]", 5, 8, "", False, False, "FORMATTED TEXT")
			m_Tax.FldDefaultErrMsg = Language.Phrase("IncorrectFloat")
		End If
		Set Tax = m_Tax
	End Property

	' Field Shipping
	Private m_Shipping

	Public Property Get Shipping()
		If Not IsObject(m_Shipping) Then
			Set m_Shipping = NewFldObj("Daily_report", "Daily report", "x_Shipping", "Shipping", "[Shipping]", 5, 8, "", False, False, "FORMATTED TEXT")
			m_Shipping.FldDefaultErrMsg = Language.Phrase("IncorrectFloat")
		End If
		Set Shipping = m_Shipping
	End Property

	' Field Ship_Province
	Private m_Ship_Province

	Public Property Get Ship_Province()
		If Not IsObject(m_Ship_Province) Then
			Set m_Ship_Province = NewFldObj("Daily_report", "Daily report", "x_Ship_Province", "Ship_Province", "[Ship_Province]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Ship_Province = m_Ship_Province
	End Property

	' Field Paymentdate
	Private m_Paymentdate

	Public Property Get Paymentdate()
		If Not IsObject(m_Paymentdate) Then
			Set m_Paymentdate = NewFldObj("Daily_report", "Daily report", "x_Paymentdate", "Paymentdate", "[Paymentdate]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set Paymentdate = m_Paymentdate
	End Property

	' Field payment_status
	Private m_payment_status

	Public Property Get payment_status()
		If Not IsObject(m_payment_status) Then
			Set m_payment_status = NewFldObj("Daily_report", "Daily report", "x_payment_status", "payment_status", "[payment_status]", 202, 8, "", False, False, "FORMATTED TEXT")
		End If
		Set payment_status = m_payment_status
	End Property
	Dim Fields ' Fields

	' Create new field object
	Private Function NewFldObj(TblVar, TblName, FldVar, FldName, FldExpression, FldType, FldDtFormat, FldVirtualExp, FldVirtual, FldForceSelect, FldViewTag)
		Dim fld
		Set fld = New cField
		fld.TblVar = TblVar
		fld.TblName = TblName
		fld.FldVar = FldVar
		fld.FldName = FldName
		fld.FldExpression = FldExpression
		fld.FldType = FldType
		fld.FldDataType = ew_FieldDataType(FldType)
		fld.FldDateTimeFormat = FldDtFormat
		fld.FldVirtualExpression = FldVirtualExp
		fld.FldIsVirtual = FldVirtual
		fld.FldForceSelection = FldForceSelect
		fld.FldViewTag = FldViewTag
		Set NewFldObj = fld
	End Function

	' Row Rendering event
	Sub Row_Rendering()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Rendered event
	Sub Row_Rendered()

		' To view properties of field class, use:
		' Response.Write <FieldName>.AsString() 

	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_InvoiceId) Then Set m_InvoiceId = Nothing
		If IsObject(m_CustomerId) Then Set m_CustomerId = Nothing
		If IsObject(m_Inv_FirstName) Then Set m_Inv_FirstName = Nothing
		If IsObject(m_Inv_LastName) Then Set m_Inv_LastName = Nothing
		If IsObject(m_inv_EmailAddress) Then Set m_inv_EmailAddress = Nothing
		If IsObject(m_payment_gross) Then Set m_payment_gross = Nothing
		If IsObject(m_payment_fee) Then Set m_payment_fee = Nothing
		If IsObject(m_Tax) Then Set m_Tax = Nothing
		If IsObject(m_Shipping) Then Set m_Shipping = Nothing
		If IsObject(m_Ship_Province) Then Set m_Ship_Province = Nothing
		If IsObject(m_Paymentdate) Then Set m_Paymentdate = Nothing
		If IsObject(m_payment_status) Then Set m_payment_status = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Daily_report_report
Set Daily_report_report = New cDaily_report_report
Set Page = Daily_report_report

' Page init processing
Call Daily_report_report.Page_Init()

' Page main processing
Call Daily_report_report.Page_Main()
%>
<!--#include file="header.asp"-->
<% If Daily_report.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<% End If %>
<% If Daily_report.Export = "" Then %>
<% End If %>
<% Daily_report_report.ShowPageHeader() %>
<p class="aspmaker ewTitle"><%= Language.Phrase("TblTypeReport") %><%= Daily_report.TableCaption %>
&nbsp;&nbsp;<% Daily_report_report.ExportOptions.Render "body", "" %>
</p>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td align="center" bgcolor="#CCCCCC"><span class="aspmaker "><strong>Province </strong></span></td>
    <td bgcolor="#CCCCCC"></td>
    <td colspan="5" align="center" bgcolor="#CCCCCC"><span class="aspmaker "><strong>Date</strong></span></td>
    <td width="2" bgcolor="#CCCCCC"></td>
    <td bgcolor="#CCCCCC">&nbsp;</td>
    <td bgcolor="#CCCCCC">&nbsp;</td>
    <td bgcolor="#CCCCCC">&nbsp;</td>
  <tr>
<td><select name="filterProv" id="filterProv">
	<option value="" selected>
Please Select
</option>

<option value="AL">
Alberta
</option>

<option value="BC">
British Columbia
</option>

<option value="MB">
Manitoba
</option>

<option value="NB">

New Brunswick
</option>

<option value="NL">
Newfoundland and Labrador
</option>

<option value="NT">
Northwest Territories
</option>

<option value="NS">
Nova Scotia
</option>

<option value="NU">
Nunavut

</option>

<option value="ON">
Ontario
</option>

<option value="PE">
Prince Edward Island
</option>

<option value="QC">
Quebec
</option>

<option value="SK">
Saskatchewan
</option>

<option value="YT">
Yukon
</option>
</select> </td>
<td width="3" bgcolor="#CCCCCC"></td>
 <td><span class="aspmaker ">From :</span>
   <input type="text" name="date" id="f_date_c" readonly="1" size="10" /></td>
 <td><img src="calendar/img.gif" id="f_trigger_c" style="cursor: pointer; border: 0px solid red;" title="Date selector"
      onmouseover="this.style.background='red';" onmouseout="this.style.background=''" /></td>
 <td>&nbsp;</td>
	  <td><span class="aspmaker "> To :</span>
	    <input type="text" name="date" id="f_date_a" readonly="1"  size="10" /></td>
 <td><img src="calendar/img.gif" id="f_trigger_a" style="cursor: pointer; border: 0px solid red;" title="Date selector"
      onmouseover="this.style.background='red';" onmouseout="this.style.background=''" /></td>
 <td bgcolor="#CCCCCC"></td>
	  <td><a href="javascript:filter();"><img src="images/Filter.png" width="50" height="25" border="0" /></a></td>
	  <td>&nbsp;</td>
	  <td><a href="javascript:clearall();"><img src="images/clear.png" width="50" height="25" border="0" /></a></td>  </tr>
  <tr>
    <td height="3px" colspan="11" bgcolor="#666666"></td></tr>
	  <tr>
    <td height="25px" colspan="11" class="ewGroupName" ><%=reportheading%></td></tr>
</table>

<script type="text/javascript">
    Calendar.setup({
        inputField     :    "f_date_c",     // id of the input field
        ifFormat       :    "%m/%d/%Y",      // format of the input field
        button         :    "f_trigger_c",  // trigger for the calendar (button ID)
        align          :    "Br",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
 

<script type="text/javascript">
    Calendar.setup({
        inputField     :    "f_date_a",     // id of the input field
        ifFormat       :    "%m/%d/%Y",      // format of the input field
        button         :    "f_trigger_a",  // trigger for the calendar (button ID)
        align          :    "Br",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
	
	document.getElementById("f_date_c").value="<%=datef%>";
	document.getElementById("f_date_a").value="<%=datet%>";
	document.getElementById("filterProv").value="<%=pr%>";
	
	function clearall()
	{
		document.getElementById("f_date_c").value="";
	document.getElementById("f_date_a").value="";
	document.getElementById("filterProv").value="";
	filter();
	}
	function filter()
	{var s;
	s=""
	 var x=document.getElementById("f_date_c");
	 if(x.value!="") s+= "datef=" + x.value + "&";
	 x=document.getElementById("f_date_a");
	if(x.value!="") s+= "datet=" + x.value + "&";
	 x=document.getElementById("filterProv");
	 if(x.value!="") s+="pr=" + x.value + "&";
	 
	 window.location="reportDaily.asp?" + s;
	}
</script>
<form method="post">
<table class="ewReportTable" cellspacing="-1">
<%
Daily_report_report.DefaultFilter = ""
Daily_report_report.ReportFilter = Daily_report_report.DefaultFilter
If Daily_report_report.DbDetailFilter <> "" Then
	If Daily_report_report.ReportFilter <> "" Then Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & " AND "
	Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "(" & Daily_report_report.DbDetailFilter & ")"
End If

' Set up filter and load Group level sql
Daily_report.CurrentFilter = Daily_report_report.ReportFilter
Daily_report_report.ReportSql = Daily_report.GroupSQL

 'Response.Write Daily_report_report.ReportSql
' Load recordset

Set Daily_report_report.Recordset = Server.CreateObject("ADODB.Recordset")
Daily_report_report.Recordset.CursorLocation = EW_CURSORLOCATION
Daily_report_report.Recordset.Open Daily_report_report.ReportSql, Conn, 1, EW_RECORDSET_LOCKTYPE

' Get First Row
If Not Daily_report_report.Recordset.Eof Then
	Daily_report.Ship_Province.DbValue = Daily_report_report.Recordset("Ship_Province")
	Daily_report_report.ReportGroups(0) = Daily_report.Ship_Province.DbValue
	Daily_report.Paymentdate.DbValue = Daily_report_report.Recordset("Paymentdate")
	Daily_report_report.ReportGroups(1) = Daily_report.Paymentdate.DbValue
End If
Daily_report_report.RecCnt = 0
Daily_report_report.ReportCounts(0) = 0
Call Daily_report_report.ChkLvlBreak()
Do While (Not Daily_report_report.Recordset.Eof)

	' Render for view
	Daily_report.RowType = EW_ROWTYPE_VIEW
	Call Daily_report.ResetAttrs()
	Call Daily_report_report.RenderRow()

	' Show group headers
	If Daily_report_report.LevelBreak(1) Then ' Reset counter and aggregation
%>
	<tr><td class="ewGroupField"><span class="aspmaker"><%= Daily_report.Ship_Province.FldCaption %></span></td>
	<td colspan=9 class="ewGroupName"><span class="aspmaker">
<div<%= Daily_report.Ship_Province.ViewAttributes %>><%= Daily_report.Ship_Province.ViewValue %></div>
</span></td></tr>
<%
	End If
	If Daily_report_report.LevelBreak(2) Then ' Reset counter and aggregation
%>
	<tr><td class="ewGroupField"><span class="aspmaker"><%=  Daily_report.Paymentdate.CurrentValue %></span></td>
	<td colspan=9 class="ewGroupName"><span class="aspmaker">
<div<%= Daily_report.Paymentdate.ViewAttributes %>></div>
</span></td></tr>
<%
	End If

	' Get detail records
	Daily_report_report.ReportFilter = Daily_report_report.DefaultFilter
	If Daily_report_report.ReportFilter <> "" Then Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & " AND "
	If IsNull(Daily_report.Ship_Province.CurrentValue) Then
		Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "([Ship_Province] IS NULL)"
	Else
		Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "([Ship_Province] = '" & ew_AdjustSql(Daily_report.Ship_Province.CurrentValue) & "')"
	End If
	If Daily_report_report.ReportFilter <> "" Then Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & " AND "
	If IsNull(Daily_report.Paymentdate.CurrentValue) Then
		Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "([Paymentdate] IS NULL)"
	Else
		Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "([Paymentdate] = '" & ew_AdjustSql(Daily_report.Paymentdate.CurrentValue) & "')"
	End If
	If Daily_report_report.DbDetailFilter <> "" Then
		If Daily_report_report.ReportFilter <> "" Then Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & " AND "
		Daily_report_report.ReportFilter = Daily_report_report.ReportFilter & "(" & Daily_report_report.DbDetailFilter & ")"
	End If

	' Set up detail SQL
	Daily_report.CurrentFilter = Daily_report_report.ReportFilter
	Daily_report_report.ReportSql = Daily_report.DetailSQL

	' Load detail records
	Set Daily_report_report.DetailRecordset = Server.CreateObject("ADODB.Recordset")
	Daily_report_report.DetailRecordset.CursorLocation = EW_CURSORLOCATION
	Daily_report_report.DetailRecordset.Open Daily_report_report.ReportSql, Conn, 1, EW_RECORDSET_LOCKTYPE
	Daily_report_report.DtlRecordCount = Daily_report_report.DetailRecordset.RecordCount

	' Initialize Aggregate
	If Not Daily_report_report.DetailRecordset.Eof Then
		Daily_report_report.RecCnt = Daily_report_report.RecCnt + 1
		Daily_report.payment_gross.DbValue = Daily_report_report.DetailRecordset("payment_gross")
		Daily_report.payment_fee.DbValue = Daily_report_report.DetailRecordset("payment_fee")
		Daily_report.Tax.DbValue = Daily_report_report.DetailRecordset("Tax")
		Daily_report.Shipping.DbValue = Daily_report_report.DetailRecordset("Shipping")
	End If
	If Daily_report_report.RecCnt = 1 Then
		Daily_report_report.ReportCounts(0) = 0
		Daily_report_report.ReportTotals(0,5) = 0
		Daily_report_report.ReportTotals(0,6) = 0
		Daily_report_report.ReportTotals(0,7) = 0
		Daily_report_report.ReportTotals(0,8) = 0
	End If
	For i = 1 to 2
		If Daily_report_report.LevelBreak(i) Then ' Reset counter and aggregation
			Daily_report_report.ReportCounts(i) = 0
			Daily_report_report.ReportTotals(i, 5) = 0
			Daily_report_report.ReportTotals(i, 6) = 0
			Daily_report_report.ReportTotals(i, 7) = 0
			Daily_report_report.ReportTotals(i, 8) = 0
		End If
	Next
	Daily_report_report.ReportCounts(0) = Daily_report_report.ReportCounts(0) + Daily_report_report.DtlRecordCount
	Daily_report_report.ReportCounts(1) = Daily_report_report.ReportCounts(1) + Daily_report_report.DtlRecordCount
	Daily_report_report.ReportCounts(2) = Daily_report_report.ReportCounts(2) + Daily_report_report.DtlRecordCount
%>
	<tr>
		<td></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.InvoiceId.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.CustomerId.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.Inv_FirstName.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.Inv_LastName.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.inv_EmailAddress.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.payment_gross.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.payment_fee.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.Tax.FldCaption %></span></td>
		<td valign="top" class="ewGroupHeader"><span class="aspmaker"><%= Daily_report.Shipping.FldCaption %></span></td>
	</tr>
<%
	Do While Not Daily_report_report.DetailRecordset.Eof
		Daily_report.InvoiceId.DbValue = Daily_report_report.DetailRecordset("InvoiceId")
		Daily_report.CustomerId.DbValue = Daily_report_report.DetailRecordset("CustomerId")
		Daily_report.Inv_FirstName.DbValue = Daily_report_report.DetailRecordset("Inv_FirstName")
		Daily_report.Inv_LastName.DbValue = Daily_report_report.DetailRecordset("Inv_LastName")
		Daily_report.inv_EmailAddress.DbValue = Daily_report_report.DetailRecordset("inv_EmailAddress")
		Daily_report.payment_gross.DbValue = Daily_report_report.DetailRecordset("payment_gross")
		Daily_report_report.ReportTotals(0, 5) = Daily_report_report.ReportTotals(0, 5) + Daily_report.payment_gross.CurrentValue
		Daily_report_report.ReportTotals(1, 5) = Daily_report_report.ReportTotals(1, 5) + Daily_report.payment_gross.CurrentValue
		Daily_report_report.ReportTotals(2, 5) = Daily_report_report.ReportTotals(2, 5) + Daily_report.payment_gross.CurrentValue
		Daily_report.payment_fee.DbValue = Daily_report_report.DetailRecordset("payment_fee")
		Daily_report_report.ReportTotals(0, 6) = Daily_report_report.ReportTotals(0, 6) + Daily_report.payment_fee.CurrentValue
		Daily_report_report.ReportTotals(1, 6) = Daily_report_report.ReportTotals(1, 6) + Daily_report.payment_fee.CurrentValue
		Daily_report_report.ReportTotals(2, 6) = Daily_report_report.ReportTotals(2, 6) + Daily_report.payment_fee.CurrentValue
		Daily_report.Tax.DbValue = Daily_report_report.DetailRecordset("Tax")
		Daily_report_report.ReportTotals(0, 7) = Daily_report_report.ReportTotals(0, 7) + Daily_report.Tax.CurrentValue
		Daily_report_report.ReportTotals(1, 7) = Daily_report_report.ReportTotals(1, 7) + Daily_report.Tax.CurrentValue
		Daily_report_report.ReportTotals(2, 7) = Daily_report_report.ReportTotals(2, 7) + Daily_report.Tax.CurrentValue
		Daily_report.Shipping.DbValue = Daily_report_report.DetailRecordset("Shipping")
		Daily_report_report.ReportTotals(0, 8) = Daily_report_report.ReportTotals(0, 8) + Daily_report.Shipping.CurrentValue
		Daily_report_report.ReportTotals(1, 8) = Daily_report_report.ReportTotals(1, 8) + Daily_report.Shipping.CurrentValue
		Daily_report_report.ReportTotals(2, 8) = Daily_report_report.ReportTotals(2, 8) + Daily_report.Shipping.CurrentValue

		' Render for view
		Daily_report.RowType = EW_ROWTYPE_VIEW
		Call Daily_report.ResetAttrs()
		Call Daily_report_report.RenderRow()
%>
	<tr>
		<td></td>
		<td><span class="aspmaker">
<div<%= Daily_report.InvoiceId.ViewAttributes %>><%= Daily_report.InvoiceId.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.CustomerId.ViewAttributes %>><%= Daily_report.CustomerId.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Inv_FirstName.ViewAttributes %>><%= Daily_report.Inv_FirstName.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Inv_LastName.ViewAttributes %>><%= Daily_report.Inv_LastName.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.inv_EmailAddress.ViewAttributes %>><%= Daily_report.inv_EmailAddress.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_gross.ViewAttributes %>><%= Daily_report.payment_gross.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_fee.ViewAttributes %>><%= Daily_report.payment_fee.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Tax.ViewAttributes %>><%= Daily_report.Tax.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Shipping.ViewAttributes %>><%= Daily_report.Shipping.ViewValue %></div>
</span></td>
	</tr>
<%
		Daily_report_report.DetailRecordset.MoveNext
	Loop
	Daily_report_report.DetailRecordset.Close
	Set Daily_report_report.DetailRecordset = Nothing

	' Save old group data
	Daily_report_report.ReportGroups(0) = Daily_report.Ship_Province.CurrentValue
	Daily_report_report.ReportGroups(1) = Daily_report.Paymentdate.CurrentValue

	' Get next record
	Daily_report_report.Recordset.MoveNext
	If Daily_report_report.Recordset.Eof Then
		Daily_report_report.RecCnt = 0 ' EOF, force all level breaks
	Else
		Daily_report.Ship_Province.DbValue = Daily_report_report.Recordset("Ship_Province")
		Daily_report.Paymentdate.DbValue = Daily_report_report.Recordset("Paymentdate")
	End If
	Call Daily_report_report.ChkLvlBreak()

	' Show Footers
	If Daily_report_report.LevelBreak(2) Then
		Daily_report.Paymentdate.CurrentValue = Daily_report_report.ReportGroups(1)

		' Render row for view
		Daily_report.RowType = EW_ROWTYPE_VIEW
		Call Daily_report.ResetAttrs()
		Call Daily_report_report.RenderRow()
		Daily_report.Paymentdate.CurrentValue = Daily_report.Paymentdate.DbValue
%>
	
<%
	Daily_report.payment_gross.CurrentValue = Daily_report_report.ReportTotals(2,5)
	Daily_report.payment_fee.CurrentValue = Daily_report_report.ReportTotals(2,6)
	Daily_report.Tax.CurrentValue = Daily_report_report.ReportTotals(2,7)
	Daily_report.Shipping.CurrentValue = Daily_report_report.ReportTotals(2,8)

	' Render row for view
	Daily_report.RowType = EW_ROWTYPE_VIEW
	Call Daily_report.ResetAttrs()
	Call Daily_report_report.RenderRow()
%>
	<tr>
		<td class="ewGroupAggregate" colspan=5><span class="aspmaker"><span class="aspmaker">Sub total(<%= FormatNumber(Daily_report_report.ReportCounts(2),0) %>&nbsp;<%= Language.Phrase("RptDtlRec") %>)</span></span></td>

		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class=" ewGroupAggregate aspmaker">
<div<%= Daily_report.payment_gross.ViewAttributes %>><%= Daily_report.payment_gross.ViewValue %></div>
</span></td>
		<td><span class="ewGroupAggregate aspmaker">
<div<%= Daily_report.payment_fee.ViewAttributes %>><%= Daily_report.payment_fee.ViewValue %></div>
</span></td>
		<td><span class="ewGroupAggregate aspmaker">
<div<%= Daily_report.Tax.ViewAttributes %>><%= Daily_report.Tax.ViewValue %></div>
</span></td>
		<td><span class="ewGroupAggregate aspmaker">
<div<%= Daily_report.Shipping.ViewAttributes %>><%= Daily_report.Shipping.ViewValue %></div>
</span></td>
	</tr>
	<tr><td colspan=11><span class="aspmaker">&nbsp;<br></span></td></tr>
<%
End If
	If Daily_report_report.LevelBreak(1) Then
		Daily_report.Ship_Province.CurrentValue = Daily_report_report.ReportGroups(0)

		' Render row for view
		Daily_report.RowType = EW_ROWTYPE_VIEW
		Call Daily_report.ResetAttrs()
		Call Daily_report_report.RenderRow()
		Daily_report.Ship_Province.CurrentValue = Daily_report.Ship_Province.DbValue
%>
	<tr><td colspan=10 class="ewGroupSummary"><span class="aspmaker"><%= Language.Phrase("RptSumHead") %>&nbsp;<%= Daily_report.Ship_Province.FldCaption %>:&nbsp;<%= Daily_report.Ship_Province.ViewValue %>&nbsp;(<%= FormatNumber(Daily_report_report.ReportCounts(1),0) %>&nbsp;<%= Language.Phrase("RptDtlRec") %>)</span></td></tr>
<%
	Daily_report.payment_gross.CurrentValue = Daily_report_report.ReportTotals(1,5)
	Daily_report.payment_fee.CurrentValue = Daily_report_report.ReportTotals(1,6)
	Daily_report.Tax.CurrentValue = Daily_report_report.ReportTotals(1,7)
	Daily_report.Shipping.CurrentValue = Daily_report_report.ReportTotals(1,8)

	' Render row for view
	Daily_report.RowType = EW_ROWTYPE_VIEW
	Call Daily_report.ResetAttrs()
	Call Daily_report_report.RenderRow()
%>
	<tr>
		<td class="ewGroupAggregate"><span class="aspmaker"><%= Language.Phrase("RptSum") %></span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_gross.ViewAttributes %>><%= Daily_report.payment_gross.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_fee.ViewAttributes %>><%= Daily_report.payment_fee.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Tax.ViewAttributes %>><%= Daily_report.Tax.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Shipping.ViewAttributes %>><%= Daily_report.Shipping.ViewValue %></div>
</span></td>
	</tr>
	<tr><td colspan=11><span class="aspmaker">&nbsp;<br></span></td></tr>
<%
End If
Loop

' Close recordset
Daily_report_report.Recordset.Close
Set Daily_report_report.Recordset = Nothing
%>
	<tr><td colspan=10><span class="aspmaker">&nbsp;<br></span></td></tr>
	<tr><td colspan=10 class="ewGrandSummary"><span class="aspmaker"><%= Language.Phrase("RptGrandTotal") %>&nbsp;(<%= FormatNumber(Daily_report_report.ReportCounts(0),0) %>&nbsp;<%= Language.Phrase("RptDtlRec") %>)</span></td></tr>
<%
	Daily_report.payment_gross.CurrentValue = Daily_report_report.ReportTotals(0,5)
	Daily_report.payment_fee.CurrentValue = Daily_report_report.ReportTotals(0,6)
	Daily_report.Tax.CurrentValue = Daily_report_report.ReportTotals(0,7)
	Daily_report.Shipping.CurrentValue = Daily_report_report.ReportTotals(0,8)

	' Render row for view
	Daily_report.RowType = EW_ROWTYPE_VIEW
	Call Daily_report.ResetAttrs()
	Call Daily_report_report.RenderRow()
%>
	<tr>
		<td class="ewGroupAggregate"><span class="aspmaker"><%= Language.Phrase("RptSum") %></span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">&nbsp;</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_gross.ViewAttributes %>><%= Daily_report.payment_gross.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.payment_fee.ViewAttributes %>><%= Daily_report.payment_fee.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Tax.ViewAttributes %>><%= Daily_report.Tax.ViewValue %></div>
</span></td>
		<td><span class="aspmaker">
<div<%= Daily_report.Shipping.ViewAttributes %>><%= Daily_report.Shipping.ViewValue %></div>
</span></td>
	</tr>
	<tr><td colspan=10><span class="aspmaker">&nbsp;<br></span></td></tr>
</table>
</form>
<%
Daily_report_report.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Daily_report.Export = "" Then %>
<script language="JavaScript" type="text/javascript">
<!--
// Write your table-specific startup script here
// document.write("page loaded");
//-->
</script>
<% End If %>
<!--#include file="footer.asp"-->
<%

' Drop page object
Set Daily_report_report = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cDaily_report_report

	' Page ID
	Public Property Get PageID()
		PageID = "report"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "Daily report"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Daily_report_report"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
	End Property

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

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
		Call Message_Showing(sMessage, "")
		If sMessage <> "" Then Response.Write "<p class=""ewMessage"">" & sMessage & "</p>"
		Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then Response.Write "<p class=""ewSuccessMessage"">" & sSuccessMessage & "</p>"
		Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then Response.Write "<p class=""ewErrorMessage"">" & sErrorMessage & "</p>"
		Session(EW_SESSION_FAILURE_MESSAGE) = "" ' Clear message in Session
	End Sub
	Dim PageHeader
	Dim PageFooter

	' Show Page Header
	Public Sub ShowPageHeader()
		Dim sHeader
		sHeader = PageHeader
		Call Page_DataRendering(sHeader)
		If sHeader <> "" Then ' Header exists, display
			Response.Write "<p class=""aspmaker"">" & sHeader & "</p>"
		End If
	End Sub

	' Show Page Footer
	Public Sub ShowPageFooter()
		Dim sFooter
		sFooter = PageFooter
		Call Page_DataRendered(sFooter)
		If sFooter <> "" Then ' Footer exists, display
			Response.Write "<p class=""aspmaker"">" & sFooter & "</p>"
		End If
	End Sub

	' -----------------------
	'  Validate Page request
	'
	Public Function IsPageRequest()
		IsPageRequest = True
	End Function

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

		' Initialize table object
		If IsEmpty(Daily_report) Then Set Daily_report = New cDaily_report
		Set Table = Daily_report

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"

		' Initialize other table object
		If IsEmpty(Logins) Then Set Logins = New cLogins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "report"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Daily report"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.Tag = "span"
		ExportOptions.Separator = "&nbsp;&nbsp;"
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
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("login.asp")
		End If

		' Get export parameters
		If Request.QueryString("export").Count > 0 Then
			Daily_report.Export = Request.QueryString("export")
		End If
		gsExport = Daily_report.Export ' Get export parameter, used in header
		gsExportFile = Daily_report.TableVar ' Get export file, used in header
		Dim Charset ' Charset used in header
		If EW_CHARSET <> "" Then
			Charset = ";charset=" & EW_CHARSET
		Else
			Charset = ""
		End If

		' Global page loading event (in userfn7.asp)
		Call Page_Loading()

		' Page load event, used in current page
		Call Page_Load()
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

		' Page unload event, used in current page
		Call Page_Unload()

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set ObjForm = Nothing

		' Go to url if specified
		If sReDirectUrl <> "" Then
			Response.Redirect sReDirectUrl
		End If
	End Sub

	'
	'  Subroutine Page_Terminate (End)
	' ----------------------------------------

	Dim ExportOptions ' Export options
	Dim RecCnt
	Dim ReportSql
	Dim ReportFilter
	Dim DefaultFilter
	Dim DbMasterFilter
	Dim DbDetailFilter
	Dim MasterRecordExists
	Dim DtlRecordCount
	Dim ReportGroups
	Dim ReportCounts
	Dim LevelBreak
	Dim ReportTotals
	Dim ReportMaxs
	Dim ReportMins
	Dim Recordset
	Dim DetailRecordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		ReDim ReportGroups(2)
		ReDim ReportCounts(2)
		ReDim LevelBreak(2)
		ReDim ReportTotals(2, 9)
		ReDim ReportMaxs(2, 9)
		ReDim ReportMins(2, 9)
	End Sub

	' -----------------------------------------------------------------
	' Check level break
	'
	Sub ChkLvlBreak()
		LevelBreak(1) = False
		LevelBreak(2) = False
		If RecCnt = 0 Then ' Start Or End of Recordset
			LevelBreak(1) = True
			LevelBreak(2) = True
		Else
			If Not ew_CompareValue(Daily_report.Ship_Province.CurrentValue, ReportGroups(0)) Then
				LevelBreak(1) = True
				LevelBreak(2) = True
			End If
			If Not ew_CompareValue(Daily_report.Paymentdate.CurrentValue, ReportGroups(1)) Then
				LevelBreak(2) = True
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call Daily_report.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' InvoiceId
		' CustomerId
		' Inv_FirstName
		' Inv_LastName
		' inv_EmailAddress
		' payment_gross
		' payment_fee
		' Tax
		' Shipping
		' Ship_Province
		' Paymentdate
		' payment_status
		' -----------
		'  View  Row
		' -----------

		If Daily_report.RowType = EW_ROWTYPE_VIEW Then ' View row

			' InvoiceId
			Daily_report.InvoiceId.ViewValue = Daily_report.InvoiceId.CurrentValue
			Daily_report.InvoiceId.ViewCustomAttributes = ""

			' CustomerId
			Daily_report.CustomerId.ViewValue = Daily_report.CustomerId.CurrentValue
			Daily_report.CustomerId.ViewCustomAttributes = ""

			' Inv_FirstName
			Daily_report.Inv_FirstName.ViewValue = Daily_report.Inv_FirstName.CurrentValue
			Daily_report.Inv_FirstName.ViewCustomAttributes = ""

			' Inv_LastName
			Daily_report.Inv_LastName.ViewValue = Daily_report.Inv_LastName.CurrentValue
			Daily_report.Inv_LastName.ViewCustomAttributes = ""

			' inv_EmailAddress
			Daily_report.inv_EmailAddress.ViewValue = Daily_report.inv_EmailAddress.CurrentValue
			Daily_report.inv_EmailAddress.ViewCustomAttributes = ""

			' payment_gross
			Daily_report.payment_gross.ViewValue = Daily_report.payment_gross.CurrentValue
			Daily_report.payment_gross.ViewCustomAttributes = ""

			' payment_fee
			Daily_report.payment_fee.ViewValue = Daily_report.payment_fee.CurrentValue
			Daily_report.payment_fee.ViewCustomAttributes = ""

			' Tax
			Daily_report.Tax.ViewValue = Daily_report.Tax.CurrentValue
			Daily_report.Tax.ViewCustomAttributes = ""

			' Shipping
			Daily_report.Shipping.ViewValue = Daily_report.Shipping.CurrentValue
			Daily_report.Shipping.ViewCustomAttributes = ""

			' Ship_Province
			Daily_report.Ship_Province.ViewValue = Daily_report.Ship_Province.CurrentValue
			Daily_report.Ship_Province.ViewCustomAttributes = ""

			' Paymentdate
			Daily_report.Paymentdate.ViewValue = Daily_report.Paymentdate.CurrentValue
			Daily_report.Paymentdate.ViewCustomAttributes = ""

			' payment_status
			Daily_report.payment_status.ViewValue = Daily_report.payment_status.CurrentValue
			Daily_report.payment_status.ViewCustomAttributes = ""

			' View refer script
			' InvoiceId

			Daily_report.InvoiceId.LinkCustomAttributes = ""
			Daily_report.InvoiceId.HrefValue = ""
			Daily_report.InvoiceId.TooltipValue = ""

			' CustomerId
			Daily_report.CustomerId.LinkCustomAttributes = ""
			Daily_report.CustomerId.HrefValue = ""
			Daily_report.CustomerId.TooltipValue = ""

			' Inv_FirstName
			Daily_report.Inv_FirstName.LinkCustomAttributes = ""
			Daily_report.Inv_FirstName.HrefValue = ""
			Daily_report.Inv_FirstName.TooltipValue = ""

			' Inv_LastName
			Daily_report.Inv_LastName.LinkCustomAttributes = ""
			Daily_report.Inv_LastName.HrefValue = ""
			Daily_report.Inv_LastName.TooltipValue = ""

			' inv_EmailAddress
			Daily_report.inv_EmailAddress.LinkCustomAttributes = ""
			Daily_report.inv_EmailAddress.HrefValue = ""
			Daily_report.inv_EmailAddress.TooltipValue = ""

			' payment_gross
			Daily_report.payment_gross.LinkCustomAttributes = ""
			Daily_report.payment_gross.HrefValue = ""
			Daily_report.payment_gross.TooltipValue = ""

			' payment_fee
			Daily_report.payment_fee.LinkCustomAttributes = ""
			Daily_report.payment_fee.HrefValue = ""
			Daily_report.payment_fee.TooltipValue = ""

			' Tax
			Daily_report.Tax.LinkCustomAttributes = ""
			Daily_report.Tax.HrefValue = ""
			Daily_report.Tax.TooltipValue = ""

			' Shipping
			Daily_report.Shipping.LinkCustomAttributes = ""
			Daily_report.Shipping.HrefValue = ""
			Daily_report.Shipping.TooltipValue = ""

			' Paymentdate
			Daily_report.Paymentdate.LinkCustomAttributes = ""
			Daily_report.Paymentdate.HrefValue = ""
			Daily_report.Paymentdate.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Daily_report.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Daily_report.Row_Rendered()
		End If
	End Sub

	' Page Load event
	Sub Page_Load()

		'Response.Write "Page Load"
	End Sub

	' Page Unload event
	Sub Page_Unload()

		'Response.Write "Page Unload"
	End Sub

	' Page Redirecting event
	Sub Page_Redirecting(url)

		'url = newurl
	End Sub

	' Message Showing event
	' typ = ""|"success"|"failure"
	Sub Message_Showing(msg, typ)

		' Example:
		'If typ = "success" Then msg = "your success message"

	End Sub

	' Page Data Rendering event
	Sub Page_DataRendering(header)

		' Example:
		'header = "your header"

	End Sub

	' Page Data Rendered event
	Sub Page_DataRendered(footer)

		' Example:
		'footer = "your footer"

	End Sub
End Class
%>
