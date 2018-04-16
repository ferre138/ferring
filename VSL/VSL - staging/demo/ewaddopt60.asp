<%@ EnableSessionState=False %>
<%
Const EW_PAGE_ID = "ewaddopt"
%>
<!--#include file="ewcfg60.asp"-->
<!--#include file="aspfn60.asp"-->
<%
On Error Resume Next
Dim LeftQuote, RightQuote, QS, Sql, Where, FieldList, ValueList
Dim LookupTableName, LinkFieldName, DisplayFieldName, DisplayField2Name
Dim LinkField, DisplayField, DisplayField2, LinkFieldQuote, DisplayFieldQuote, DisplayField2Quote
Dim bError
Dim bUseLinkField, bUseDisplayField, bUseDisplayField2
LeftQuote = "["
RightQuote = "]"
bError = False
QS = Split(Request.Querystring, "&")
If IsArray(QS) Then
	If UBound(QS) >= 0 Then
		LookupTableName = GetValue("ltn")

		' Parent fields
		Dim ParentFieldName, ParentField, ParentFieldQuote
		ParentFieldName = GetValue("pfn")
		ParentField = GetValue("pf")
		ParentFieldQuote = GetValue("pfq")

		' Link and display fields
		LinkFieldName = GetValue("lfn")
		DisplayFieldName = GetValue("dfn")
		DisplayField2Name = GetValue("df2n")
		LinkField = GetValue("lf")
		If DisplayFieldName = LinkFieldName Then
			DisplayField = LinkField
		Else
			DisplayField = GetValue("df")
		End If
		If DisplayField2Name = LinkFieldName Then
			DisplayField2 = LinkField
		ElseIf DisplayField2Name = DisplayFieldName Then
			DisplayField2 = DisplayField
		Else
			DisplayField2 = GetValue("df2")
		End If
		LinkFieldQuote = GetValue("lfq")
		DisplayFieldQuote = GetValue("dfq")
		DisplayField2Quote = GetValue("df2q")
	Else
		Response.Write "Invalid Parameter"
		Response.End
	End If
Else
	Response.Write "Invalid Parameter"
	Response.End
End If
If LookupTableName = "" Then
	Response.Write "Missing lookup table name"
	Response.End
End If
If DisplayFieldName = "" Then
	Response.Write "Missing display field name"
	Response.End
End If
Dim bUseParentField
bUseParentField = (ParentFieldName <> "" And ParentField <> "")
bUseLinkField = (LinkFieldName <> "" And LinkField <> "")
bUseDisplayField = (DisplayFieldName <> "" And DisplayFieldName <> LinkFieldName)
If bUseDisplayField Then
	If bUseParentField And ParentFieldName = DisplayFieldName Then
		DisplayField = ParentField
	End If
	bUseDisplayField = (DisplayField <> "")
End If
bUseDisplayField2 = (DisplayField2Name <> "" And DisplayField2Name <> LinkFieldName And DisplayField2Name <> DisplayFieldName)
If bUseDisplayField2 Then
	If bUseParentField And ParentFieldName = DisplayField2Name Then
		DisplayField2 = ParentField
	End If
	bUseDisplayField2 = (DisplayField2 <> "")
End If
Sql = ""
If bUseLinkField Then
	Sql = Sql & LeftQuote & LinkFieldName & RightQuote
End If
If bUseDisplayField Then
	If Sql <> "" Then Sql = Sql & ","
	Sql = Sql & LeftQuote & DisplayFieldName & RightQuote
End If
If bUseDisplayField2 Then
	If Sql <> "" Then Sql = Sql & ","
	Sql = Sql & LeftQuote & DisplayField2Name & RightQuote
End If
Sql = "SELECT DISTINCT " & Sql & " FROM " & LeftQuote & LookupTableName & RightQuote
Where = ""
If bUseLinkField Then
	Where = LeftQuote & LinkFieldName & RightQuote & "=" & LinkFieldQuote & ew_AdjustSql(LinkField) & LinkFieldQuote
End If
If bUseDisplayField Then
	If Where <> "" Then Where = Where & " AND "
	Where = Where & LeftQuote & DisplayFieldName & RightQuote & "=" & DisplayFieldQuote & ew_AdjustSql(DisplayField) & DisplayFieldQuote
End If
If bUseDisplayField2 Then
	If Where <> "" Then Where = Where & " AND "
	Where = Where & LeftQuote & DisplayField2Name & RightQuote & "=" & DisplayField2Quote & ew_AdjustSql(DisplayField2) & DisplayField2Quote
End If
Sql = Sql & " WHERE " & Where
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
Set rs = conn.Execute(Sql)
If Err.Number <> 0 Then
	Response.Write Err.Description
	bError = True
End If
If Not bError Then
	If rs.Eof Then ' Add new option
		FieldList = ""
		ValueList = ""
		If bUseParentField Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & ParentFieldName & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & ParentFieldQuote & ew_AdjustSql(ParentField) & ParentFieldQuote
		End If
		If bUseLinkField Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & LinkFieldName & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & LinkFieldQuote & ew_AdjustSql(LinkField) & LinkFieldQuote
		End If
		If bUseDisplayField And ParentFieldName <> DisplayFieldName Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & DisplayFieldName & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & DisplayFieldQuote & ew_AdjustSql(DisplayField) & DisplayFieldQuote
		End If
		If bUseDisplayField2 And ParentFieldName <> DisplayField2Name Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & DisplayField2Name & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & DisplayField2Quote & ew_AdjustSql(DisplayField2) & DisplayField2Quote
		End If
		conn.Execute("INSERT INTO " & LeftQuote & LookupTableName & RightQuote & " (" & FieldList & ") VALUES (" & ValueList & ")")
		If Err.Number <> 0 Then
			Response.Write Err.Description
			bError = True
		End If
	Else
		Response.Write "Option already exists"
		bError = True
	End If
End If
rs.Close
Set rs = Nothing
If Not bError Then
	If LinkField = "" Then ' Get new link field value
		Sql = "SELECT " & LeftQuote & LinkFieldName & RightQuote & " FROM " & LeftQuote & LookupTableName & RightQuote & " WHERE " & Where
		Set rs = conn.Execute(Sql)
		If Not rs.Eof Then
			LinkField = rs(0)
			If DisplayFieldName = LinkFieldName Then DisplayField = LinkField
			If DisplayField2Name = LinkFieldName Then DisplayField2 = LinkField
		End If
		rs.Close
		Set rs = Nothing
	End If
End If
conn.Close
Set conn = Nothing
If Not bError Then
	Response.Clear
	Response.Write "OK" & vbCr
	Response.Write LinkField & vbCr
	Response.Write DisplayField & vbCr
	Response.Write DisplayField2
End If
Response.End

Function GetValue(Key)
	Dim kv, i
	For i = 0 To UBound(QS)
		kv = Split(QS(i), "=")
		If (kv(0) = Key) Then
			GetValue = ew_Decode(kv(1))
			Exit Function
		End If
	Next
	GetValue = ""
End Function
%>
