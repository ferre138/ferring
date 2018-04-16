<%@ EnableSessionState=False %>
<%
Const EW_PAGE_ID = "ewlookup"
%>
<!--#include file="ewcfg60.asp"-->
<!--#include file="aspfn60.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%
Dim QS, Sql, Value, LnkType, LnkCount, LnkFld, LnkDisp1, LnkDisp2
Dim I, arValue, LnkFldType
QS = Split(Request.Querystring, "&")
If IsArray(QS) Then
	If UBound(QS) >= 0 Then
		Sql = GetValue("s")
		Sql = TEAdecrypt(Sql, EW_RANDOM_KEY)
		Value = GetValue("q")
		Value = ew_AdjustSql(Value)
		LnkType = GetValue("lt") ' Get link type
		If LnkType = "2" Then ' Auto fill
			LnkCount = 1
			LnkFld = -1
			LnkDisp1 = 0
			LnkDisp2 = -1
		ElseIf LnkType = "1" Then ' Auto suggest
			LnkCount = 2
			LnkFld = -1
			LnkDisp1 = 0
			LnkDisp2 = 1
		Else
			LnkCount = GetValue("lc") ' Link field count
			If Not IsNumeric(LnkCount) Then
				Response.End
			ElseIf CInt(LnkCount) <= 0 Then
				Response.End
			End If
			LnkFld = 0 ' Link field default = 0
			LnkDisp1 = GetValue("ld1") ' Link display field
			If Not IsNumeric(LnkDisp1) Then
				Response.End
			ElseIf CInt(LnkDisp1) < -1 Or CInt(LnkDisp1) >= CInt(LnkCount) Then
				Response.End
			End If
			LnkDisp2 = GetValue("ld2") ' Link display field 2
			If Not IsNumeric(LnkDisp2) Then
				Response.End
			ElseIf CInt(LnkDisp2) < -1 Or CInt(LnkDisp2) >= CInt(LnkCount) Then
				Response.End
			End If
			LnkFldType = GetValue("lft") ' Link field data type
		End If
		If Sql <> "" Then
			If Value <> "" Then
				arValue = Split(Value, ",")
				For I = 0 To UBound(arValue)
					arValue(I) = ew_QuotedValue(arValue(I), LnkFldType)
				Next
				Sql = Replace(Sql, "@FILTER_VALUE", Join(arValue, ","))
			End If
			GetLookupValues(Sql)
		End If
	End If
End If

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

Sub GetLookupValues(Sql)

	' Connect to database
	Dim conn, rs, rsarr, str, i
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open EW_DB_CONNECTION_STRING
	Set rs = conn.Execute(Sql)
	If Not rs.EOF Then
		rsarr = rs.GetRows
	End If

	' Close database
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing

	' Output
	If IsArray(rsarr) Then
		If LnkType = "2" Then ' Auto fill
			i = 0
			Do While i < UBound(rsarr, 1)
				str = rsarr(i, 0)
				str = RemoveCrLf(str)
				If rsarr(i+1, 0)&"" <> "" Then
					str = str & ", " & RemoveCrLf(rsarr(i+1, 0)&"")
				End If
				Response.Write str & vbCr
				i = i + 2
			Loop
		Else
			For i = 0 To UBound(rsarr, 2)
				If UBound(rsarr, 1) = CInt(LnkCount) -1 Then

					' Process link field
					If LnkType <> "1" Then
						str = rsarr(LnkFld, i)
						str = RemoveCrLf(str)
						Response.Write str & vbCr
					End If

					' Process display field
					If CInt(LnkDisp1) >= 0 Then
						str = rsarr(LnkDisp1, i)
						str = RemoveCrLf(str)
					Else
						str = ""
					End If
					Response.Write str & vbCr

					' Process display field 2
					If CInt(LnkDisp2) >= 0 Then
						str = rsarr(LnkDisp2, i)
						str = RemoveCrLf(str)
					Else
						str = ""
					End If
					Response.Write str & vbCr
				End If
			Next
		End If
	End If
End Sub

Function RemoveCrLf(s)
	Dim wrkstr
	wrkstr = s
	If Len(wrkstr) > 0 Then
		wrkstr = Replace(wrkstr, vbCrLf, " ")
		wrkstr = Replace(wrkstr, vbCr, " ")
		wrkstr = Replace(wrkstr, vbLf, " ")
	End If
	RemoveCrLf = wrkstr
End Function
%>
