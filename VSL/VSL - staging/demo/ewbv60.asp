<%
Const EW_PAGE_ID = "ewbv.asp"
%>
<!--#include file="ewcfg60.asp"-->
<!--#include file="aspfn60.asp"-->
<%
Dim tbl, fld, ft, fn, fs, b, obj
Dim resize, width, height, interpolation
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"

' Get resize parameters
resize = Request.QueryString("resize").Count > 0
If Request.QueryString("width").Count > 0 Then
	width = Request.QueryString("width")
End If
If Request.QueryString("height").Count > 0 Then
	height = Request.QueryString("height")
End If
If Request.QueryString("width").Count <= 0 And Request.QueryString("height").Count <= 0 Then
	width = EW_THUMBNAIL_DEFAULT_WIDTH
	height = EW_THUMBNAIL_DEFAULT_HEIGHT
End If
If Request.QueryString("interpolation").Count > 0 Then
	interpolation = Request.QueryString("interpolation")
Else
	interpolation = EW_THUMBNAIL_DEFAULT_INTERPOLATION
End If

' Resize image from physical file
If Request.QueryString("fn").Count > 0 Then
	Dim fso
	fn = Request.QueryString("fn")
	fn = Server.MapPath(fn)
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	If fso.FileExists(fn) Then
		Response.BinaryWrite ew_ResizeFileToBinary(fn, width, height, interpolation)
	End If
	Set fso = Nothing
	Response.End

' Display image from Session
Else
	If Request.QueryString("tbl").Count > 0 Then
		tbl = Request.QueryString("tbl")
	Else
		Response.End
	End If
	If Request.QueryString("fld").Count > 0 Then
		fld = Request.QueryString("fld")
	Else
		Response.End
	End If

	' Get blob field
	Set obj = New cUpload
	obj.TblVar = tbl
	obj.FldVar = fld
	obj.RestoreFromSession()
	b = obj.Value
	If IsNull(b) Then Response.End
	ft = obj.ContentType
	fn = obj.FileName

	'If ft <> "" Then
	'	Response.ContentType = ft
	'End If
	'If fn <> "" Then
	'	Response.AddHeader "Content-Disposition", "attachment; filename=""" & fn & """"
	'End If

	If resize Then Call obj.Resize(width, height, interpolation)
	Response.BinaryWrite obj.Value
	Set obj = Nothing
	Response.End
End If
%>
