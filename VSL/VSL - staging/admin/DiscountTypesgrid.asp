<% Session.Timeout = 20 %>
<%

' Define page object
Dim DiscountTypes_grid
Set DiscountTypes_grid = New cDiscountTypes_grid
Set MasterPage = Page
Set Page = DiscountTypes_grid

' Page init processing
Call DiscountTypes_grid.Page_Init()

' Page main processing
Call DiscountTypes_grid.Page_Main()
%>
<% If DiscountTypes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var DiscountTypes_grid = new ew_Page("DiscountTypes_grid");
// page properties
DiscountTypes_grid.PageID = "grid"; // page ID
DiscountTypes_grid.FormID = "fDiscountTypesgrid"; // form ID
var EW_PAGE_ID = DiscountTypes_grid.PageID; // for backward compatibility
// extend page with ValidateForm function
DiscountTypes_grid.ValidateForm = function(fobj) {
	ew_PostAutoSuggest(fobj);
	if (!this.ValidateRequired)
		return true; // ignore validation
	if (fobj.a_confirm && fobj.a_confirm.value == "F")
		return true;
	var i, elm, aelm, infix;
	var rowcnt = (fobj.key_count) ? Number(fobj.key_count.value) : 1;
	var addcnt = 0;
	for (i=0; i<rowcnt; i++) {
		infix = (fobj.key_count) ? String(i+1) : "";
		var chkthisrow = true;
		if (fobj.a_list && fobj.a_list.value == "gridinsert")
			chkthisrow = !(this.EmptyRow(fobj, infix));
		else
			chkthisrow = true;
		if (chkthisrow) {
			addcnt += 1;
		elm = fobj.elements["x" + infix + "_FreePerQty"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.FreePerQty.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_SpecialPrice"];
		if (elm && !ew_CheckNumber(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(DiscountTypes.SpecialPrice.FldErrMsg) %>");
		// Set up row object
		var row = {};
		row["index"] = infix;
		for (var j = 0; j < fobj.elements.length; j++) {
			var el = fobj.elements[j];
			var len = infix.length + 2;
			if (el.name.substr(0, len) == "x" + infix + "_") {
				var elname = "x_" + el.name.substr(len);
				if (ewLang.isObject(row[elname])) { // already exists
					if (ewLang.isArray(row[elname])) {
						row[elname][row[elname].length] = el; // add to array
					} else {
						row[elname] = [row[elname], el]; // convert to array
					}
				} else {
					row[elname] = el;
				}
			}
		}
		fobj.row = row;
		// Call Form Custom Validate event
		if (!this.Form_CustomValidate(fobj)) return false;
		} // End Grid Add checking
	}
	return true;
}
// Extend page with empty row check
DiscountTypes_grid.EmptyRow = function(fobj, infix) {
	if (ew_ValueChanged(fobj, infix, "DiscountType")) return false;
	if (ew_ValueChanged(fobj, infix, "DiscountTitle")) return false;
	if (ew_ValueChanged(fobj, infix, "freeShipping")) return false;
	if (ew_ValueChanged(fobj, infix, "FreePerQty")) return false;
	if (ew_ValueChanged(fobj, infix, "SpecialPrice")) return false;
	return true;
}
// extend page with Form_CustomValidate function
DiscountTypes_grid.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
<% If EW_CLIENT_VALIDATE Then %>
DiscountTypes_grid.ValidateRequired = true; // uses JavaScript validation
<% Else %>
DiscountTypes_grid.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<% End If %>
<% DiscountTypes_grid.ShowPageHeader() %>
<%
If DiscountTypes.CurrentAction = "gridadd" Then
	If DiscountTypes.CurrentMode <> "copy" Then DiscountTypes.CurrentFilter = "0=1"
End If

' Load recordset
Set DiscountTypes_grid.Recordset = DiscountTypes_grid.LoadRecordset()
If DiscountTypes.CurrentAction = "gridadd" Then
	If DiscountTypes.CurrentMode = "copy" Then
		DiscountTypes_grid.TotalRecs = DiscountTypes_grid.Recordset.RecordCount
		DiscountTypes_grid.StartRec = 1
		DiscountTypes_grid.DisplayRecs = DiscountTypes_grid.TotalRecs
	Else
		DiscountTypes_grid.StartRec = 1
		DiscountTypes_grid.DisplayRecs = DiscountTypes.GridAddRowCount
	End If
	DiscountTypes_grid.TotalRecs = DiscountTypes_grid.DisplayRecs
	DiscountTypes_grid.StopRec = DiscountTypes_grid.DisplayRecs
Else
	DiscountTypes_grid.TotalRecs = DiscountTypes_grid.Recordset.RecordCount
	DiscountTypes_grid.StartRec = 1
	DiscountTypes_grid.DisplayRecs = DiscountTypes_grid.TotalRecs ' Display all records
End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><% If DiscountTypes.CurrentMode = "add" Or DiscountTypes.CurrentMode = "copy" Then %><%= Language.Phrase("Add") %><% ElseIf DiscountTypes.CurrentMode = "edit" Then %><%= Language.Phrase("Edit") %><% End If %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= DiscountTypes.TableCaption %></p>
</p>
<% DiscountTypes_grid.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If (DiscountTypes.CurrentMode = "add" Or DiscountTypes.CurrentMode = "copy" Or DiscountTypes.CurrentMode = "edit") And DiscountTypes.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridUpperPanel">
<% If DiscountTypes.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
<div id="gmp_DiscountTypes" class="ewGridMiddlePanel">
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= DiscountTypes.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call DiscountTypes_grid.RenderListOptions()

' Render list options (header, left)
DiscountTypes_grid.ListOptions.Render "header", "left"
%>
<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
	<% If DiscountTypes.SortUrl(DiscountTypes.DiscountType) = "" Then %>
		<td><%= DiscountTypes.DiscountType.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.DiscountType.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.DiscountType.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.DiscountType.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
	<% If DiscountTypes.SortUrl(DiscountTypes.DiscountTitle) = "" Then %>
		<td><%= DiscountTypes.DiscountTitle.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.DiscountTitle.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.DiscountTitle.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.DiscountTitle.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
	<% If DiscountTypes.SortUrl(DiscountTypes.freeShipping) = "" Then %>
		<td><%= DiscountTypes.freeShipping.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.freeShipping.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.freeShipping.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.freeShipping.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
	<% If DiscountTypes.SortUrl(DiscountTypes.FreePerQty) = "" Then %>
		<td><%= DiscountTypes.FreePerQty.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.FreePerQty.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.FreePerQty.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.FreePerQty.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
	<% If DiscountTypes.SortUrl(DiscountTypes.SpecialPrice) = "" Then %>
		<td><%= DiscountTypes.SpecialPrice.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= DiscountTypes.SpecialPrice.FldCaption %></td><td style="width: 10px;"><% If DiscountTypes.SpecialPrice.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf DiscountTypes.SpecialPrice.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
DiscountTypes_grid.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
DiscountTypes_grid.StartRec = 1
DiscountTypes_grid.StopRec = DiscountTypes_grid.TotalRecs ' Show all records

' Restore number of post back records
If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
	ObjForm.Index = 0
	If ObjForm.HasValue("key_count") And (DiscountTypes.CurrentAction = "gridadd" Or DiscountTypes.CurrentAction = "gridedit" Or DiscountTypes.CurrentAction = "F") Then
		DiscountTypes_grid.KeyCount = ObjForm.GetValue("key_count")
		DiscountTypes_grid.StopRec = DiscountTypes_grid.KeyCount
	End If
End If

' Move to first record
DiscountTypes_grid.RecCnt = DiscountTypes_grid.StartRec - 1
If Not DiscountTypes_grid.Recordset.Eof Then
	DiscountTypes_grid.Recordset.MoveFirst
	If DiscountTypes_grid.StartRec > 1 Then DiscountTypes_grid.Recordset.Move DiscountTypes_grid.StartRec - 1
ElseIf Not DiscountTypes.AllowAddDeleteRow And DiscountTypes_grid.StopRec = 0 Then
	DiscountTypes_grid.StopRec = DiscountTypes.GridAddRowCount
End If

' Initialize Aggregate
DiscountTypes.RowType = EW_ROWTYPE_AGGREGATEINIT
Call DiscountTypes.ResetAttrs()
Call DiscountTypes_grid.RenderRow()
DiscountTypes_grid.RowCnt = 0
If DiscountTypes.CurrentAction = "gridadd" Then DiscountTypes_grid.RowIndex = 0
If DiscountTypes.CurrentAction = "gridedit" Then DiscountTypes_grid.RowIndex = 0

' Output date rows
Do While CLng(DiscountTypes_grid.RecCnt) < CLng(DiscountTypes_grid.StopRec)
	DiscountTypes_grid.RecCnt = DiscountTypes_grid.RecCnt + 1
	If CLng(DiscountTypes_grid.RecCnt) >= CLng(DiscountTypes_grid.StartRec) Then
		DiscountTypes_grid.RowCnt = DiscountTypes_grid.RowCnt + 1
		If DiscountTypes.CurrentAction = "gridadd" Or DiscountTypes.CurrentAction = "gridedit" Or DiscountTypes.CurrentAction = "F" Then
			DiscountTypes_grid.RowIndex = DiscountTypes_grid.RowIndex + 1
			ObjForm.Index = DiscountTypes_grid.RowIndex
			If ObjForm.HasValue("k_action") Then
				DiscountTypes_grid.RowAction = ObjForm.GetValue("k_action") & ""
			ElseIf DiscountTypes.CurrentAction = "gridadd" Then
				DiscountTypes_grid.RowAction = "insert"
			Else
				DiscountTypes_grid.RowAction = ""
			End If
		End If

	' Set up key count
	DiscountTypes_grid.KeyCount = DiscountTypes_grid.RowIndex
	Call DiscountTypes.ResetAttrs()
	DiscountTypes.CssClass = ""
	If DiscountTypes.CurrentAction = "gridadd" Then
		If DiscountTypes.CurrentMode = "copy" Then
			Call DiscountTypes_grid.LoadRowValues(DiscountTypes_grid.Recordset) ' Load row values
			DiscountTypes_grid.RowOldKey = DiscountTypes_grid.SetRecordKey(DiscountTypes_grid.Recordset) ' Set old record key
		Else
			Call DiscountTypes_grid.LoadDefaultValues() ' Load default values
			DiscountTypes_grid.RowOldKey = "" ' Clear old key value
		End If
	Else
		Call DiscountTypes_grid.LoadRowValues(DiscountTypes_grid.Recordset) ' Load row values
	End If
	DiscountTypes.RowType = EW_ROWTYPE_VIEW ' Render view
	If DiscountTypes.CurrentAction = "gridadd" Then ' Grid add
		DiscountTypes.RowType = EW_ROWTYPE_ADD ' Render add
	End If
	If DiscountTypes.CurrentAction = "gridadd" And DiscountTypes.EventCancelled Then ' Insert failed
		Call DiscountTypes_grid.RestoreCurrentRowFormValues(DiscountTypes_grid.RowIndex) ' Restore form values
	End If
	If DiscountTypes.CurrentAction = "gridedit" Then ' Grid edit
		If DiscountTypes.EventCancelled Then ' Update failed
			Call DiscountTypes_grid.RestoreCurrentRowFormValues(DiscountTypes_grid.RowIndex) ' Restore form values
		End If
		If DiscountTypes_grid.RowAction = "insert" Then
			DiscountTypes.RowType = EW_ROWTYPE_ADD ' Render add
		Else
			DiscountTypes.RowType = EW_ROWTYPE_EDIT ' Render edit
		End If
	End If
	If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit row
		DiscountTypes_grid.EditRowCnt = DiscountTypes_grid.EditRowCnt + 1
	End If
	If DiscountTypes.CurrentAction = "F" Then ' Confirm row
		Call DiscountTypes_grid.RestoreCurrentRowFormValues(DiscountTypes_grid.RowIndex) ' Restore form values
	End If
	If DiscountTypes.RowType = EW_ROWTYPE_ADD Or DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Add / Edit row
		If DiscountTypes.CurrentAction = "edit" Then
			DiscountTypes.RowAttrs.AddAttributes Array()
			DiscountTypes.CssClass = "ewTableEditRow"
		Else
			DiscountTypes.RowAttrs.AddAttributes Array()
		End If
		If Not IsEmpty(DiscountTypes_grid.RowIndex) Then
			DiscountTypes.RowAttrs.AddAttributes Array(Array("data-rowindex", DiscountTypes_grid.RowIndex), Array("id", "r" & DiscountTypes_grid.RowIndex & "_DiscountTypes"))
		End If
	Else
		DiscountTypes.RowAttrs.AddAttributes Array()
	End If

	' Render row
	Call DiscountTypes_grid.RenderRow()

	' Render list options
	Call DiscountTypes_grid.RenderListOptions()

	' Skip delete row / empty row for confirm page
	If DiscountTypes_grid.RowAction <> "delete" And DiscountTypes_grid.RowAction <> "insertdelete" And Not (DiscountTypes_grid.RowAction = "insert" And DiscountTypes.CurrentAction = "F" And DiscountTypes_grid.EmptyRow()) Then
%>
	<tr<%= DiscountTypes.RowAttributes %>>
<%

' Render list options (body, left)
DiscountTypes_grid.ListOptions.Render "body", "left"
%>
	<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
		<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" size="30" maxlength="255" value="<%= DiscountTypes.DiscountType.EditValue %>"<%= DiscountTypes.DiscountType.EditAttributes %>>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" value="<%= Server.HTMLEncode(DiscountTypes.DiscountType.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" size="30" maxlength="255" value="<%= DiscountTypes.DiscountType.EditValue %>"<%= DiscountTypes.DiscountType.EditAttributes %>>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ListViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" value="<%= Server.HTMLEncode(DiscountTypes.DiscountType.CurrentValue&"") %>">
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" value="<%= Server.HTMLEncode(DiscountTypes.DiscountType.OldValue&"") %>">
<% End If %>
<a name="<%= DiscountTypes_grid.PageObjName & "_row_" & DiscountTypes_grid.RowCnt %>" id="<%= DiscountTypes_grid.PageObjName & "_row_" & DiscountTypes_grid.RowCnt %>"></a>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountTypeId" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTypeId.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then %>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTypeId" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTypeId.CurrentValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" size="30" maxlength="255" value="<%= DiscountTypes.DiscountTitle.EditValue %>"<%= DiscountTypes.DiscountTitle.EditAttributes %>>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTitle.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" size="30" maxlength="255" value="<%= DiscountTypes.DiscountTitle.EditValue %>"<%= DiscountTypes.DiscountTitle.EditAttributes %>>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ListViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTitle.CurrentValue&"") %>">
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTitle.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
		<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="1"<%= selwrk %><%= DiscountTypes.freeShipping.EditAttributes %>>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="<%= Server.HTMLEncode(DiscountTypes.freeShipping.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="1"<%= selwrk %><%= DiscountTypes.freeShipping.EditAttributes %>>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="<%= Server.HTMLEncode(DiscountTypes.freeShipping.CurrentValue&"") %>">
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="<%= Server.HTMLEncode(DiscountTypes.freeShipping.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" size="30" value="<%= DiscountTypes.FreePerQty.EditValue %>"<%= DiscountTypes.FreePerQty.EditAttributes %>>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" value="<%= Server.HTMLEncode(DiscountTypes.FreePerQty.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" size="30" value="<%= DiscountTypes.FreePerQty.EditValue %>"<%= DiscountTypes.FreePerQty.EditAttributes %>>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ListViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" value="<%= Server.HTMLEncode(DiscountTypes.FreePerQty.CurrentValue&"") %>">
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" value="<%= Server.HTMLEncode(DiscountTypes.FreePerQty.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" size="30" value="<%= DiscountTypes.SpecialPrice.EditValue %>"<%= DiscountTypes.SpecialPrice.EditAttributes %>>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" value="<%= Server.HTMLEncode(DiscountTypes.SpecialPrice.OldValue&"") %>">
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" size="30" value="<%= DiscountTypes.SpecialPrice.EditValue %>"<%= DiscountTypes.SpecialPrice.EditAttributes %>>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ListViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" value="<%= Server.HTMLEncode(DiscountTypes.SpecialPrice.CurrentValue&"") %>">
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" value="<%= Server.HTMLEncode(DiscountTypes.SpecialPrice.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
<%

' Render list options (body, right)
DiscountTypes_grid.ListOptions.Render "body", "right"
%>
	</tr>
<% If DiscountTypes.RowType = EW_ROWTYPE_ADD Then %>
<% End If %>
<% If DiscountTypes.RowType = EW_ROWTYPE_EDIT Then %>
<% End If %>
<%
	End If
	End If ' End delete row checking
	If DiscountTypes.CurrentAction <> "gridadd" Or DiscountTypes.CurrentMode = "copy" Then
		If Not DiscountTypes_grid.Recordset.Eof Then DiscountTypes_grid.Recordset.MoveNext()
	End If
Loop
%>
<%
	If DiscountTypes.CurrentMode = "add" Or DiscountTypes.CurrentMode = "copy" Or DiscountTypes.CurrentMode = "edit" Then
		DiscountTypes_grid.RowIndex = "$rowindex$"
		DiscountTypes_grid.LoadDefaultValues()

		' Set row properties
		Call DiscountTypes.ResetAttrs()
		DiscountTypes.RowAttrs.AddAttributes Array()
		If Not IsEmpty(DiscountTypes_grid.RowIndex) Then
			DiscountTypes.RowAttrs.AddAttributes Array(Array("data-rowindex", DiscountTypes_grid.RowIndex), Array("id", "r" & DiscountTypes_grid.RowIndex & "_DiscountTypes"))
		End If
		DiscountTypes.RowType = EW_ROWTYPE_ADD

		' Render row
		Call DiscountTypes_grid.RenderRow()

		' Render list options
		Call DiscountTypes_grid.RenderListOptions()

		' Add id and class to the template row
		DiscountTypes.RowAttrs.UpdateAttribute "id", "r0_DiscountTypes"
		DiscountTypes.RowAttrs.AddAttribute "class", "ewTemplate", True
%>
	<tr<%= DiscountTypes.RowAttributes %>>
<%

' Render list options (body, left)
DiscountTypes_grid.ListOptions.Render "body", "left"
%>
	<% If DiscountTypes.DiscountType.Visible Then ' DiscountType %>
		<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<% If DiscountTypes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" size="30" maxlength="255" value="<%= DiscountTypes.DiscountType.EditValue %>"<%= DiscountTypes.DiscountType.EditAttributes %>>
<% Else %>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountType" value="<%= Server.HTMLEncode(DiscountTypes.DiscountType.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountType" value="<%= Server.HTMLEncode(DiscountTypes.DiscountType.OldValue&"") %>">
</td>
	<% End If %>
	<% If DiscountTypes.DiscountTitle.Visible Then ' DiscountTitle %>
		<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<% If DiscountTypes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" size="30" maxlength="255" value="<%= DiscountTypes.DiscountTitle.EditValue %>"<%= DiscountTypes.DiscountTitle.EditAttributes %>>
<% Else %>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="x<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTitle.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" id="o<%= DiscountTypes_grid.RowIndex %>_DiscountTitle" value="<%= Server.HTMLEncode(DiscountTypes.DiscountTitle.OldValue&"") %>">
</td>
	<% End If %>
	<% If DiscountTypes.freeShipping.Visible Then ' freeShipping %>
		<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If DiscountTypes.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="1"<%= selwrk %><%= DiscountTypes.freeShipping.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="x<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="<%= Server.HTMLEncode(DiscountTypes.freeShipping.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" id="o<%= DiscountTypes_grid.RowIndex %>_freeShipping" value="<%= Server.HTMLEncode(DiscountTypes.freeShipping.OldValue&"") %>">
</td>
	<% End If %>
	<% If DiscountTypes.FreePerQty.Visible Then ' FreePerQty %>
		<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<% If DiscountTypes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" size="30" value="<%= DiscountTypes.FreePerQty.EditValue %>"<%= DiscountTypes.FreePerQty.EditAttributes %>>
<% Else %>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="x<%= DiscountTypes_grid.RowIndex %>_FreePerQty" value="<%= Server.HTMLEncode(DiscountTypes.FreePerQty.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" id="o<%= DiscountTypes_grid.RowIndex %>_FreePerQty" value="<%= Server.HTMLEncode(DiscountTypes.FreePerQty.OldValue&"") %>">
</td>
	<% End If %>
	<% If DiscountTypes.SpecialPrice.Visible Then ' SpecialPrice %>
		<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<% If DiscountTypes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" size="30" value="<%= DiscountTypes.SpecialPrice.EditValue %>"<%= DiscountTypes.SpecialPrice.EditAttributes %>>
<% Else %>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ViewValue %></div>
<input type="hidden" name="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="x<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" value="<%= Server.HTMLEncode(DiscountTypes.SpecialPrice.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" id="o<%= DiscountTypes_grid.RowIndex %>_SpecialPrice" value="<%= Server.HTMLEncode(DiscountTypes.SpecialPrice.OldValue&"") %>">
</td>
	<% End If %>
<%

' Render list options (body, right)
DiscountTypes_grid.ListOptions.Render "body", "right"
%>
	</tr>
<%
End If
%>
</tbody>
</table>
<% If DiscountTypes.CurrentMode = "add" Or DiscountTypes.CurrentMode = "copy" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridinsert">
<input type="hidden" name="key_count" id="key_count" value="<%= DiscountTypes_grid.KeyCount %>">
<%= DiscountTypes_grid.MultiSelectKey %>
<% End If %>
<% If DiscountTypes.CurrentMode = "edit" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridupdate">
<input type="hidden" name="key_count" id="key_count" value="<%= DiscountTypes_grid.KeyCount %>">
<%= DiscountTypes_grid.MultiSelectKey %>
<% End If %>
<input type="hidden" name="detailpage" id="detailpage" value="DiscountTypes_grid">
</div>
<%

' Close recordset and connection
DiscountTypes_grid.Recordset.Close
Set DiscountTypes_grid.Recordset = Nothing
%>
<% If (DiscountTypes.CurrentMode = "add" Or DiscountTypes.CurrentMode = "copy" Or DiscountTypes.CurrentMode = "edit") And DiscountTypes.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridLowerPanel">
<% If DiscountTypes.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
</td></tr></table>
<% If DiscountTypes.Export = "" And DiscountTypes.CurrentAction = "" Then %>
<% End If %>
<%
DiscountTypes_grid.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Drop page object
Set DiscountTypes_grid = Nothing
Set Page = MasterPage
%>
