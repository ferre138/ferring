<% Session.Timeout = 20 %>
<%

' Define page object
Dim Discountcodes_grid
Set Discountcodes_grid = New cDiscountcodes_grid
Set MasterPage = Page
Set Page = Discountcodes_grid

' Page init processing
Call Discountcodes_grid.Page_Init()

' Page main processing
Call Discountcodes_grid.Page_Main()
%>
<% If Discountcodes.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var Discountcodes_grid = new ew_Page("Discountcodes_grid");
// page properties
Discountcodes_grid.PageID = "grid"; // page ID
Discountcodes_grid.FormID = "fDiscountcodesgrid"; // form ID
var EW_PAGE_ID = Discountcodes_grid.PageID; // for backward compatibility
// extend page with ValidateForm function
Discountcodes_grid.ValidateForm = function(fobj) {
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
		elm = fobj.elements["x" + infix + "_OrderId"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.OrderId.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Use_date"];
		if (elm && !ew_CheckDate(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(Discountcodes.Use_date.FldErrMsg) %>");
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
Discountcodes_grid.EmptyRow = function(fobj, infix) {
	if (ew_ValueChanged(fobj, infix, "DiscountCode")) return false;
	if (ew_ValueChanged(fobj, infix, "Active")) return false;
	if (ew_ValueChanged(fobj, infix, "used")) return false;
	if (ew_ValueChanged(fobj, infix, "OrderId")) return false;
	if (ew_ValueChanged(fobj, infix, "Use_date")) return false;
	if (ew_ValueChanged(fobj, infix, "DiscountTypeId")) return false;
	return true;
}
// extend page with Form_CustomValidate function
Discountcodes_grid.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
Discountcodes_grid.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
Discountcodes_grid.ValidateRequired = true; // uses JavaScript validation
<% Else %>
Discountcodes_grid.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<link rel="stylesheet" type="text/css" media="all" href="calendar/calendar-win2k-cold-1.css" title="win2k-1">
<script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="calendar/calendar-setup.js"></script>
<% End If %>
<% Discountcodes_grid.ShowPageHeader() %>
<%
If Discountcodes.CurrentAction = "gridadd" Then
	If Discountcodes.CurrentMode <> "copy" Then Discountcodes.CurrentFilter = "0=1"
End If

' Load recordset
Set Discountcodes_grid.Recordset = Discountcodes_grid.LoadRecordset()
If Discountcodes.CurrentAction = "gridadd" Then
	If Discountcodes.CurrentMode = "copy" Then
		Discountcodes_grid.TotalRecs = Discountcodes_grid.Recordset.RecordCount
		Discountcodes_grid.StartRec = 1
		Discountcodes_grid.DisplayRecs = Discountcodes_grid.TotalRecs
	Else
		Discountcodes_grid.StartRec = 1
		Discountcodes_grid.DisplayRecs = Discountcodes.GridAddRowCount
	End If
	Discountcodes_grid.TotalRecs = Discountcodes_grid.DisplayRecs
	Discountcodes_grid.StopRec = Discountcodes_grid.DisplayRecs
Else
	Discountcodes_grid.TotalRecs = Discountcodes_grid.Recordset.RecordCount
	Discountcodes_grid.StartRec = 1
	Discountcodes_grid.DisplayRecs = Discountcodes_grid.TotalRecs ' Display all records
End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><% If Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Then %><%= Language.Phrase("Add") %><% ElseIf Discountcodes.CurrentMode = "edit" Then %><%= Language.Phrase("Edit") %><% End If %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= Discountcodes.TableCaption %></p>
</p>
<% Discountcodes_grid.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If (Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Or Discountcodes.CurrentMode = "edit") And Discountcodes.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridUpperPanel">
<% If Discountcodes.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
<div id="gmp_Discountcodes" class="ewGridMiddlePanel">
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= Discountcodes.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Discountcodes_grid.RenderListOptions()

' Render list options (header, left)
Discountcodes_grid.ListOptions.Render "header", "left"
%>
<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
	<% If Discountcodes.SortUrl(Discountcodes.DiscountCode) = "" Then %>
		<td><%= Discountcodes.DiscountCode.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.DiscountCode.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.DiscountCode.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.DiscountCode.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.Active.Visible Then ' Active %>
	<% If Discountcodes.SortUrl(Discountcodes.Active) = "" Then %>
		<td><%= Discountcodes.Active.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.Active.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.Active.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.Active.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.used.Visible Then ' used %>
	<% If Discountcodes.SortUrl(Discountcodes.used) = "" Then %>
		<td><%= Discountcodes.used.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.used.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.used.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.used.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.OrderId.Visible Then ' OrderId %>
	<% If Discountcodes.SortUrl(Discountcodes.OrderId) = "" Then %>
		<td><%= Discountcodes.OrderId.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.OrderId.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.OrderId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.OrderId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.Use_date.Visible Then ' Use_date %>
	<% If Discountcodes.SortUrl(Discountcodes.Use_date) = "" Then %>
		<td><%= Discountcodes.Use_date.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.Use_date.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.Use_date.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.Use_date.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
	<% If Discountcodes.SortUrl(Discountcodes.DiscountTypeId) = "" Then %>
		<td><%= Discountcodes.DiscountTypeId.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= Discountcodes.DiscountTypeId.FldCaption %></td><td style="width: 10px;"><% If Discountcodes.DiscountTypeId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Discountcodes.DiscountTypeId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Discountcodes_grid.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
Discountcodes_grid.StartRec = 1
Discountcodes_grid.StopRec = Discountcodes_grid.TotalRecs ' Show all records

' Restore number of post back records
If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
	ObjForm.Index = 0
	If ObjForm.HasValue("key_count") And (Discountcodes.CurrentAction = "gridadd" Or Discountcodes.CurrentAction = "gridedit" Or Discountcodes.CurrentAction = "F") Then
		Discountcodes_grid.KeyCount = ObjForm.GetValue("key_count")
		Discountcodes_grid.StopRec = Discountcodes_grid.KeyCount
	End If
End If

' Move to first record
Discountcodes_grid.RecCnt = Discountcodes_grid.StartRec - 1
If Not Discountcodes_grid.Recordset.Eof Then
	Discountcodes_grid.Recordset.MoveFirst
	If Discountcodes_grid.StartRec > 1 Then Discountcodes_grid.Recordset.Move Discountcodes_grid.StartRec - 1
ElseIf Not Discountcodes.AllowAddDeleteRow And Discountcodes_grid.StopRec = 0 Then
	Discountcodes_grid.StopRec = Discountcodes.GridAddRowCount
End If

' Initialize Aggregate
Discountcodes.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Discountcodes.ResetAttrs()
Call Discountcodes_grid.RenderRow()
Discountcodes_grid.RowCnt = 0
If Discountcodes.CurrentAction = "gridadd" Then Discountcodes_grid.RowIndex = 0
If Discountcodes.CurrentAction = "gridedit" Then Discountcodes_grid.RowIndex = 0

' Output date rows
Do While CLng(Discountcodes_grid.RecCnt) < CLng(Discountcodes_grid.StopRec)
	Discountcodes_grid.RecCnt = Discountcodes_grid.RecCnt + 1
	If CLng(Discountcodes_grid.RecCnt) >= CLng(Discountcodes_grid.StartRec) Then
		Discountcodes_grid.RowCnt = Discountcodes_grid.RowCnt + 1
		If Discountcodes.CurrentAction = "gridadd" Or Discountcodes.CurrentAction = "gridedit" Or Discountcodes.CurrentAction = "F" Then
			Discountcodes_grid.RowIndex = Discountcodes_grid.RowIndex + 1
			ObjForm.Index = Discountcodes_grid.RowIndex
			If ObjForm.HasValue("k_action") Then
				Discountcodes_grid.RowAction = ObjForm.GetValue("k_action") & ""
			ElseIf Discountcodes.CurrentAction = "gridadd" Then
				Discountcodes_grid.RowAction = "insert"
			Else
				Discountcodes_grid.RowAction = ""
			End If
		End If

	' Set up key count
	Discountcodes_grid.KeyCount = Discountcodes_grid.RowIndex
	Call Discountcodes.ResetAttrs()
	Discountcodes.CssClass = ""
	If Discountcodes.CurrentAction = "gridadd" Then
		If Discountcodes.CurrentMode = "copy" Then
			Call Discountcodes_grid.LoadRowValues(Discountcodes_grid.Recordset) ' Load row values
			Discountcodes_grid.RowOldKey = Discountcodes_grid.SetRecordKey(Discountcodes_grid.Recordset) ' Set old record key
		Else
			Call Discountcodes_grid.LoadDefaultValues() ' Load default values
			Discountcodes_grid.RowOldKey = "" ' Clear old key value
		End If
	Else
		Call Discountcodes_grid.LoadRowValues(Discountcodes_grid.Recordset) ' Load row values
	End If
	Discountcodes.RowType = EW_ROWTYPE_VIEW ' Render view
	If Discountcodes.CurrentAction = "gridadd" Then ' Grid add
		Discountcodes.RowType = EW_ROWTYPE_ADD ' Render add
	End If
	If Discountcodes.CurrentAction = "gridadd" And Discountcodes.EventCancelled Then ' Insert failed
		Call Discountcodes_grid.RestoreCurrentRowFormValues(Discountcodes_grid.RowIndex) ' Restore form values
	End If
	If Discountcodes.CurrentAction = "gridedit" Then ' Grid edit
		If Discountcodes.EventCancelled Then ' Update failed
			Call Discountcodes_grid.RestoreCurrentRowFormValues(Discountcodes_grid.RowIndex) ' Restore form values
		End If
		If Discountcodes_grid.RowAction = "insert" Then
			Discountcodes.RowType = EW_ROWTYPE_ADD ' Render add
		Else
			Discountcodes.RowType = EW_ROWTYPE_EDIT ' Render edit
		End If
	End If
	If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit row
		Discountcodes_grid.EditRowCnt = Discountcodes_grid.EditRowCnt + 1
	End If
	If Discountcodes.CurrentAction = "F" Then ' Confirm row
		Call Discountcodes_grid.RestoreCurrentRowFormValues(Discountcodes_grid.RowIndex) ' Restore form values
	End If
	If Discountcodes.RowType = EW_ROWTYPE_ADD Or Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Add / Edit row
		If Discountcodes.CurrentAction = "edit" Then
			Discountcodes.RowAttrs.AddAttributes Array()
			Discountcodes.CssClass = "ewTableEditRow"
		Else
			Discountcodes.RowAttrs.AddAttributes Array()
		End If
		If Not IsEmpty(Discountcodes_grid.RowIndex) Then
			Discountcodes.RowAttrs.AddAttributes Array(Array("data-rowindex", Discountcodes_grid.RowIndex), Array("id", "r" & Discountcodes_grid.RowIndex & "_Discountcodes"))
		End If
	Else
		Discountcodes.RowAttrs.AddAttributes Array()
	End If

	' Render row
	Call Discountcodes_grid.RenderRow()

	' Render list options
	Call Discountcodes_grid.RenderListOptions()

	' Skip delete row / empty row for confirm page
	If Discountcodes_grid.RowAction <> "delete" And Discountcodes_grid.RowAction <> "insertdelete" And Not (Discountcodes_grid.RowAction = "insert" And Discountcodes.CurrentAction = "F" And Discountcodes_grid.EmptyRow()) Then
%>
	<tr<%= Discountcodes.RowAttributes %>>
<%

' Render list options (body, left)
Discountcodes_grid.ListOptions.Render "body", "left"
%>
	<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" size="30" maxlength="6" value="<%= Discountcodes.DiscountCode.EditValue %>"<%= Discountcodes.DiscountCode.EditAttributes %>>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" value="<%= Server.HTMLEncode(Discountcodes.DiscountCode.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" size="30" maxlength="6" value="<%= Discountcodes.DiscountCode.EditValue %>"<%= Discountcodes.DiscountCode.EditAttributes %>>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= Discountcodes.DiscountCode.ViewAttributes %>><%= Discountcodes.DiscountCode.ListViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" value="<%= Server.HTMLEncode(Discountcodes.DiscountCode.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" value="<%= Server.HTMLEncode(Discountcodes.DiscountCode.OldValue&"") %>">
<% End If %>
<a name="<%= Discountcodes_grid.PageObjName & "_row_" & Discountcodes_grid.RowCnt %>" id="<%= Discountcodes_grid.PageObjName & "_row_" & Discountcodes_grid.RowCnt %>"></a>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Discountid" id="o<%= Discountcodes_grid.RowIndex %>_Discountid" value="<%= Server.HTMLEncode(Discountcodes.Discountid.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then %>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_Discountid" id="x<%= Discountcodes_grid.RowIndex %>_Discountid" value="<%= Server.HTMLEncode(Discountcodes.Discountid.CurrentValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.Active.Visible Then ' Active %>
		<td<%= Discountcodes.Active.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_Active" id="x<%= Discountcodes_grid.RowIndex %>_Active" value="1"<%= selwrk %><%= Discountcodes.Active.EditAttributes %>>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Active" id="o<%= Discountcodes_grid.RowIndex %>_Active" value="<%= Server.HTMLEncode(Discountcodes.Active.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_Active" id="x<%= Discountcodes_grid.RowIndex %>_Active" value="1"<%= selwrk %><%= Discountcodes.Active.EditAttributes %>>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<% If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_Active" id="x<%= Discountcodes_grid.RowIndex %>_Active" value="<%= Server.HTMLEncode(Discountcodes.Active.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Active" id="o<%= Discountcodes_grid.RowIndex %>_Active" value="<%= Server.HTMLEncode(Discountcodes.Active.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.used.Visible Then ' used %>
		<td<%= Discountcodes.used.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.used.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_used" id="x<%= Discountcodes_grid.RowIndex %>_used" value="1"<%= selwrk %><%= Discountcodes.used.EditAttributes %>>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_used" id="o<%= Discountcodes_grid.RowIndex %>_used" value="<%= Server.HTMLEncode(Discountcodes.used.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.used.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_used" id="x<%= Discountcodes_grid.RowIndex %>_used" value="1"<%= selwrk %><%= Discountcodes.used.EditAttributes %>>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<% If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_used" id="x<%= Discountcodes_grid.RowIndex %>_used" value="<%= Server.HTMLEncode(Discountcodes.used.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_used" id="o<%= Discountcodes_grid.RowIndex %>_used" value="<%= Server.HTMLEncode(Discountcodes.used.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.OrderId.Visible Then ' OrderId %>
		<td<%= Discountcodes.OrderId.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_OrderId" id="x<%= Discountcodes_grid.RowIndex %>_OrderId" size="30" value="<%= Discountcodes.OrderId.EditValue %>"<%= Discountcodes.OrderId.EditAttributes %>>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_OrderId" id="o<%= Discountcodes_grid.RowIndex %>_OrderId" value="<%= Server.HTMLEncode(Discountcodes.OrderId.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_OrderId" id="x<%= Discountcodes_grid.RowIndex %>_OrderId" size="30" value="<%= Discountcodes.OrderId.EditValue %>"<%= Discountcodes.OrderId.EditAttributes %>>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= Discountcodes.OrderId.ViewAttributes %>>
<% If Discountcodes.OrderId.LinkAttributes <> "" Then %>
<a<%= Discountcodes.OrderId.LinkAttributes %>><%= Discountcodes.OrderId.ListViewValue %></a>
<% Else %>
<%= Discountcodes.OrderId.ListViewValue %>
<% End If %>
</div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_OrderId" id="x<%= Discountcodes_grid.RowIndex %>_OrderId" value="<%= Server.HTMLEncode(Discountcodes.OrderId.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_OrderId" id="o<%= Discountcodes_grid.RowIndex %>_OrderId" value="<%= Server.HTMLEncode(Discountcodes.OrderId.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.Use_date.Visible Then ' Use_date %>
		<td<%= Discountcodes.Use_date.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_Use_date" id="x<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Discountcodes.Use_date.EditValue %>"<%= Discountcodes.Use_date.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" name="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x<%= Discountcodes_grid.RowIndex %>_Use_date", // input field id
	ifFormat: "%Y/%m/%d", // date format
	button: "cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" // button id
});
</script>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Use_date" id="o<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Server.HTMLEncode(Discountcodes.Use_date.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_Use_date" id="x<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Discountcodes.Use_date.EditValue %>"<%= Discountcodes.Use_date.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" name="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x<%= Discountcodes_grid.RowIndex %>_Use_date", // input field id
	ifFormat: "%Y/%m/%d", // date format
	button: "cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" // button id
});
</script>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= Discountcodes.Use_date.ViewAttributes %>><%= Discountcodes.Use_date.ListViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_Use_date" id="x<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Server.HTMLEncode(Discountcodes.Use_date.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Use_date" id="o<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Server.HTMLEncode(Discountcodes.Use_date.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<% If Discountcodes.DiscountTypeId.SessionValue <> "" Then %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
<input type="hidden" id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= ew_HtmlEncode(Discountcodes.DiscountTypeId.CurrentValue) %>">
<% Else %>
<select id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId"<%= Discountcodes.DiscountTypeId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Discountcodes.DiscountTypeId.EditValue) Then
	arwrk = Discountcodes.DiscountTypeId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Discountcodes.DiscountTypeId.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
</option>
<%
	Next
End If
If emptywrk Then Discountcodes.DiscountTypeId.OldValue = ""
%>
</select>
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" id="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(Discountcodes.DiscountTypeId.OldValue&"") %>">
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<% If Discountcodes.DiscountTypeId.SessionValue <> "" Then %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
<input type="hidden" id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= ew_HtmlEncode(Discountcodes.DiscountTypeId.CurrentValue) %>">
<% Else %>
<select id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId"<%= Discountcodes.DiscountTypeId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Discountcodes.DiscountTypeId.EditValue) Then
	arwrk = Discountcodes.DiscountTypeId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Discountcodes.DiscountTypeId.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
</option>
<%
	Next
End If
If emptywrk Then Discountcodes.DiscountTypeId.OldValue = ""
%>
</select>
<% End If %>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(Discountcodes.DiscountTypeId.CurrentValue&"") %>">
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" id="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(Discountcodes.DiscountTypeId.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
<%

' Render list options (body, right)
Discountcodes_grid.ListOptions.Render "body", "right"
%>
	</tr>
<% If Discountcodes.RowType = EW_ROWTYPE_ADD Then %>
<% End If %>
<% If Discountcodes.RowType = EW_ROWTYPE_EDIT Then %>
<% End If %>
<%
	End If
	End If ' End delete row checking
	If Discountcodes.CurrentAction <> "gridadd" Or Discountcodes.CurrentMode = "copy" Then
		If Not Discountcodes_grid.Recordset.Eof Then Discountcodes_grid.Recordset.MoveNext()
	End If
Loop
%>
<%
	If Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Or Discountcodes.CurrentMode = "edit" Then
		Discountcodes_grid.RowIndex = "$rowindex$"
		Discountcodes_grid.LoadDefaultValues()

		' Set row properties
		Call Discountcodes.ResetAttrs()
		Discountcodes.RowAttrs.AddAttributes Array()
		If Not IsEmpty(Discountcodes_grid.RowIndex) Then
			Discountcodes.RowAttrs.AddAttributes Array(Array("data-rowindex", Discountcodes_grid.RowIndex), Array("id", "r" & Discountcodes_grid.RowIndex & "_Discountcodes"))
		End If
		Discountcodes.RowType = EW_ROWTYPE_ADD

		' Render row
		Call Discountcodes_grid.RenderRow()

		' Render list options
		Call Discountcodes_grid.RenderListOptions()

		' Add id and class to the template row
		Discountcodes.RowAttrs.UpdateAttribute "id", "r0_Discountcodes"
		Discountcodes.RowAttrs.AddAttribute "class", "ewTemplate", True
%>
	<tr<%= Discountcodes.RowAttributes %>>
<%

' Render list options (body, left)
Discountcodes_grid.ListOptions.Render "body", "left"
%>
	<% If Discountcodes.DiscountCode.Visible Then ' DiscountCode %>
		<td<%= Discountcodes.DiscountCode.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" size="30" maxlength="6" value="<%= Discountcodes.DiscountCode.EditValue %>"<%= Discountcodes.DiscountCode.EditAttributes %>>
<% Else %>
<div<%= Discountcodes.DiscountCode.ViewAttributes %>><%= Discountcodes.DiscountCode.ViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="x<%= Discountcodes_grid.RowIndex %>_DiscountCode" value="<%= Server.HTMLEncode(Discountcodes.DiscountCode.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" id="o<%= Discountcodes_grid.RowIndex %>_DiscountCode" value="<%= Server.HTMLEncode(Discountcodes.DiscountCode.OldValue&"") %>">
</td>
	<% End If %>
	<% If Discountcodes.Active.Visible Then ' Active %>
		<td<%= Discountcodes.Active.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.Active.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_Active" id="x<%= Discountcodes_grid.RowIndex %>_Active" value="1"<%= selwrk %><%= Discountcodes.Active.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.Active.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.Active.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_Active" id="x<%= Discountcodes_grid.RowIndex %>_Active" value="<%= Server.HTMLEncode(Discountcodes.Active.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Active" id="o<%= Discountcodes_grid.RowIndex %>_Active" value="<%= Server.HTMLEncode(Discountcodes.Active.OldValue&"") %>">
</td>
	<% End If %>
	<% If Discountcodes.used.Visible Then ' used %>
		<td<%= Discountcodes.used.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<% selwrk = ew_IIf(ew_ConvertToBool(Discountcodes.used.CurrentValue), " checked=""checked""", "") %>
<input type="checkbox" name="x<%= Discountcodes_grid.RowIndex %>_used" id="x<%= Discountcodes_grid.RowIndex %>_used" value="1"<%= selwrk %><%= Discountcodes.used.EditAttributes %>>
<% Else %>
<% If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.used.ViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.used.ViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_used" id="x<%= Discountcodes_grid.RowIndex %>_used" value="<%= Server.HTMLEncode(Discountcodes.used.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_used" id="o<%= Discountcodes_grid.RowIndex %>_used" value="<%= Server.HTMLEncode(Discountcodes.used.OldValue&"") %>">
</td>
	<% End If %>
	<% If Discountcodes.OrderId.Visible Then ' OrderId %>
		<td<%= Discountcodes.OrderId.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_OrderId" id="x<%= Discountcodes_grid.RowIndex %>_OrderId" size="30" value="<%= Discountcodes.OrderId.EditValue %>"<%= Discountcodes.OrderId.EditAttributes %>>
<% Else %>
<div<%= Discountcodes.OrderId.ViewAttributes %>>
<% If Discountcodes.OrderId.LinkAttributes <> "" Then %>
<a<%= Discountcodes.OrderId.LinkAttributes %>><%= Discountcodes.OrderId.ViewValue %></a>
<% Else %>
<%= Discountcodes.OrderId.ViewValue %>
<% End If %>
</div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_OrderId" id="x<%= Discountcodes_grid.RowIndex %>_OrderId" value="<%= Server.HTMLEncode(Discountcodes.OrderId.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_OrderId" id="o<%= Discountcodes_grid.RowIndex %>_OrderId" value="<%= Server.HTMLEncode(Discountcodes.OrderId.OldValue&"") %>">
</td>
	<% End If %>
	<% If Discountcodes.Use_date.Visible Then ' Use_date %>
		<td<%= Discountcodes.Use_date.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<input type="text" name="x<%= Discountcodes_grid.RowIndex %>_Use_date" id="x<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Discountcodes.Use_date.EditValue %>"<%= Discountcodes.Use_date.EditAttributes %>>
&nbsp;<img src="images/calendar.png" id="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" name="cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" alt="<%= Language.Phrase("PickDate") %>" title="<%= Language.Phrase("PickDate") %>" style="cursor:pointer;cursor:hand;">
<script type="text/javascript">
Calendar.setup({
	inputField: "x<%= Discountcodes_grid.RowIndex %>_Use_date", // input field id
	ifFormat: "%Y/%m/%d", // date format
	button: "cal_x<%= Discountcodes_grid.RowIndex %>_Use_date" // button id
});
</script>
<% Else %>
<div<%= Discountcodes.Use_date.ViewAttributes %>><%= Discountcodes.Use_date.ViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_Use_date" id="x<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Server.HTMLEncode(Discountcodes.Use_date.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_Use_date" id="o<%= Discountcodes_grid.RowIndex %>_Use_date" value="<%= Server.HTMLEncode(Discountcodes.Use_date.OldValue&"") %>">
</td>
	<% End If %>
	<% If Discountcodes.DiscountTypeId.Visible Then ' DiscountTypeId %>
		<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
<% If Discountcodes.CurrentAction <> "F" Then %>
<% If Discountcodes.DiscountTypeId.SessionValue <> "" Then %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
<input type="hidden" id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= ew_HtmlEncode(Discountcodes.DiscountTypeId.CurrentValue) %>">
<% Else %>
<select id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId"<%= Discountcodes.DiscountTypeId.EditAttributes %>>
<%
emptywrk = True
If IsArray(Discountcodes.DiscountTypeId.EditValue) Then
	arwrk = Discountcodes.DiscountTypeId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = Discountcodes.DiscountTypeId.CurrentValue&"" Then
			selwrk = " selected=""selected"""
			emptywrk = False
		Else
			selwrk = ""
		End If
%>
<option value="<%= Server.HtmlEncode(arwrk(0, rowcntwrk)&"") %>"<%= selwrk %>>
<%= arwrk(1, rowcntwrk) %>
</option>
<%
	Next
End If
If emptywrk Then Discountcodes.DiscountTypeId.OldValue = ""
%>
</select>
<% End If %>
<% Else %>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ViewValue %></div>
<input type="hidden" name="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" id="x<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(Discountcodes.DiscountTypeId.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" id="o<%= Discountcodes_grid.RowIndex %>_DiscountTypeId" value="<%= Server.HTMLEncode(Discountcodes.DiscountTypeId.OldValue&"") %>">
</td>
	<% End If %>
<%

' Render list options (body, right)
Discountcodes_grid.ListOptions.Render "body", "right"
%>
	</tr>
<%
End If
%>
</tbody>
</table>
<% If Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridinsert">
<input type="hidden" name="key_count" id="key_count" value="<%= Discountcodes_grid.KeyCount %>">
<%= Discountcodes_grid.MultiSelectKey %>
<% End If %>
<% If Discountcodes.CurrentMode = "edit" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridupdate">
<input type="hidden" name="key_count" id="key_count" value="<%= Discountcodes_grid.KeyCount %>">
<%= Discountcodes_grid.MultiSelectKey %>
<% End If %>
<input type="hidden" name="detailpage" id="detailpage" value="Discountcodes_grid">
</div>
<%

' Close recordset and connection
Discountcodes_grid.Recordset.Close
Set Discountcodes_grid.Recordset = Nothing
%>
<% If (Discountcodes.CurrentMode = "add" Or Discountcodes.CurrentMode = "copy" Or Discountcodes.CurrentMode = "edit") And Discountcodes.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridLowerPanel">
<% If Discountcodes.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
</td></tr></table>
<% If Discountcodes.Export = "" And Discountcodes.CurrentAction = "" Then %>
<% End If %>
<%
Discountcodes_grid.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Drop page object
Set Discountcodes_grid = Nothing
Set Page = MasterPage
%>
