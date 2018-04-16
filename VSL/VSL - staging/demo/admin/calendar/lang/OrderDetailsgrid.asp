<% Session.Timeout = 20 %>
<%

' Define page object
Dim OrderDetails_grid
Set OrderDetails_grid = New cOrderDetails_grid
Set MasterPage = Page
Set Page = OrderDetails_grid

' Page init processing
Call OrderDetails_grid.Page_Init()

' Page main processing
Call OrderDetails_grid.Page_Main()
%>
<% If OrderDetails.Export = "" Then %>
<script type="text/javascript">
<!--
// Create page object
var OrderDetails_grid = new ew_Page("OrderDetails_grid");
// page properties
OrderDetails_grid.PageID = "grid"; // page ID
OrderDetails_grid.FormID = "fOrderDetailsgrid"; // form ID
var EW_PAGE_ID = OrderDetails_grid.PageID; // for backward compatibility
// extend page with ValidateForm function
OrderDetails_grid.ValidateForm = function(fobj) {
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
		elm = fobj.elements["x" + infix + "_Quantity"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(OrderDetails.Quantity.FldErrMsg) %>");
		elm = fobj.elements["x" + infix + "_Price"];
		if (elm && !ew_CheckInteger(elm.value))
			return ew_OnError(this, elm, "<%= ew_JsEncode2(OrderDetails.Price.FldErrMsg) %>");
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
OrderDetails_grid.EmptyRow = function(fobj, infix) {
	if (ew_ValueChanged(fobj, infix, "ProductId")) return false;
	if (ew_ValueChanged(fobj, infix, "Quantity")) return false;
	if (ew_ValueChanged(fobj, infix, "Price")) return false;
	return true;
}
// extend page with Form_CustomValidate function
OrderDetails_grid.Form_CustomValidate =  
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
OrderDetails_grid.SelectAllKey = function(elem) {
	ew_SelectAll(elem);
}
<% If EW_CLIENT_VALIDATE Then %>
OrderDetails_grid.ValidateRequired = true; // uses JavaScript validation
<% Else %>
OrderDetails_grid.ValidateRequired = false; // no JavaScript validation
<% End If %>
//-->
</script>
<% End If %>
<% OrderDetails_grid.ShowPageHeader() %>
<%
If OrderDetails.CurrentAction = "gridadd" Then
	If OrderDetails.CurrentMode <> "copy" Then OrderDetails.CurrentFilter = "0=1"
End If

' Load recordset
Set OrderDetails_grid.Recordset = OrderDetails_grid.LoadRecordset()
If OrderDetails.CurrentAction = "gridadd" Then
	If OrderDetails.CurrentMode = "copy" Then
		OrderDetails_grid.TotalRecs = OrderDetails_grid.Recordset.RecordCount
		OrderDetails_grid.StartRec = 1
		OrderDetails_grid.DisplayRecs = OrderDetails_grid.TotalRecs
	Else
		OrderDetails_grid.StartRec = 1
		OrderDetails_grid.DisplayRecs = OrderDetails.GridAddRowCount
	End If
	OrderDetails_grid.TotalRecs = OrderDetails_grid.DisplayRecs
	OrderDetails_grid.StopRec = OrderDetails_grid.DisplayRecs
Else
	OrderDetails_grid.TotalRecs = OrderDetails_grid.Recordset.RecordCount
	OrderDetails_grid.StartRec = 1
	OrderDetails_grid.DisplayRecs = OrderDetails_grid.TotalRecs ' Display all records
End If
%>
<p class="aspmaker ewTitle" style="white-space: nowrap;"><% If OrderDetails.CurrentMode = "add" Or OrderDetails.CurrentMode = "copy" Then %><%= Language.Phrase("Add") %><% ElseIf OrderDetails.CurrentMode = "edit" Then %><%= Language.Phrase("Edit") %><% End If %>&nbsp;<%= Language.Phrase("TblTypeTABLE") %><%= OrderDetails.TableCaption %></p>
</p>
<% OrderDetails_grid.ShowMessage %>
<br>
<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<% If (OrderDetails.CurrentMode = "add" Or OrderDetails.CurrentMode = "copy" Or OrderDetails.CurrentMode = "edit") And OrderDetails.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridUpperPanel">
<% If OrderDetails.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
<div id="gmp_OrderDetails" class="ewGridMiddlePanel">
<table cellspacing="0" data-rowhighlightclass="ewTableHighlightRow" data-rowselectclass="ewTableSelectRow" data-roweditclass="ewTableEditRow" class="ewTable ewTableSeparate">
<%= OrderDetails.TableCustomInnerHTML %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call OrderDetails_grid.RenderListOptions()

' Render list options (header, left)
OrderDetails_grid.ListOptions.Render "header", "left"
%>
<% If OrderDetails.ProductId.Visible Then ' ProductId %>
	<% If OrderDetails.SortUrl(OrderDetails.ProductId) = "" Then %>
		<td><%= OrderDetails.ProductId.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.ProductId.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.ProductId.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.ProductId.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If OrderDetails.Quantity.Visible Then ' Quantity %>
	<% If OrderDetails.SortUrl(OrderDetails.Quantity) = "" Then %>
		<td><%= OrderDetails.Quantity.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.Quantity.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.Quantity.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.Quantity.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<% If OrderDetails.Price.Visible Then ' Price %>
	<% If OrderDetails.SortUrl(OrderDetails.Price) = "" Then %>
		<td><%= OrderDetails.Price.FldCaption %></td>
	<% Else %>
		<td><div>
			<table cellspacing="0" class="ewTableHeaderBtn"><thead><tr><td><%= OrderDetails.Price.FldCaption %></td><td style="width: 10px;"><% If OrderDetails.Price.Sort = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf OrderDetails.Price.Sort = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></td></tr></thead></table>
		</div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
OrderDetails_grid.ListOptions.Render "header", "right"
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
OrderDetails_grid.StartRec = 1
OrderDetails_grid.StopRec = OrderDetails_grid.TotalRecs ' Show all records

' Restore number of post back records
If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
	ObjForm.Index = 0
	If ObjForm.HasValue("key_count") And (OrderDetails.CurrentAction = "gridadd" Or OrderDetails.CurrentAction = "gridedit" Or OrderDetails.CurrentAction = "F") Then
		OrderDetails_grid.KeyCount = ObjForm.GetValue("key_count")
		OrderDetails_grid.StopRec = OrderDetails_grid.KeyCount
	End If
End If

' Move to first record
OrderDetails_grid.RecCnt = OrderDetails_grid.StartRec - 1
If Not OrderDetails_grid.Recordset.Eof Then
	OrderDetails_grid.Recordset.MoveFirst
	If OrderDetails_grid.StartRec > 1 Then OrderDetails_grid.Recordset.Move OrderDetails_grid.StartRec - 1
ElseIf Not OrderDetails.AllowAddDeleteRow And OrderDetails_grid.StopRec = 0 Then
	OrderDetails_grid.StopRec = OrderDetails.GridAddRowCount
End If

' Initialize Aggregate
OrderDetails.RowType = EW_ROWTYPE_AGGREGATEINIT
Call OrderDetails.ResetAttrs()
Call OrderDetails_grid.RenderRow()
OrderDetails_grid.RowCnt = 0
If OrderDetails.CurrentAction = "gridadd" Then OrderDetails_grid.RowIndex = 0
If OrderDetails.CurrentAction = "gridedit" Then OrderDetails_grid.RowIndex = 0

' Output date rows
Do While CLng(OrderDetails_grid.RecCnt) < CLng(OrderDetails_grid.StopRec)
	OrderDetails_grid.RecCnt = OrderDetails_grid.RecCnt + 1
	If CLng(OrderDetails_grid.RecCnt) >= CLng(OrderDetails_grid.StartRec) Then
		OrderDetails_grid.RowCnt = OrderDetails_grid.RowCnt + 1
		If OrderDetails.CurrentAction = "gridadd" Or OrderDetails.CurrentAction = "gridedit" Or OrderDetails.CurrentAction = "F" Then
			OrderDetails_grid.RowIndex = OrderDetails_grid.RowIndex + 1
			ObjForm.Index = OrderDetails_grid.RowIndex
			If ObjForm.HasValue("k_action") Then
				OrderDetails_grid.RowAction = ObjForm.GetValue("k_action") & ""
			ElseIf OrderDetails.CurrentAction = "gridadd" Then
				OrderDetails_grid.RowAction = "insert"
			Else
				OrderDetails_grid.RowAction = ""
			End If
		End If

	' Set up key count
	OrderDetails_grid.KeyCount = OrderDetails_grid.RowIndex
	Call OrderDetails.ResetAttrs()
	OrderDetails.CssClass = ""
	If OrderDetails.CurrentAction = "gridadd" Then
		If OrderDetails.CurrentMode = "copy" Then
			Call OrderDetails_grid.LoadRowValues(OrderDetails_grid.Recordset) ' Load row values
			OrderDetails_grid.RowOldKey = OrderDetails_grid.SetRecordKey(OrderDetails_grid.Recordset) ' Set old record key
		Else
			Call OrderDetails_grid.LoadDefaultValues() ' Load default values
			OrderDetails_grid.RowOldKey = "" ' Clear old key value
		End If
	Else
		Call OrderDetails_grid.LoadRowValues(OrderDetails_grid.Recordset) ' Load row values
	End If
	OrderDetails.RowType = EW_ROWTYPE_VIEW ' Render view
	If OrderDetails.CurrentAction = "gridadd" Then ' Grid add
		OrderDetails.RowType = EW_ROWTYPE_ADD ' Render add
	End If
	If OrderDetails.CurrentAction = "gridadd" And OrderDetails.EventCancelled Then ' Insert failed
		Call OrderDetails_grid.RestoreCurrentRowFormValues(OrderDetails_grid.RowIndex) ' Restore form values
	End If
	If OrderDetails.CurrentAction = "gridedit" Then ' Grid edit
		If OrderDetails.EventCancelled Then ' Update failed
			Call OrderDetails_grid.RestoreCurrentRowFormValues(OrderDetails_grid.RowIndex) ' Restore form values
		End If
		If OrderDetails_grid.RowAction = "insert" Then
			OrderDetails.RowType = EW_ROWTYPE_ADD ' Render add
		Else
			OrderDetails.RowType = EW_ROWTYPE_EDIT ' Render edit
		End If
	End If
	If OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Edit row
		OrderDetails_grid.EditRowCnt = OrderDetails_grid.EditRowCnt + 1
	End If
	If OrderDetails.CurrentAction = "F" Then ' Confirm row
		Call OrderDetails_grid.RestoreCurrentRowFormValues(OrderDetails_grid.RowIndex) ' Restore form values
	End If
	If OrderDetails.RowType = EW_ROWTYPE_ADD Or OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Add / Edit row
		If OrderDetails.CurrentAction = "edit" Then
			OrderDetails.RowAttrs.AddAttributes Array()
			OrderDetails.CssClass = "ewTableEditRow"
		Else
			OrderDetails.RowAttrs.AddAttributes Array()
		End If
		If Not IsEmpty(OrderDetails_grid.RowIndex) Then
			OrderDetails.RowAttrs.AddAttributes Array(Array("data-rowindex", OrderDetails_grid.RowIndex), Array("id", "r" & OrderDetails_grid.RowIndex & "_OrderDetails"))
		End If
	Else
		OrderDetails.RowAttrs.AddAttributes Array()
	End If

	' Render row
	Call OrderDetails_grid.RenderRow()

	' Render list options
	Call OrderDetails_grid.RenderListOptions()

	' Skip delete row / empty row for confirm page
	If OrderDetails_grid.RowAction <> "delete" And OrderDetails_grid.RowAction <> "insertdelete" And Not (OrderDetails_grid.RowAction = "insert" And OrderDetails.CurrentAction = "F" And OrderDetails_grid.EmptyRow()) Then
%>
	<tr<%= OrderDetails.RowAttributes %>>
<%

' Render list options (body, left)
OrderDetails_grid.ListOptions.Render "body", "left"
%>
	<% If OrderDetails.ProductId.Visible Then ' ProductId %>
		<td<%= OrderDetails.ProductId.CellAttributes %>>
<% If OrderDetails.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<select id="x<%= OrderDetails_grid.RowIndex %>_ProductId" name="x<%= OrderDetails_grid.RowIndex %>_ProductId"<%= OrderDetails.ProductId.EditAttributes %>>
<%
emptywrk = True
If IsArray(OrderDetails.ProductId.EditValue) Then
	arwrk = OrderDetails.ProductId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = OrderDetails.ProductId.CurrentValue&"" Then
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
If emptywrk Then OrderDetails.ProductId.OldValue = ""
%>
</select>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_ProductId" id="o<%= OrderDetails_grid.RowIndex %>_ProductId" value="<%= Server.HTMLEncode(OrderDetails.ProductId.OldValue&"") %>">
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<select id="x<%= OrderDetails_grid.RowIndex %>_ProductId" name="x<%= OrderDetails_grid.RowIndex %>_ProductId"<%= OrderDetails.ProductId.EditAttributes %>>
<%
emptywrk = True
If IsArray(OrderDetails.ProductId.EditValue) Then
	arwrk = OrderDetails.ProductId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = OrderDetails.ProductId.CurrentValue&"" Then
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
If emptywrk Then OrderDetails.ProductId.OldValue = ""
%>
</select>
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= OrderDetails.ProductId.ViewAttributes %>><%= OrderDetails.ProductId.ListViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_ProductId" id="x<%= OrderDetails_grid.RowIndex %>_ProductId" value="<%= Server.HTMLEncode(OrderDetails.ProductId.CurrentValue&"") %>">
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_ProductId" id="o<%= OrderDetails_grid.RowIndex %>_ProductId" value="<%= Server.HTMLEncode(OrderDetails.ProductId.OldValue&"") %>">
<% End If %>
<a name="<%= OrderDetails_grid.PageObjName & "_row_" & OrderDetails_grid.RowCnt %>" id="<%= OrderDetails_grid.PageObjName & "_row_" & OrderDetails_grid.RowCnt %>"></a>
<% If OrderDetails.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_OrderDetailsId" id="o<%= OrderDetails_grid.RowIndex %>_OrderDetailsId" value="<%= Server.HTMLEncode(OrderDetails.OrderDetailsId.OldValue&"") %>">
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_EDIT Then %>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_OrderDetailsId" id="x<%= OrderDetails_grid.RowIndex %>_OrderDetailsId" value="<%= Server.HTMLEncode(OrderDetails.OrderDetailsId.CurrentValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If OrderDetails.Quantity.Visible Then ' Quantity %>
		<td<%= OrderDetails.Quantity.CellAttributes %>>
<% If OrderDetails.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Quantity" id="x<%= OrderDetails_grid.RowIndex %>_Quantity" size="30" value="<%= OrderDetails.Quantity.EditValue %>"<%= OrderDetails.Quantity.EditAttributes %>>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Quantity" id="o<%= OrderDetails_grid.RowIndex %>_Quantity" value="<%= Server.HTMLEncode(OrderDetails.Quantity.OldValue&"") %>">
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Quantity" id="x<%= OrderDetails_grid.RowIndex %>_Quantity" size="30" value="<%= OrderDetails.Quantity.EditValue %>"<%= OrderDetails.Quantity.EditAttributes %>>
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= OrderDetails.Quantity.ViewAttributes %>><%= OrderDetails.Quantity.ListViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_Quantity" id="x<%= OrderDetails_grid.RowIndex %>_Quantity" value="<%= Server.HTMLEncode(OrderDetails.Quantity.CurrentValue&"") %>">
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Quantity" id="o<%= OrderDetails_grid.RowIndex %>_Quantity" value="<%= Server.HTMLEncode(OrderDetails.Quantity.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
	<% If OrderDetails.Price.Visible Then ' Price %>
		<td<%= OrderDetails.Price.CellAttributes %>>
<% If OrderDetails.RowType = EW_ROWTYPE_ADD Then ' Add Record %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Price" id="x<%= OrderDetails_grid.RowIndex %>_Price" size="30" value="<%= OrderDetails.Price.EditValue %>"<%= OrderDetails.Price.EditAttributes %>>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Price" id="o<%= OrderDetails_grid.RowIndex %>_Price" value="<%= Server.HTMLEncode(OrderDetails.Price.OldValue&"") %>">
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_EDIT Then ' Edit Record %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Price" id="x<%= OrderDetails_grid.RowIndex %>_Price" size="30" value="<%= OrderDetails.Price.EditValue %>"<%= OrderDetails.Price.EditAttributes %>>
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_VIEW Then ' View Record %>
<div<%= OrderDetails.Price.ViewAttributes %>><%= OrderDetails.Price.ListViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_Price" id="x<%= OrderDetails_grid.RowIndex %>_Price" value="<%= Server.HTMLEncode(OrderDetails.Price.CurrentValue&"") %>">
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Price" id="o<%= OrderDetails_grid.RowIndex %>_Price" value="<%= Server.HTMLEncode(OrderDetails.Price.OldValue&"") %>">
<% End If %>
</td>
	<% End If %>
<%

' Render list options (body, right)
OrderDetails_grid.ListOptions.Render "body", "right"
%>
	</tr>
<% If OrderDetails.RowType = EW_ROWTYPE_ADD Then %>
<% End If %>
<% If OrderDetails.RowType = EW_ROWTYPE_EDIT Then %>
<% End If %>
<%
	End If
	End If ' End delete row checking
	If OrderDetails.CurrentAction <> "gridadd" Or OrderDetails.CurrentMode = "copy" Then
		If Not OrderDetails_grid.Recordset.Eof Then OrderDetails_grid.Recordset.MoveNext()
	End If
Loop
%>
<%
	If OrderDetails.CurrentMode = "add" Or OrderDetails.CurrentMode = "copy" Or OrderDetails.CurrentMode = "edit" Then
		OrderDetails_grid.RowIndex = "$rowindex$"
		OrderDetails_grid.LoadDefaultValues()

		' Set row properties
		Call OrderDetails.ResetAttrs()
		OrderDetails.RowAttrs.AddAttributes Array()
		If Not IsEmpty(OrderDetails_grid.RowIndex) Then
			OrderDetails.RowAttrs.AddAttributes Array(Array("data-rowindex", OrderDetails_grid.RowIndex), Array("id", "r" & OrderDetails_grid.RowIndex & "_OrderDetails"))
		End If
		OrderDetails.RowType = EW_ROWTYPE_ADD

		' Render row
		Call OrderDetails_grid.RenderRow()

		' Render list options
		Call OrderDetails_grid.RenderListOptions()

		' Add id and class to the template row
		OrderDetails.RowAttrs.UpdateAttribute "id", "r0_OrderDetails"
		OrderDetails.RowAttrs.AddAttribute "class", "ewTemplate", True
%>
	<tr<%= OrderDetails.RowAttributes %>>
<%

' Render list options (body, left)
OrderDetails_grid.ListOptions.Render "body", "left"
%>
	<% If OrderDetails.ProductId.Visible Then ' ProductId %>
		<td<%= OrderDetails.ProductId.CellAttributes %>>
<% If OrderDetails.CurrentAction <> "F" Then %>
<select id="x<%= OrderDetails_grid.RowIndex %>_ProductId" name="x<%= OrderDetails_grid.RowIndex %>_ProductId"<%= OrderDetails.ProductId.EditAttributes %>>
<%
emptywrk = True
If IsArray(OrderDetails.ProductId.EditValue) Then
	arwrk = OrderDetails.ProductId.EditValue
	For rowcntwrk = 0 To UBound(arwrk, 2)
		If arwrk(0, rowcntwrk)&"" = OrderDetails.ProductId.CurrentValue&"" Then
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
If emptywrk Then OrderDetails.ProductId.OldValue = ""
%>
</select>
<% Else %>
<div<%= OrderDetails.ProductId.ViewAttributes %>><%= OrderDetails.ProductId.ViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_ProductId" id="x<%= OrderDetails_grid.RowIndex %>_ProductId" value="<%= Server.HTMLEncode(OrderDetails.ProductId.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_ProductId" id="o<%= OrderDetails_grid.RowIndex %>_ProductId" value="<%= Server.HTMLEncode(OrderDetails.ProductId.OldValue&"") %>">
</td>
	<% End If %>
	<% If OrderDetails.Quantity.Visible Then ' Quantity %>
		<td<%= OrderDetails.Quantity.CellAttributes %>>
<% If OrderDetails.CurrentAction <> "F" Then %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Quantity" id="x<%= OrderDetails_grid.RowIndex %>_Quantity" size="30" value="<%= OrderDetails.Quantity.EditValue %>"<%= OrderDetails.Quantity.EditAttributes %>>
<% Else %>
<div<%= OrderDetails.Quantity.ViewAttributes %>><%= OrderDetails.Quantity.ViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_Quantity" id="x<%= OrderDetails_grid.RowIndex %>_Quantity" value="<%= Server.HTMLEncode(OrderDetails.Quantity.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Quantity" id="o<%= OrderDetails_grid.RowIndex %>_Quantity" value="<%= Server.HTMLEncode(OrderDetails.Quantity.OldValue&"") %>">
</td>
	<% End If %>
	<% If OrderDetails.Price.Visible Then ' Price %>
		<td<%= OrderDetails.Price.CellAttributes %>>
<% If OrderDetails.CurrentAction <> "F" Then %>
<input type="text" name="x<%= OrderDetails_grid.RowIndex %>_Price" id="x<%= OrderDetails_grid.RowIndex %>_Price" size="30" value="<%= OrderDetails.Price.EditValue %>"<%= OrderDetails.Price.EditAttributes %>>
<% Else %>
<div<%= OrderDetails.Price.ViewAttributes %>><%= OrderDetails.Price.ViewValue %></div>
<input type="hidden" name="x<%= OrderDetails_grid.RowIndex %>_Price" id="x<%= OrderDetails_grid.RowIndex %>_Price" value="<%= Server.HTMLEncode(OrderDetails.Price.FormValue&"") %>">
<% End If %>
<input type="hidden" name="o<%= OrderDetails_grid.RowIndex %>_Price" id="o<%= OrderDetails_grid.RowIndex %>_Price" value="<%= Server.HTMLEncode(OrderDetails.Price.OldValue&"") %>">
</td>
	<% End If %>
<%

' Render list options (body, right)
OrderDetails_grid.ListOptions.Render "body", "right"
%>
	</tr>
<%
End If
%>
</tbody>
</table>
<% If OrderDetails.CurrentMode = "add" Or OrderDetails.CurrentMode = "copy" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridinsert">
<input type="hidden" name="key_count" id="key_count" value="<%= OrderDetails_grid.KeyCount %>">
<%= OrderDetails_grid.MultiSelectKey %>
<% End If %>
<% If OrderDetails.CurrentMode = "edit" Then %>
<input type="hidden" name="a_list" id="a_list" value="gridupdate">
<input type="hidden" name="key_count" id="key_count" value="<%= OrderDetails_grid.KeyCount %>">
<%= OrderDetails_grid.MultiSelectKey %>
<% End If %>
<input type="hidden" name="detailpage" id="detailpage" value="OrderDetails_grid">
</div>
<%

' Close recordset and connection
OrderDetails_grid.Recordset.Close
Set OrderDetails_grid.Recordset = Nothing
%>
<% If (OrderDetails.CurrentMode = "add" Or OrderDetails.CurrentMode = "copy" Or OrderDetails.CurrentMode = "edit") And OrderDetails.CurrentAction <> "F" Then ' add/copy/edit mode %>
<div class="ewGridLowerPanel">
<% If OrderDetails.AllowAddDeleteRow Then %>
<% If Security.IsLoggedIn() Then %>
<span class="aspmaker">
<a href="javascript:void(0);" onclick="ew_AddGridRow(this);"><img src='images/addblankrow.gif' alt='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' title='<%= ew_HtmlEncode(Language.Phrase("AddBlankRow")) %>' width='16' height='16' border='0'></a>&nbsp;&nbsp;
</span>
<% End If %>
<% End If %>
</div>
<% End If %>
</td></tr></table>
<% If OrderDetails.Export = "" And OrderDetails.CurrentAction = "" Then %>
<% End If %>
<%
OrderDetails_grid.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<%

' Drop page object
Set OrderDetails_grid = Nothing
Set Page = MasterPage
%>
