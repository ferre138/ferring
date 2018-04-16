<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
	<tbody>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.DiscountType.FldCaption %></td>
			<td<%= DiscountTypes.DiscountType.CellAttributes %>>
<div<%= DiscountTypes.DiscountType.ViewAttributes %>><%= DiscountTypes.DiscountType.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.DiscountTitle.FldCaption %></td>
			<td<%= DiscountTypes.DiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.DiscountTitle.ViewAttributes %>><%= DiscountTypes.DiscountTitle.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.freeShipping.FldCaption %></td>
			<td<%= DiscountTypes.freeShipping.CellAttributes %>>
<% If ew_ConvertToBool(DiscountTypes.freeShipping.CurrentValue) Then %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= DiscountTypes.freeShipping.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.FreePerQty.FldCaption %></td>
			<td<%= DiscountTypes.FreePerQty.CellAttributes %>>
<div<%= DiscountTypes.FreePerQty.ViewAttributes %>><%= DiscountTypes.FreePerQty.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.SpecialPrice.FldCaption %></td>
			<td<%= DiscountTypes.SpecialPrice.CellAttributes %>>
<div<%= DiscountTypes.SpecialPrice.ViewAttributes %>><%= DiscountTypes.SpecialPrice.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.fDiscountTitle.FldCaption %></td>
			<td<%= DiscountTypes.fDiscountTitle.CellAttributes %>>
<div<%= DiscountTypes.fDiscountTitle.ViewAttributes %>><%= DiscountTypes.fDiscountTitle.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.StartDate.FldCaption %></td>
			<td<%= DiscountTypes.StartDate.CellAttributes %>>
<div<%= DiscountTypes.StartDate.ViewAttributes %>><%= DiscountTypes.StartDate.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.EndDate.FldCaption %></td>
			<td<%= DiscountTypes.EndDate.CellAttributes %>>
<div<%= DiscountTypes.EndDate.ViewAttributes %>><%= DiscountTypes.EndDate.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= DiscountTypes.DiscountPerc.FldCaption %></td>
			<td<%= DiscountTypes.DiscountPerc.CellAttributes %>>
<div<%= DiscountTypes.DiscountPerc.ViewAttributes %>><%= DiscountTypes.DiscountPerc.ListViewValue %></div>
</td>
		</tr>
	</tbody>
</table>
</div>
</td></tr></table>
<br>
