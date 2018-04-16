<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
	<tbody>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.DiscountCode.FldCaption %></td>
			<td<%= Discountcodes.DiscountCode.CellAttributes %>>
<div<%= Discountcodes.DiscountCode.ViewAttributes %>><%= Discountcodes.DiscountCode.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.Active.FldCaption %></td>
			<td<%= Discountcodes.Active.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.Active.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.Active.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.used.FldCaption %></td>
			<td<%= Discountcodes.used.CellAttributes %>>
<% If ew_ConvertToBool(Discountcodes.used.CurrentValue) Then %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" checked onclick="this.form.reset();" disabled="disabled">
<% Else %>
<input type="checkbox" value="<%= Discountcodes.used.ListViewValue %>" onclick="this.form.reset();" disabled="disabled">
<% End If %></td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.OrderId.FldCaption %></td>
			<td<%= Discountcodes.OrderId.CellAttributes %>>
<div<%= Discountcodes.OrderId.ViewAttributes %>><%= Discountcodes.OrderId.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.Use_date.FldCaption %></td>
			<td<%= Discountcodes.Use_date.CellAttributes %>>
<div<%= Discountcodes.Use_date.ViewAttributes %>><%= Discountcodes.Use_date.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Discountcodes.DiscountTypeId.FldCaption %></td>
			<td<%= Discountcodes.DiscountTypeId.CellAttributes %>>
<div<%= Discountcodes.DiscountTypeId.ViewAttributes %>><%= Discountcodes.DiscountTypeId.ListViewValue %></div>
</td>
		</tr>
	</tbody>
</table>
</div>
</td></tr></table>
<br>
