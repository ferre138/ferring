<table cellspacing="0" class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table cellspacing="0" class="ewTable ewTableSeparate">
	<tbody>
		<tr>
			<td class="ewTableHeader"><%= Orders.OrderId.FldCaption %></td>
			<td<%= Orders.OrderId.CellAttributes %>>
<div<%= Orders.OrderId.ViewAttributes %>><%= Orders.OrderId.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.CustomerId.FldCaption %></td>
			<td<%= Orders.CustomerId.CellAttributes %>>
<div<%= Orders.CustomerId.ViewAttributes %>>
<% If Orders.CustomerId.LinkAttributes <> "" Then %>
<a<%= Orders.CustomerId.LinkAttributes %>><%= Orders.CustomerId.ListViewValue %></a>
<% Else %>
<%= Orders.CustomerId.ListViewValue %>
<% End If %>
</div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Amount.FldCaption %></td>
			<td<%= Orders.Amount.CellAttributes %>>
<div<%= Orders.Amount.ViewAttributes %>><%= Orders.Amount.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Ship_FirstName.FldCaption %></td>
			<td<%= Orders.Ship_FirstName.CellAttributes %>>
<div<%= Orders.Ship_FirstName.ViewAttributes %>><%= Orders.Ship_FirstName.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Ship_LastName.FldCaption %></td>
			<td<%= Orders.Ship_LastName.CellAttributes %>>
<div<%= Orders.Ship_LastName.ViewAttributes %>><%= Orders.Ship_LastName.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.payment_status.FldCaption %></td>
			<td<%= Orders.payment_status.CellAttributes %>>
<div<%= Orders.payment_status.ViewAttributes %>><%= Orders.payment_status.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Ordered_Date.FldCaption %></td>
			<td<%= Orders.Ordered_Date.CellAttributes %>>
<div<%= Orders.Ordered_Date.ViewAttributes %>><%= Orders.Ordered_Date.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.payer_email.FldCaption %></td>
			<td<%= Orders.payer_email.CellAttributes %>>
<div<%= Orders.payer_email.ViewAttributes %>><%= Orders.payer_email.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.payment_gross.FldCaption %></td>
			<td<%= Orders.payment_gross.CellAttributes %>>
<div<%= Orders.payment_gross.ViewAttributes %>><%= Orders.payment_gross.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.payment_fee.FldCaption %></td>
			<td<%= Orders.payment_fee.CellAttributes %>>
<div<%= Orders.payment_fee.ViewAttributes %>><%= Orders.payment_fee.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Tax.FldCaption %></td>
			<td<%= Orders.Tax.CellAttributes %>>
<div<%= Orders.Tax.ViewAttributes %>><%= Orders.Tax.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.Shipping.FldCaption %></td>
			<td<%= Orders.Shipping.CellAttributes %>>
<div<%= Orders.Shipping.ViewAttributes %>><%= Orders.Shipping.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.EmailSent.FldCaption %></td>
			<td<%= Orders.EmailSent.CellAttributes %>>
<div<%= Orders.EmailSent.ViewAttributes %>><%= Orders.EmailSent.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.EmailDate.FldCaption %></td>
			<td<%= Orders.EmailDate.CellAttributes %>>
<div<%= Orders.EmailDate.ViewAttributes %>><%= Orders.EmailDate.ListViewValue %></div>
</td>
		</tr>
		<tr>
			<td class="ewTableHeader"><%= Orders.PromoCodeUsed.FldCaption %></td>
			<td<%= Orders.PromoCodeUsed.CellAttributes %>>
<div<%= Orders.PromoCodeUsed.ViewAttributes %>><%= Orders.PromoCodeUsed.ListViewValue %></div>
</td>
		</tr>
	</tbody>
</table>
</div>
</td></tr></table>
<br>
