<!-- Begin Left Nav -->
<!--table width="100%" border="0" cellspacing="0" cellpadding="2"-->
<table width="180" border="0" cellspacing="0" cellpadding="2" class="flyoutMenu">
	<tr>
		<td>
<% If IsLoggedIn() Then %>
	<!--tr><td><span class="aspmaker"><a href="Customerslist.asp?cmd=resetall">Customers</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="Customerslist.asp?cmd=resetall">Customers</a></td>
		</tr>
	</table>
<% End If %>
<% If IsLoggedIn() Then %>
	<!--tr><td><span class="aspmaker"><a href="Shippinglist.asp?cmd=resetall">Shipping</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="Shippinglist.asp?cmd=resetall">Shipping</a></td>
		</tr>
	</table>
<% End If %>
<% If IsLoggedIn() Then %>
	<!--tr><td><span class="aspmaker"><a href="Productslist.asp?cmd=resetall">Products</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="Productslist.asp?cmd=resetall">Products</a></td>
		</tr>
	</table>
<% End If %>
<% If IsLoggedIn() And (Not IsSysAdmin()) Then %>
	<!--tr><td><span class="aspmaker"><a href="changepwd.asp">Change Password</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="changepwd.asp">Change Password</a></td>
		</tr>
	</table>
<% End If %>
<% If IsLoggedIn() Then %>
	<!--tr><td><span class="aspmaker"><a href="logout.asp">Logout</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="logout.asp">Logout</a></td>
		</tr>
	</table>
<% ElseIf Right(Request.ServerVariables("URL"), Len("login.asp")) <> "login.asp" Then %>
	<!--tr><td><span class="aspmaker"><a href="login.asp">Login</a></span></td></tr-->
	<table width="175" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td class="flyoutLink"><a href="login.asp">Login</a></td>
		</tr>
	</table>
<% End If %>
<!--/table-->
		</td>
	</tr>
</table>
