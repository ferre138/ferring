<%
Function TotalPurchase()
	dim arrCart,amount
	amount=0
	If Not IsCartEmpty() Then
		arrCart=Session("cart")
		For i=0 To UBound(arrCart)
			amount=amount+arrCart(i,2)*arrCart(i,3)
		Next
	End If
	TotalPurchase=amount
End Function

Function CartExists()
	CartExists=IsArray(Session("cart"))
End Function

Function IsCartEmpty()
	bln=true
	If CartExists() Then
		dim arr
		arr=GetCart()
		bln= (uBound(arr)=-1)
	End If
	IsCartEmpty=bln
End Function

Function GetCart()
	GetCart=Session("cart")
End Function

%>