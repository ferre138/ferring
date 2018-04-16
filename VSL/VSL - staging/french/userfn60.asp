<%

' Global user functions
' Page Loading event
Sub Page_Loading()

'***Response.Write "Page Loading"
End Sub

' Page Unloaded event
Sub Page_Unloaded()

'***Response.Write "Page Unloaded"
End Sub

function isNewCustomer()
dim cn
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open EW_DB_CONNECTION_STRING

	Dim Sec
	Set Sec = New cAdvancedSecurity
nc=false
If  Sec.IsLoggedIn() then

	Set rst = Server.CreateObject("ADODB.Recordset")
	'strSql= " SELECT NewCustomer   FROM Customers WHERE (((Customers.CustomerID)=" & sec.CurrentUserID & "));"
	strSql= "SELECT Count(Orders.OrderId) AS CountOfOrderId FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerId " 
	strSql= strSql & "WHERE (((Customers.CustomerID)=" & sec.CurrentUserID & ") AND ((Orders.payment_status)='Completed'));"

	rst.open strSql,cn
	if(not rst.eof) then 
		nc=(rst("CountOfOrderId")=0)
	end if
	rst.close
end if
set rst =nothing
set cn =nothing
set Sec =nothing
isNewCustomer=nc
end function

%>
