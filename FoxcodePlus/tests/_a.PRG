
text to lcSelect noshow textmerge
	select	Customers.CustomerID
	from Orders Rodrigo
	inner join Customers (nolock) on Orders.CustomerID = customers.CustomerID
	inner join [Order Details] OrderItems (nolock) on Orders.OrderID = OrderItems.OrderID
	inner join Products (nolock) on OrderItems.ProductID = Products.ProductID
	order by CustomerID, OrderID, 
endtext





























text to lcInsert noshow textmerge
	insert into Customers (ContactTitle, CustomerID) values ("Rodrigo")
endtext


text to lcUpdate noshow textmerge
	update Shippers set CompanyName = 'Bruscain' where ShipperID = '0001'
endtext 


text to lcDelect noshow textmerge
	delete from Customers set
endtext    

replace myfield with "myvalue" + 




if sqlexec(lnCnx, lcSelect, "MyCursor") = 1
	messagebox("Sql IntelliSense")	
endif  




if 1 = 0

endif 











text to lcSelect noshow textmerge
	select	Customers.CustomerID, Customers.CompanyName, Customers.City,
			Orders.OrderID, Orders.OrderDate, orderitems.ProductID, 
			Products.ProductName 
	from Orders Rodrigo
	inner join Customers (nolock) on Orders.CustomerID = customers.CustomerID
	inner join [Order Details] OrderItems (nolock) on Orders.OrderID = OrderItems.OrderID
	inner join Products (nolock) on OrderItems.ProductID = Products.ProductID
	inner join  
	order by CustomerID, OrderID, 
endtext
