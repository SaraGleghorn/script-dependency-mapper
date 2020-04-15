	/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

-- Get sales
SELECT	d.CalendarYear,
		d.MonthNumberOfYear,
		SUM(s.salesAmount) as sales_value,
		e.ManagerName,
		COUNT(e.employeeKey) as employees,
		ROUND(SUM(s.salesAmount) / COUNT(e.employeeKey),2) as avg_sale_val_per_employee
FROM factResellerSales s
INNER JOIN sg.present_employee e
ON s.EmployeeKey = e.EmployeeKey
INNER JOIN dimDate d
ON s.OrderDateKey = d.DateKey
GROUP BY d.CalendarYear,
		d.MonthNumberOfYear, 
		e.ManagerName

/* Testing result:
	Should be at position 2 on the X-axis, as it refers to a table that is created in another script,
*/