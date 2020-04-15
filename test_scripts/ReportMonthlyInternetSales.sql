/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

SELECT	d.CalendarYear,
		d.MonthNumberOfYear,
		d.EnglishMonthName,
		COUNT(*) as OrderVolume,
		SUM(salesAmount) as SalesValue,
		SUM(taxAmt) as TaxValue,
		SUM(freight) as Freight
FROM factInternetSales s
INNER JOIN DimDate d
ON s.OrderDateKey = d.DateKey
GROUP BY d.CalendarYear,
		d.MonthNumberOfYear,
		d.EnglishMonthName
ORDER BY d.CalendarYear, d.MonthNumberOfYear ASC

/* Testing Result:
	This should be at level 1 on the X axis, using only a pre-existing table
	There should be no scripts with dependencies, because this script does not change the db
	*/