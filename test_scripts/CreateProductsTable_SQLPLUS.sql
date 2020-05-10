/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

WHENEVER SQLERROR EXIT SQL.CODE

BEGIN

EXECUTE IMMEDIATE 'CREATE TABLE product_range AS (
	SELECT COUNT(p.ProductKey) as Products,
		sc.EnglishProductSubcategoryName,
		c.EnglishProductCategoryName
	FROM dbo.DimProduct p
	LEFT JOIN dbo.DimProductSubCategory sc
	ON p.ProductSubcategoryKey = sc.ProductSubcategoryKey
	LEFT JOIN dbo.DimProductCategory c 
	ON sc.ProductCategoryKey = c.ProductCategoryKey
	GROUP BY 
		sc.EnglishProductSubcategoryName,
		c.EnglishProductCategoryName
	);';
	
END;
/

BEGIN
	
	EXECUTE IMMEDIATE 'COMMENT ON TABLE product_range IS ''Use:Counts the volume of unique products''';
	
END;
/
/* Testing result:
	Tables should be detected, despite being inside an Execute Immediate clause.
	Should be at position 1 on the X-axis, as it refers to a table that is created in another script,
*/