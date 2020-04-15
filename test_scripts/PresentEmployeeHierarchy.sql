/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

-- Get all current employees and their managers

With one_level AS (
	SELECT	e.EmployeeKey,
			e.FirstName,
			e.LastName,
			e.title,
			e.DepartmentName,
			e.ParentEmployeeKey
	FROM DimEmployee e
	WHERE ISNULL(e.endDate, '9999-12-31') > GETDATE()
	)

SELECT	e1.EmployeeKey,
		CONCAT(e1.FirstName, ' ', e1.LastName) as EmployeeName,
		e1.title,
		e1.DepartmentName,
		e2.EmployeeKey as ManagerKey,
		CASE WHEN e2.EmployeeKey IS NOT NULL THEN CONCAT(e2.FirstName, ' ', e2.LAstName) ELSE NULL END as ManagerName,
		e2.title as ManagerTitle,
		e2.DepartmentName as ManagerDepartment,
		e3.EmployeeKey As SecondLevelManager,
		CASE WHEN e3.EmployeeKey IS NOT NULL THEN CONCAT(e3.FirstName, ' ', e3.LAstName) ELSE NULL END as SecondLevelManagerName,
		e3.DepartmentName as SecondLevelManagerDepartment,
		e4.EmployeeKey As ThirdLevelManager,
		CASE WHEN e4.EmployeeKey IS NOT NULL THEN CONCAT(e4.FirstName, ' ', e4.LAstName) ELSE NULL END as ThirdLevelManagerName,
		e4.DepartmentName as ThirdLevelManagerDepartment
INTO sg.present_employee 
FROM one_level e1
LEFT JOIN one_level e2
ON	e1.ParentEmployeeKey = e2.EmployeeKey
LEFT JOIN one_level e3
ON	e2.ParentEmployeeKey = e3.EmployeeKey
LEFT JOIN one_level e4
ON	e3.ParentEmployeeKey = e4.EmployeeKey;

/* Testing result:
	This should be at location 1 on the X-axis, having only dependencies on pre-existing tables
*/