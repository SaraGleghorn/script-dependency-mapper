	/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

-- Cleanup
DROP TABLE sg.present_employee e;

/* Testing result:
	Should be last in the chain on the X-axis, as it refers to a table that is created in another script
*/