/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */

-- Tab=Surveys per Month
SELECT
	d.CalendarYear,
	d.MonthNumberOfYear,
	SUM(survey_vol) as survey_vol
FROM sg.survey_days sd
INNER JOIN DimDate d
ON sd.DateKey = d.DateKey
GROUP BY d.CalendarYear, d.MonthNumberOfYear;

/* Testing Result:
	As the prerequisite table was created and dropped in the same script, this should be flagged as having an invalid requirement.
	*/