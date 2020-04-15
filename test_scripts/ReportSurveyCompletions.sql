/* Purpose ********************************************************************
Testing Script for Sara's Script Dependency Mapper
Using the MS SQL Adventureworks dataset
**************************************************************************** */


SELECT	d.DateKey,
		COUNT(sr.CustomerKey) as survey_vol
INTO sg.survey_days
FROM DimDate d
LEFT JOIN FactSurveyResponse sr
ON d.DateKey = sr.DateKey
GROUP BY d.datekey;

-- Tab=Surveys per Day
SELECT *
FROM sg.survey_days sd;

-- Tab=Surveys per Week
SELECT
	d.CalendarYear,
	d.WeekNumberOfYear,
	SUM(survey_vol) as survey_vol
FROM sg.survey_days sd
INNER JOIN DimDate d
ON sd.DateKey = d.DateKey
GROUP BY d.CalendarYear, d.WeekNumberOfYear;

-- Tab=Surveys per Month
SELECT
	d.CalendarYear,
	d.MonthNumberOfYear,
	SUM(survey_vol) as survey_vol
FROM sg.survey_days sd
INNER JOIN DimDate d
ON sd.DateKey = d.DateKey
GROUP BY d.CalendarYear, d.MonthNumberOfYear;

DROP TABLE sg.survey_days;

/* Testing Result:
	As the only table created here is dropped after use, any scripts that use this as a prerequisite should be flagged
	*/