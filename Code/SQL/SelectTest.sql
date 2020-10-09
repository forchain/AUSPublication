SELECT  distinct w.LastName AS LastName 
       ,w.FirstInitial      AS FirstInitial 
       ,w.FullName          AS WeightFullName 
       ,w.FirstName         AS WeightFirstName 
       ,w.MiddleName        AS WeightMiddleName 
       ,w.FirstInitial      AS WeightFirstInitial 
       ,w.MiddleInitial     AS WeightMiddleInitial 
       ,a.Code              AS Code 
       ,a.FullName          AS AuthorFullName 
       ,a.FirstName         AS AuthorFirstName 
       ,a.FirstInitial      AS AuthorFirstInitial 
       ,a.MiddleName        AS AuthorMiddleName 
       ,a.MiddleInitial     AS AuthorMiddleInitial 
       ,CalcMatchingScore(WeightFirstName,WeightMiddleName,WeightMiddleInitial,Code,AuthorFirstName,AuthorFirstInitial,AuthorMiddleName,AuthorMiddleInitial) AS Score
       into 
FROM [Weight] AS w
LEFT JOIN Author AS a
ON w.LastName = a.LastName AND w.FirstInitial = a.FirstInitial
ORDER BY w.FullName