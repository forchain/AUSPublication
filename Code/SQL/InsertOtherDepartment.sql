INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT DepartmentID 
       ,DepartmentName 
       ,1
FROM ImportAuthor
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  