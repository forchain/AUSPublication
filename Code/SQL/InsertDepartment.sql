INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT DepartmentID 
       ,DepartmentName 
       ,[College.ID]
FROM ImportAuthor
INNER JOIN College
ON ImportAuthor.CollegeName = College.[Name]
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  