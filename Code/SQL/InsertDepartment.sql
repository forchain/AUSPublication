INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT SelectPersonnel.DepartmentID 
       ,SelectPersonnel.DepartmentName 
       ,[College.ID]
FROM SelectPersonnel
INNER JOIN College
ON SelectPersonnel.CollegeName = College.[Name]
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  