INSERT INTO Author ( Code, FullName, AuthorName, AbbrName, JobID, DepartmentID )
SELECT  DISTINCT Code
       ,FullName
       ,AuthorName
       ,AbbrName
       ,Job.ID
       ,DepartmentID
FROM 
(ImportAuthor 
	INNER JOIN Job
	ON ImportAuthor.JobTitle = Job.Title 
)