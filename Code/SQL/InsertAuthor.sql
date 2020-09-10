INSERT INTO Author ( Code, FullName, AuthorName, AbbrName, JobID, DepartmentID )
SELECT  DISTINCT Code
       ,FullName
       ,AuthorName
       ,AbbrName
       ,Job.ID
       ,DepartmentID
FROM 
( SelectPersonnel
	INNER JOIN Job
	ON SelectPersonnel.JobTitle = Job.Title 
)
WHERE Code Not IN ( SELECT distinct Code FROM Author ) 