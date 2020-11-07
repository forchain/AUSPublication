INSERT INTO Author ( Code, FullName, JobID, DepartmentID, LastName, FirstName, FirstInitial, MiddleName, MiddleInitial )
SELECT  DISTINCT Code 
       ,FullName 
       ,Job.ID 
       ,DepartmentID 
       ,GetAuthorLastName(FullName)      AS LastName 
       ,GetAuthorFirstName(FullName)     AS FirstName 
       ,GetAuthorFirstInitial(FullName)  AS FirstInitial 
       ,GetAuthorMiddleName(FullName)    AS MiddleName 
       ,GetAuthorMiddleInitial(FullName) AS MiddleInitial
FROM 
(ImportAuthor
	INNER JOIN Job
	ON ImportAuthor.JobTitle = Job.Title 
)