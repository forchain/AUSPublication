INSERT INTO Author ( FirstName, LastName, PositionID, DepartmentID )
SELECT  GetFirstName(AuthorName) 
       ,GetLastName(AuthorName) 
       ,0 
       ,0
FROM 
(
	SELECT  distinct AuthorName
	FROM SelectUnknownAuthor
)