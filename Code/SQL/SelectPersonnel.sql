SELECT  *
FROM 
(
	SELECT  DISTINCT GetAuthorName([Full Name]) AS AuthorName 
	       ,[Full Name] As FullName
	       ,GetAbbrName(AuthorName)             AS AbbrName 
	       ,FixTitle([Job Title]) As JobTitle
	       ,Department  As DepartmentID
	       ,ExtractInDepName([Department Description]) As DepartmentName
	       ,ExtractInCollName([Department Description]) As CollegeName
	FROM RawFacultyIn union
	SELECT  DISTINCT GetAuthorName([Name]) AS AuthorName
	       ,[Name] As FullName
	       ,GetAbbrName(AuthorName)        AS AbbrName
	       ,FixTitle(Title) As JobTitle
	       ,ExtractOutDepID(Department) As DepartmentID
	       ,ExtractOutDepName(Department) As DepartmentName
	       ,ExtractOutCollName(Department) As CollegeName
	FROM  RawFacultyOut
) As Personnel