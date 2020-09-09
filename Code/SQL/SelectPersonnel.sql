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
	FROM LinkFacultyIn union
	SELECT  DISTINCT GetAuthorName([Name]) AS AuthorName
	       ,[Name] As FullName
	       ,GetAbbrName(AuthorName)        AS AbbrName
	       ,FixTitle(Title) As JobTitle
	       ,ExtractOutDepID(Department) As DepartmentID
	       ,ExtractOutDepName(Department) As DepartmentName
	       ,ExtractOutCollName(Department) As CollegeName
	FROM  LinkFacultyOut
	union

	SELECT  DISTINCT GetAuthorName([Full Name]) AS AuthorName
	       ,[Full Name] As FullName
	       ,GetAbbrName(AuthorName)             AS AbbrName 
	       , "Senior"
	       ,Department  As DepartmentID
	       ,ExtractInDepName([Department Description]) As DepartmentName
	       ,"Others" As CollegeName
	FROM  LinkSenior
	union
	SELECT  DISTINCT GetAuthorName([Name]) AS AuthorName
	       ,[Name] As FullName
	       ,GetAbbrName(AuthorName)        AS AbbrName
	       , "Staff"
	       ,ExtractOutDepID(Department) As DepartmentID
	       ,ExtractOutDepName(Department) As DepartmentName
	       ,"Others" As CollegeName
	FROM  LinkStaff
) As Personnel