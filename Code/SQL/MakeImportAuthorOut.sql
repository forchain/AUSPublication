SELECT  DISTINCT GetAuthorName([Name]) AS AuthorName 
       ,ID                             AS Code 
       ,[Name]                         AS FullName 
       ,GetAbbrName(AuthorName)        AS AbbrName 
       ,FixTitle(Title)                AS JobTitle 
       ,ExtractOutDepID(Department)    AS DepartmentID 
       ,ExtractOutDepName(Department)  AS DepartmentName 
       ,ExtractOutCollName(Department) AS CollegeName into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author )  