SELECT  DISTINCT ID                    AS Code 
       ,[Name]                         AS FullName 
       ,FixTitle(Title)                AS JobTitle 
       ,ExtractOutDepID(Department)    AS DepartmentID 
       ,ExtractOutDepName(Department)  AS DepartmentName 
       ,ExtractOutCollName(Department) AS CollegeName 
       ,False AS IsStudent 
       into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author )  