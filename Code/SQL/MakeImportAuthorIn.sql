SELECT  DISTINCT ID                                 AS Code 
       ,[Full Name]                                 AS FullName 
       ,FixTitle([Job Title])                       AS JobTitle 
       ,Department                                  AS DepartmentID 
       ,ExtractInDepName([Department Description])  AS DepartmentName 
       ,ExtractInCollName([Department Description]) AS CollegeName 
       ,False as IsStudent
       into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author)  