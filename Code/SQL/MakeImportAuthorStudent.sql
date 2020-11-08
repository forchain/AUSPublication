SELECT  DISTINCT IIF(IsNull([ID]),GetTempCode(DepartmentID,JobTitle,FullName),[ID]) AS Code 
       ,[First Name] & " " & [Last name]                                                      AS FullName 
       ,FixTitle([Position])                                                            AS JobTitle 
       ,ExtractOutDepID([Department])                                                AS DepartmentID 
       ,ExtractOutDepName([Department])                                                 AS DepartmentName 
       ,ExtractOutCollName([Department])                                                AS CollegeName 
       ,True                                                                            AS IsStudent into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author)  