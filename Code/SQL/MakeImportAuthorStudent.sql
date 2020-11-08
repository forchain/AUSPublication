SELECT  DISTINCT IIF(IsNull([ID]),GetTempCode(DepartmentID,JobTitle,FullName),[ID]) AS Code 
       ,[First Name] & " " & [Last name]                                            AS FullName 
       ,FixTitle([Position])                                                        AS JobTitle 
       ,ExtractOutDepID([Department])                                               AS DepartmentID 
       ,ExtractOutDepName([Department])                                             AS DepartmentName 
       ,ExtractOutCollName([Department])                                            AS CollegeName 
       ,True                                                                        AS IsStudent into ImportAuthor
FROM LinkAuthor
WHERE IIF(IsNull(ID), ([First Name] & " " & [Last name]) not IN ( SELECT DISTINCT FullName FROM Author) , ID not IN ( SELECT DISTINCT Code FROM Author) )  