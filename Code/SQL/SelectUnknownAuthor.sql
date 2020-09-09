SELECT
    distinct 1000000+ First(ID) , FullName , AuthorName , GetAbbrName(AuthorName)  as AbbrName, 0 as JobID, 0 as DepartmentID
FROM
    [Weight]
WHERE
     AuthorName not IN (
        SELECT
            distinct AuthorName
        FROM
            Author
    )
Group by AuthorName