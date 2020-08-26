SELECT ID, PaperID, AuthorName
FROM [Weight]
WHERE AuthorName not IN ( SELECT distinct FullName FROM Author)  