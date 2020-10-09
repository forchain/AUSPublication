 -- MakeImportPaper
SELECT  [UT (Unique WOS ID)]                                    AS WosID
       ,DOI
       ,[Article Title]                                         AS Title
       ,GetYear([Publication Year],[Early Access Date])         AS [Year]
       ,CByte(LinkPaper.[Index])                                AS [Index]
       ,Addresses
       ,SerializeAuthorNames(Addresses,[Researcher Ids],ORCIDs) AS AuthorNames
       ,[Researcher Ids]                                        AS ResearcherIDs
       ,ORCIDs
       ,CountAuthors(Addresses)                                 AS AuthorCount
       ,ExtractAuthorsText(Addresses)                           AS FullNames into ImportPaper
FROM LinkPaper
WHERE [UT (Unique WOS ID)] not IN ( SELECT WosID FROM Paper ); 