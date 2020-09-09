INSERT INTO Paper ( WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, AuthorCount )
SELECT  [UT (Unique WOS ID)]
       ,DOI
       ,[Article Title]
       ,[Year]
       ,[Index]
       ,Addresses
       ,SerializeAuthorNames(Addresses,[Researcher Ids],ORCIDs)
       ,CountAuthors(Addresses)
FROM LinkPaper;