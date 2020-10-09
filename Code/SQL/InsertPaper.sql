INSERT INTO Paper ( WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, AuthorCount, FullNames )
SELECT  WoSID
       ,DOI
       ,Title
       ,[Year]
       ,[Index]
       ,Addresses
       ,AuthorNames
       ,AuthorCount
       ,FullNames
FROM ImportPaper