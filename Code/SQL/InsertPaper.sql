INSERT INTO Paper ( WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, AuthorCount )
SELECT  WoSID
       ,DOI
       ,Title
       ,[Year]
       ,[Index]
       ,Addresses
       ,AuthorNames
       ,AuthorCount
FROM ImportPaper 