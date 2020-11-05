INSERT INTO Paper ( WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, FullNames )
SELECT  WoSID
       ,DOI
       ,Title
       ,[Year]
       ,[Index]
       ,Addresses
       ,AuthorNames
       ,FullNames
FROM ImportPaper