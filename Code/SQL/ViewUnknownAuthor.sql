SELECT  WosID                   AS [WoS ID]
       ,DOI
       ,Title 
       ,Year
       ,GetIndexName(p.[Index]) AS [Index]
       ,Addresses
       ,AuthorNames             AS Authors
FROM Paper as p
WHERE ID IN ( SELECT PaperID FROM UnknownAuthor )  