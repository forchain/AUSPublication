SELECT  WosID                   AS [WoS ID] 
       ,DOI 
       ,Title 
       ,Year 
       ,GetIndexName(p.[Index]) AS [Index] 
       ,Addresses 
       ,AuthorNames             AS Authors
FROM Paper AS p
WHERE AuthorNames = '' or (AuthorNames like '*.*');  