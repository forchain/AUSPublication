SELECT  WeightID 
       ,WosID                   AS [WoS ID] 
       ,DOI 
       ,PaperTitle              AS [Paper Title] 
       ,[Year] 
       ,GetIndexName(p.[Index]) AS [Index] 
       ,Addresses 
       ,AuthorName              AS [Author name]
FROM 
(
	SELECT  *
	FROM SelectWeight
	WHERE AuthorID is null  
) AS w