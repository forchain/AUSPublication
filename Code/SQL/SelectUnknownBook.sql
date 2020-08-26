SELECT  ID, DOI, Title, Authors, [Address], IBSN, [Year]
FROM Paper
WHERE (not IsNull(ISBN)) 
AND ( [BKCI-S] = 0 AND [BKCI-SSH] = 0 );  