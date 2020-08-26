INSERT INTO Paper ( DOI, Title, Authors, [Address], ISSN, eISSN, ISBN, [Year], [Weight], AHCI, SCIE, SSCI, ESCI , [BKCI-S], [BKCI-SSH])
SELECT  RawPaper.[DI (DOI)] 
       ,RawPaper.Title 
       ,RawPaper.[Author Full name] 
       ,RawPaper.[Author Address] 
       ,RawPaper.SN 
       ,RawPaper.EI 
       ,RawPaper.BN 
       ,[Year] 
       ,FormatNumber( 1 / TokenNum([RawPaper].[Author Full name],";"),2 ) AS [Weight] 
       ,IIf(IsNull([RawAHCI].[Journal title]),0,[Weight])                 AS AHCI 
       ,IIf(IsNull([RawSCIE].[Journal title]),0,[Weight])                 AS SCIE 
       ,IIf(IsNull([RawSSCI].[Journal title]),0,[Weight])                 AS SSCI 
       ,IIf(IsNull([RawESCI].[Journal title]),0,[Weight])                 AS ESCI 
       ,IIf(IsNull([RawBKCI-S].[Book title]),0,[Weight])                  AS [BKCI-S] 
       ,IIf(IsNull([RawBKCI-SSH].[Book title]),0,[Weight])                AS [BKCI-SSH]
FROM 
( ( ( ( ( ( RawPaper
	LEFT JOIN RawSCIE
	ON (RawPaper.SN = RawSCIE.ISSN) OR (RawPaper.EI = RawSCIE.eISSN) )
	LEFT JOIN RawSSCI
	ON (RawPaper.SN = RawSSCI.ISSN) OR (RawPaper.EI = RawSSCI.eISSN) )
	LEFT JOIN RawESCI
	ON (RawPaper.SN = RawESCI.ISSN) OR (RawPaper.EI = RawESCI.eISSN) )
	LEFT JOIN RawAHCI
	ON (RawPaper.SN = RawAHCI.ISSN) OR (RawPaper.EI = RawAHCI.eISSN) )
	LEFT JOIN [RawBKCI-S]
	ON (RawPaper.BN = [RawBKCI-S].ISBN) )
	LEFT JOIN [RawBKCI-SSH]
	ON (RawPaper.BN = [RawBKCI-SSH].ISBN) 
) ;