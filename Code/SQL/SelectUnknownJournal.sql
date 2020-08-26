SELECT  DOI, Title, Authors, ISSN, eISSN, Year
FROM Paper
WHERE ( (not IsNull(ISSN)) or (not IsNull(eISSN)) ) 
AND ( ( SCIE = 0 AND ESCI = 0 AND SSCI = 0 AND AHCI = 0 ) ); 