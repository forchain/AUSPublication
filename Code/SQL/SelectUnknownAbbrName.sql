SELECT  distinct *
FROM 
(
	SELECT  distinct Author1 AS AbbrName
	FROM Paper
	WHERE Mid(Author1, 2, 1) = "." Union 
	SELECT  distinct Author2 AS AbbrName
	FROM Paper
	WHERE Mid(Author2, 2, 1) = "." Union 
	SELECT  distinct Author3 AS AbbrName
	FROM Paper
	WHERE Mid(Author3, 2, 1) = "." Union 
	SELECT  distinct Author4 AS AbbrName
	FROM Paper
	WHERE Mid(Author4, 2, 1) = "." Union 
	SELECT  distinct Author5 AS AbbrName
	FROM Paper
	WHERE Mid(Author5, 2, 1) = "." Union 
	SELECT  distinct Author6 AS AbbrName
	FROM Paper
	WHERE Mid(Author6, 2, 1) = "." Union 
	SELECT  distinct Author7 AS AbbrName
	FROM Paper
	WHERE Mid(Author7, 2, 1) = "." Union 
	SELECT  distinct Author8 AS AbbrName
	FROM Paper
	WHERE Mid(Author1, 2, 1) = "." Union 
	SELECT  distinct Author9 AS AbbrName
	FROM Paper
	WHERE Mid(Author9, 2, 1) = "."  
)
WHERE Not IsNull(AbbrName )  