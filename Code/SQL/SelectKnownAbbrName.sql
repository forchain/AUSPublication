SELECT  DISTINCT First(AuthorName) as FullName, AbbrName
FROM 
(
	SELECT  DISTINCT Paper.Author1     AS AuthorName 
	       ,GetAbbrName(Paper.Author1) AS AbbrName
	FROM Paper
	WHERE Mid(Author1, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author2     AS AuthorName 
	       ,GetAbbrName(Paper.Author2) AS AbbrName
	FROM Paper
	WHERE Mid(Author2, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author3     AS AuthorName 
	       ,GetAbbrName(Paper.Author3) AS AbbrName
	FROM Paper
	WHERE Mid(Author3, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author4     AS AuthorName 
	       ,GetAbbrName(Paper.Author4) AS AbbrName
	FROM Paper
	WHERE Mid(Author4, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author5     AS AuthorName 
	       ,GetAbbrName(Paper.Author5) AS AbbrName
	FROM Paper
	WHERE Mid(Author5, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author6     AS AuthorName 
	       ,GetAbbrName(Paper.Author6) AS AbbrName
	FROM Paper
	WHERE Mid(Author6, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author7     AS AuthorName 
	       ,GetAbbrName(Paper.Author7) AS AbbrName
	FROM Paper
	WHERE Mid(Author7, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author8     AS AuthorName 
	       ,GetAbbrName(Paper.Author8) AS AbbrName
	FROM Paper
	WHERE Mid(Author8, 2, 1) <> '.' Union 
	SELECT  DISTINCT Paper.Author9     AS AuthorName 
	       ,GetAbbrName(Paper.Author9) AS AbbrName
	FROM Paper
	WHERE Mid(Author9, 2, 1) <> '.'  
) 
Group by AbbrName
Having  Count(AuthorName) = 1
