 update Score AS s
INNER JOIN 
(
	SELECT  ScoreID
	FROM Match
	WHERE Matched 
	GROUP BY  ScoreID
	HAVING COUNT(ScoreID) = 1 
) AS m
ON (IsNull(s.AuthorID) AND s.ID = m.ScoreID)

SET s.AuthorID = m.AuthorID