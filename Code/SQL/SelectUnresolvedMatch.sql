SELECT  *
FROM Match 
WHERE Not Matched or ID IN ( SELECT First(ID) FROM Match WHERE Matched GROUP BY ScoreID HAVING COUNT(ScoreID) > 1 )  