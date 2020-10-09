 Delete
FROM Match
WHERE IsNull(AuthorID) 
AND ScoreID IN ( SELECT distinct ScoreID FROM ImportMatch)  