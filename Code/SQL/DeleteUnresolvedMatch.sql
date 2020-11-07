 Delete
FROM Match
WHERE IsNull(AuthorID) 
AND ScoreID = ParamScoreID and ID <> ParamMatchID 