INSERT INTO Match ( ScoreID, PaperID, WoSID, PaperFullName, PaperLastName, PaperFirstName, PaperMiddleName, PaperFirstInitial, PaperMiddleInitial ,AuthorID, AuthorCode, AuthorFullName, AuthorLastName, AuthorFirstName, AuthorMiddleName, AuthorFirstInitial, AuthorMiddleInitial, FirstNameRequired , FirstNameMatched , MiddleNameRequired , MiddleNameMatched , MiddleInitialRequired , MiddleInitialMatched , Matched)
SELECT  ScoreID 
       ,PaperID 
       ,WosID 
       ,PaperFullName 
       ,PaperLastName 
       ,PaperFirstName 
       ,PaperMiddleName 
       ,PaperFirstName 
       ,PaperMiddleInitial 
       ,AuthorID 
       ,AuthorCode 
       ,AuthorFullName 
       ,AuthorLastName 
       ,AuthorFirstName 
       ,AuthorMiddleName 
       ,AuthorFirstInitial 
       ,AuthorMiddleInitial 
       ,true 
       ,true 
       ,true 
       ,true 
       ,true 
       ,true 
       ,IsMatched(PaperFirstName,PaperMiddleName,PaperMiddleInitial,AuthorID,AuthorFirstName,AuthorFirstInitial,AuthorMiddleName,AuthorMiddleInitial, true, true, true, true, true, true) AS Matched
FROM ImportMatch;