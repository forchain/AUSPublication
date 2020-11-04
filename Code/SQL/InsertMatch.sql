INSERT INTO Match ( ScoreID, PaperID, WoSID, PaperFullName, PaperLastName, PaperFirstName, PaperMiddleName, PaperFirstInitial, PaperMiddleInitial ,AuthorID, AuthorCode, AuthorFullName, AuthorLastName, AuthorFirstName, AuthorMiddleName, AuthorFirstInitial, AuthorMiddleInitial, FirstNameCheck , MiddleNameCheck , MiddleInitialCheck, Matched )
SELECT  ScoreID 
       ,PaperID 
       ,WosID 
       ,PaperFullName 
       ,PaperLastName 
       ,PaperFirstName 
       ,PaperMiddleName 
       ,PaperFirstInitial 
       ,PaperMiddleInitial 
       ,CVar(AuthorID)
       ,AuthorCode 
       ,AuthorFullName 
       ,AuthorLastName 
       ,AuthorFirstName 
       ,AuthorMiddleName 
       ,AuthorFirstInitial 
       ,AuthorMiddleInitial 
       ,CBool(FirstNameCheck )
       ,CBool(MiddleNameCheck )
       ,CBool(MiddleInitialCheck )
       ,IsMatched(PaperFirstName,PaperMiddleName,PaperMiddleInitial,AuthorID,AuthorFirstName,AuthorMiddleName,AuthorMiddleInitial,CBool(FirstNameCheck),CBool(MiddleNameCheck),CBool(MiddleInitialCheck))
FROM ImportMatch;