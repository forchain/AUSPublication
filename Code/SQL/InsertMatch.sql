INSERT INTO Match ( ScoreID, PaperID, WoSID, PaperFullName, PaperLastName, PaperFirstName, PaperMiddleName, PaperFirstInitial, PaperMiddleInitial ,AuthorID, AuthorCode, AuthorFullName, AuthorLastName, AuthorFirstName, AuthorMiddleName, AuthorFirstInitial, AuthorMiddleInitial, JobID , JobTitle , JobDisplay , JobOrder , IsStudent , FirstNameCheck , MiddleNameCheck , MiddleInitialCheck, Matched )
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
       ,j.ID as JobID
       ,j.Title as JobTitle
       ,j.Display as JobDisplay
       ,j.Order as JobOrder
       ,j.IsStudent as IsStudent
       ,CBool(FirstNameCheck) 
       ,CBool(MiddleNameCheck) 
       ,CBool(MiddleInitialCheck) 
       ,IsMatched(PaperFirstName,PaperMiddleName,PaperMiddleInitial,AuthorID,AuthorFirstName,AuthorMiddleName,AuthorMiddleInitial,CBool(FirstNameCheck),CBool(MiddleNameCheck),CBool(MiddleInitialCheck))
FROM ImportMatch AS im
Left JOIN Job AS j
ON im.JobID = j.ID