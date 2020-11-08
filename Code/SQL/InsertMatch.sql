INSERT INTO Match ( ScoreID, PaperID, WoSID, PaperFullName, PaperLastName, PaperFirstName, PaperMiddleName, PaperFirstInitial, PaperMiddleInitial, AuthorID, AuthorCode, AuthorFullName, AuthorLastName, AuthorFirstName, AuthorMiddleName, AuthorFirstInitial, AuthorMiddleInitial, JobID, JobTitle, JobDisplay, JobOrder, IsStudent, DepartmentID, DepartmentName, FirstNameCheck, MiddleNameCheck, MiddleInitialCheck, Matched )
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
       ,j.ID        AS JobID
       ,j.Title     AS JobTitle
       ,j.Display   AS JobDisplay
       ,j.Order     AS JobOrder
       ,j.IsStudent AS IsStudent
       ,d.ID        AS DepartmentID
       ,d.NAME      AS DepartmentName
       ,CBool(FirstNameCheck)
       ,CBool(MiddleNameCheck)
       ,CBool(MiddleInitialCheck)
       ,IsMatched( PaperFirstName,PaperMiddleName,PaperMiddleInitial,AuthorID,AuthorFirstName,AuthorMiddleName,AuthorMiddleInitial,CBool(FirstNameCheck),CBool(MiddleNameCheck),CBool(MiddleInitialCheck) )
FROM (ImportMatch AS im
LEFT JOIN Job AS j
ON im.JobID = j.ID) LEFT join Department as d on (d.ID = im.DepartmentID)