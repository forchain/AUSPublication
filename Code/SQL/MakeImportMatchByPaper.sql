SELECT  CLng(s.ID)      AS ScoreID 
       ,CLng(s.PaperID) AS PaperID 
       ,s.WoSID         AS WosID 
       ,s.FullName      AS PaperFullName 
       ,s.LastName      AS PaperLastName 
       ,s.FirstName     AS PaperFirstName 
       ,s.MiddleName    AS PaperMiddleName 
       ,s.FirstInitial  AS PaperFirstInitial 
       ,s.MiddleInitial AS PaperMiddleInitial 
       ,CVar(a.ID)      AS AuthorID 
       ,a.Code          AS AuthorCode 
       ,a.JobID         AS JobID 
       ,a.DepartmentID  AS DepartmentID 
       ,a.FullName      AS AuthorFullName 
       ,a.LastName      AS AuthorLastName 
       ,a.FirstName     AS AuthorFirstName 
       ,a.MiddleName    AS AuthorMiddleName 
       ,a.FirstInitial  AS AuthorFirstInitial 
       ,a.MiddleInitial AS AuthorMiddleInitial into ImportMatch
FROM ImportScore AS s
LEFT JOIN Author AS a
ON s.LastName = a.LastName AND s.FirstInitial = a.FirstInitial 