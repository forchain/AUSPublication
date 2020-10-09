SELECT  s.ID            AS ScoreID
       ,s.PaperID       AS PaperID 
       ,s.WoSID         AS WosID 
       ,s.FullName      AS PaperFullName 
       ,s.LastName      AS PaperLastName 
       ,s.FirstName     AS PaperFirstName 
       ,s.MiddleName    AS PaperMiddleName 
       ,s.FirstInitial  AS PaperFirstName 
       ,s.MiddleInitial AS PaperMiddleInitial 
       ,a.ID            AS AuthorID 
       ,a.Code          AS AuthorCode 
       ,a.FullName      AS AuthorFullName 
       ,a.LastName      AS AuthorLastName 
       ,a.FirstName     AS AuthorFirstName 
       ,a.MiddleName    AS AuthorMiddleName 
       ,a.FirstInitial  AS AuthorFirstInitial 
       ,a.MiddleInitial AS AuthorMiddleInitial into ImportMatchh
FROM Score AS s
INNER JOIN ImportAuhtor AS a
ON s.LastName = a.LastName AND s.FirstInitial = a.FirstInitial 