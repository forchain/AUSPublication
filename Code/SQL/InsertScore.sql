INSERT INTO Score ( PaperID, WoSID, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial )
SELECT  PaperID
       ,WoSID
       ,FullName
       ,LastName
       ,FirstName
       ,MiddleName
       ,FirstInitial
       ,MiddleInitial
FROM ImportScore