INSERT INTO Score ( PaperID, WoSID, [Index], AuthorCount, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial )
SELECT  PaperID
       ,WoSID
       ,[Index]
       ,AuthorCount
       ,FullName
       ,LastName
       ,FirstName
       ,MiddleName
       ,FirstInitial
       ,MiddleInitial
FROM ImportScore