Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author9 = Abbr.AbbrName
    )
SET
    Paper.Author9 = Abbr.FullName