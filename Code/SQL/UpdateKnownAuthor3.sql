Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author3 = Abbr.AbbrName
    )
SET
    Paper.Author3 = Abbr.FullName