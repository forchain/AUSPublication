Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author6 = Abbr.AbbrName
    )
SET
    Paper.Author6 = Abbr.FullName