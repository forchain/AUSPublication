Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author4 = Abbr.AbbrName
    )
SET
    Paper.Author4 = Abbr.FullName