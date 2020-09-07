Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author5 = Abbr.AbbrName
    )
SET
    Paper.Author5 = Abbr.FullName