Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author2 = Abbr.AbbrName
    )
SET
    Paper.Author2 = Abbr.FullName