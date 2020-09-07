Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author8 = Abbr.AbbrName
    )
SET
    Paper.Author8 = Abbr.FullName