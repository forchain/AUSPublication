Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author1 = Abbr.AbbrName
    )
SET
    Paper.Author1 = Abbr.FullName