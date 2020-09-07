Update
    (
        Paper
        INNER JOIN Abbr ON Paper.Author7 = Abbr.AbbrName
    )
SET
    Paper.Author7 = Abbr.FullName