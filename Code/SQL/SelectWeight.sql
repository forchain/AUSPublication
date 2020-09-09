SELECT
    *
FROM
    (
        [WeIGHT] as w
        INNER JOIN ViewPaper as p ON w.PaperID = p.ID
    )
    LEFT JOIN Author as a ON a.AuthorName = w.AuthorName