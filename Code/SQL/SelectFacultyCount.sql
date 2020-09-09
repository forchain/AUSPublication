SELECT
    PaperID,
    Count(a.ID) as FacultyCount,
    First(AuthorCount) as AllCount
FROM
    SelectWeight
GROUP BY
    PaperID