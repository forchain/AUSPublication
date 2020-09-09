SELECT
    PaperID,
    Count(AuthorID) as FacultyCount,
    First(AuthorCount) as AllCount
FROM
    SelectWeight
GROUP BY
    PaperID