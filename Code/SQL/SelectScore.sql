SELECT
       *,
       1 AS Abs,
       CalcScore(AuthorID, [Index], 0, FacultyCount, AllCount) as Weighted,
       CalcScore(AuthorID, [Index], 1, FacultyCount, AllCount) as AHCI,
       CalcScore(AuthorID, [Index], 2, FacultyCount, AllCount) as BHCI,
       CalcScore(AuthorID, [Index], 3, FacultyCount, AllCount) as BSCI,
       CalcScore(AuthorID, [Index], 4, FacultyCount, AllCount) as ESCI,
       CalcScore(AuthorID, [Index], 5, FacultyCount, AllCount) as SCIE,
       CalcScore(AuthorID, [Index], 6, FacultyCount, AllCount) as SSCI
FROM
       SelectWeight AS w
       INNER JOIN SelectFacultyCount AS f ON w.PaperID = f.PaperID