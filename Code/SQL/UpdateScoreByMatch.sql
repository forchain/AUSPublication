update
    Score AS s
    INNER JOIN ResolvedMatch AS rm ON (s.ID = rm.ScoreID)
SET
    s.AuthorID = m.AuthorID,
    s.AuthorCode = m.AuthorCode,
    s.AuthorName = m.AuthorFullName,
    s.AHCI = CalcScore(
        IsStudent,
        [Index],
        1,
        FacultyCount,
        AuthorCount
    ),
    s.BHCI = CalcScore(
        IsStudent,
        [Index],
        2,
        FacultyCount,
        AuthorCount
    ),
    s.BSCI = CalcScore(
        IsStudent,
        [Index],
        3,
        FacultyCount,
        AuthorCount
    ),
    s.ESCI = CalcScore(
        IsStudent,
        [Index],
        4,
        FacultyCount,
        AuthorCount
    ),
    s.SCIE = CalcScore(
        IsStudent,
        [Index],
        5,
        FacultyCount,
        AuthorCount
    ),
    s.SSCI = CalcScore(
        IsStudent,
        [Index],
        6,
        FacultyCount,
        AuthorCount
    ),
    s.Weighted = CalcScore(
        IsStudent,
        [Index],
        0,
        FacultyCount,
        AuthorCount
    ),
    s.Abs = 1
    AND m.ID = MatchID