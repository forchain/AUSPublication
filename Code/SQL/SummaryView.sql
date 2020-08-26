SELECT
    DISTINCT Author.ID AS AuthorID,
    First(Author.FirstName) AS FirstName,
    First(Author.LastName) AS LastName,
    First(Author.FullName) AS AuthorName,
    First(Position.Name) AS PositionName,
    First(Position.ID) AS PositionID,
    First(Author.DepartmentID) AS DepartmentID,
    [AbsSum] + [WeightSum] AS Total,
    COUNT([Weight].ID) AS AbsSum,
    [SCIEWeight] + [SSCIWeight] + [AHCIWeight] + [BKCI-SWeight] + [BKCI-SSHWeight] + [ESCIWeight] AS WeightSum,
    SUM(FormatNumber(IIf(IsNull([SCIE]), 0, [SCIE]), 2)) AS SCIEWeight,
    COUNT(IIf([SCIE] > 0, 1, Null)) AS AbsSCIE,
    SUM(FormatNumber(IIf(IsNull([SSCI]), 0, [SSCI]), 2)) AS SSCIWeight,
    SUM(FormatNumber(IIf(IsNull([AHCI]), 0, [AHCI]), 2)) AS [AHCIWeight],
    SUM(FormatNumber(IIf(IsNull([BKCI-S]), 0, [BKCI-S]), 2)) AS [BKCI-SWeight],
    SUM(
        FormatNumber(IIf(IsNull([BKCI-SSH]), 0, [BKCI-SSH]), 2)
    ) AS [BKCI-SSHWeight],
    SUM(FormatNumber(IIf(IsNull([ESCI]), 0, [ESCI]), 2)) AS ESCIWeight,
    COUNT(IIf([SSCI] > 0, 1, Null)) AS AbsSSCI,
    COUNT(IIf([AHCI] > 0, 1, Null)) AS [AbsAHCI],
    COUNT(IIf([BKCI-S] > 0, 1, Null)) AS [AbsBKCI-S],
    COUNT(IIf([BKCI-SSH] > 0, 1, Null)) AS [AbsBKCI-SSH],
    COUNT(IIf([ESCI] > 0, 1, Null)) AS AbsESCI
FROM
    [Position]
    INNER JOIN (
        Author
        INNER JOIN (
            Paper
            INNER JOIN [Weight] ON Paper.ID = [Weight].PaperID
        ) ON Author.FullName = [Weight].AuthorName
    ) ON Position.ID = Author.PositionID
GROUP BY
    Author.ID;