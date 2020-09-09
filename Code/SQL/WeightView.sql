SELECT
    DISTINCT Author.ID,
    College.ID,
    (
        0 + [SCIEWeight] + [SSCIWeight] + [A-HCIWeight] + [BKCI-SWeight] + [BKCI-SSHWeight] + [ESCIWeight]
    ) AS Weighted,
    FormatNumber(IIf(IsNull([SCIE]), 0, [SCIE]), 2) AS SCIEWeight,
    FormatNumber(IIf(IsNull([SSCI]), 0, [SSCI]), 2) AS SSCIWeight,
    FormatNumber(IIf(IsNull([A-HCI]), 0, [A-HCI]), 2) AS [A-HCIWeight],
    FormatNumber(IIf(IsNull([BKCI-S]), 0, [BKCI-S]), 2) AS [BKCI-SWeight],
    FormatNumber(IIf(IsNull([BKCI-SSH]), 0, [BKCI-SSH]), 2) AS [BKCI-SSHWeight],
    FormatNumber(IIf(IsNull([ESCI]), 0, [ESCI]), 2) AS ESCIWeight
FROM
    (
        (
            College
            INNER JOIN Department ON College.ID = Department.CollegeID
        )
        INNER JOIN Author ON Department.ID = Author.DepartmentID
    )
    INNER JOIN (
        Paper
        INNER JOIN Weight ON Paper.ID = Weight.PaperID
    ) ON Author.ID = Weight.AuthorID;