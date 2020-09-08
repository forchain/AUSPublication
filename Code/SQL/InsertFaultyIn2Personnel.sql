INSERT INTO
    Personnel (AuthorName, FullName, AbbrName, Title, DepartmentID, DepartmentName, CollegeName)
SELECT
    DISTINCT GetAuthorName([Full Name]) As AuthorName,
    [Full Name],
    GetAbbrName(AuthorName) As AbbrName,
    FixTitle([Job Title]),
    Department,
    ExtractInDepName([Department Description]),
    ExtractInCollName([Department Description])
FROM
RawPersonnel;