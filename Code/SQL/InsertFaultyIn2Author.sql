INSERT INTO
    Author (AuthorName, FullName, AbbrName, Title, DepartmentID)
SELECT
    DISTINCT GetAuthorName([Full Name]) As AuthorName,
    [Full Name],
    GetAbbrName(AuthorName) As AbbrName,
    FixTitle([Job Title]),
    Department
FROM
RawAuthor;