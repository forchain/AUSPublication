INSERT INTO
    Author (AuthorName, FullName, AbbrName, Title, DepartmentID)
SELECT
    DISTINCT GetAuthorName([Name]) As AuthorName,
    [Name],
    GetAbbrName(AuthorName) As AbbrName,
    FixTitle(Title),
    ExtractDepID(Department)
FROM
RawAuthor;