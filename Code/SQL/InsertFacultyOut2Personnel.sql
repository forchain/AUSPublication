INSERT INTO
    Personnel (AuthorName, FullName, AbbrName, Title, DepartmentID, DepartmentName, CollegeName)
SELECT
    DISTINCT GetAuthorName([Name]) As AuthorName,
    [Name],
    GetAbbrName(AuthorName) As AbbrName,
    FixTitle(Title),
    ExtractOutDepID(Department),
    ExtractOutDepName(Department),
    ExtractOutCollName(Department)
FROM
RawPersonnel;