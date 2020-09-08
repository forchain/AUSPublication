INSERT INTO
    Author (
        FullName,
        AuthorName,
        AbbrName,
        JobID,
        DepartmentID
    )
SELECT
    DISTINCT FullName,
    AuthorName,
    AbbrName,
    Job.ID,
    DepartmentID
FROM
    (
        SelectPersonnel
        Inner Join Job on SelectPersonnel.JobTitle = Job.Title
    )