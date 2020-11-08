Update
    Match
SET
    AuthorID = ParamAuthorID,
    AuthorCode = ParamAuthorCode,
    AuthorFullName = ParamAuthorFullName,
    AuthorLastName = ParamAuthorLastName,
    AuthorFirstName = ParamAuthorFirstName,
    AuthorMiddleName = ParamAuthorMiddleName,
    AuthorFirstInitial = ParamAuthorFirstInitial,
    AuthorMiddleInitial = ParamAuthorMiddleInitial,
    JobID = ParamJobID,
    JobTitle = ParamJobTitle,
    JobDisplay = ParamJobDisplay,
    JobOrder = ParamJobOrder,
    IsStudent = ParamIsStudent,
    DepartmentID = ParamDepartmentID,
    DepartmentName = ParamDepartmentName
WHERE
    ID = ParamMatchID