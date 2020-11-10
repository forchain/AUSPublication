 update Score AS s
INNER JOIN ResolvedMatch as rm
ON   (s.ID = rm.ScoreID )

SET s.AuthorID = rm.AuthorID
, s.AuthorCode = rm.AuthorCode
, s.AuthorName = rm.AuthorFullName

, s.FullName = rm.PaperFullName
, s.LastName = rm.PaperLastName
, s.FirstName = rm.PaperFirstName
, s.FirstInitial = rm.PaperFirstInitial
, s.MiddleName = rm.PaperMiddleName
, s.MiddleInitial = rm.PaperMiddleInitial

, s.JobID = rm.JobID
, s.JobTitle = rm.JobTitle
, s.JobDisplay = rm.JobDisplay
, s.JobOrder = rm.JobOrder
, s.DepartmentID = rm.DepartmentID
, s.DepartmentName = rm.DepartmentName
, s.AHCI = CalcScore(IsStudent, [Index], 1, FacultyCount, AuthorCount )
, s.BHCI = CalcScore(IsStudent, [Index], 2, FacultyCount, AuthorCount )
, s.BSCI = CalcScore(IsStudent, [Index], 3, FacultyCount, AuthorCount )
, s.ESCI = CalcScore(IsStudent, [Index], 4, FacultyCount, AuthorCount )
, s.SCIE = CalcScore(IsStudent, [Index], 5, FacultyCount, AuthorCount )
, s.SSCI = CalcScore(IsStudent, [Index], 6, FacultyCount, AuthorCount )
, s.Weighted = CalcScore(IsStudent, [Index], 0, FacultyCount, AuthorCount )
, s.Abs = 1