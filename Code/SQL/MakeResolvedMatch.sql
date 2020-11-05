SELECT  * into ResolvedMatch
FROM SelectResolvedMatch AS rm
INNER JOIN SelectFacultyCount AS fc
ON rm.PaperID = fc.PaperID