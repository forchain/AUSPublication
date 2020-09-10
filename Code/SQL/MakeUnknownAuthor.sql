SELECT  DISTINCT SelectWeight.PaperID INTO UnknownAuthor
FROM SelectWeight
WHERE (((IsNull([AuthorID]))<>False)); 