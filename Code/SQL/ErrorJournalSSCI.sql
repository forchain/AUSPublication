SELECT  *
FROM ErrorPaperWeight
INNER JOIN RawSSCI
ON (ErrorPaperWeight.Title = RawSSCI.[Journal title])