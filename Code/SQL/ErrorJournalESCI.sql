SELECT  *
FROM ErrorPaperWeight
INNER JOIN RawESCI
ON (ErrorPaperWeight.Title = RawESCI.[Journal title])