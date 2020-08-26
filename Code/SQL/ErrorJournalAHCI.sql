SELECT  *
FROM ErrorPaperWeight
INNER JOIN RawAHCI
ON (ErrorPaperWeight.Title = RawAHCI.[Journal title])