SELECT  *
FROM ErrorPaperWeight
INNER JOIN RawSCIE
ON (ErrorPaperWeight.Title = RawSCIE.[Journal title])