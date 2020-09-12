SELECT  Paper.*
FROM ImportPaper
INNER JOIN Paper
ON ImportPaper.WoSID = Paper.WoSID