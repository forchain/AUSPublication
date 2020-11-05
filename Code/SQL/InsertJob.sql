INSERT INTO Job (Title, Display, [Order],IsStudent)
SELECT  DISTINCT JobTitle
       ,JobTitle
       ,1
       ,InStr(JobTitle,"Stduent") <> 0
FROM ImportAuthor
WHERE JobTitle not IN ( SELECT DISTINCT Title FROM Job )  