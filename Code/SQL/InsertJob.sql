INSERT INTO Job (Title, Display, [Order],IsStudent)
SELECT  DISTINCT JobTitle
       ,JobTitle
       ,IIf(IsStudent, 2, 1)
       ,IsStudent
FROM ImportAuthor
WHERE JobTitle not IN ( SELECT DISTINCT Title FROM Job )  