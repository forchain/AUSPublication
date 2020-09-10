INSERT INTO
       Job (Title, Display, [Order])
SELECT
       DISTINCT JobTitle,
       JobTitle,
       1
FROM
       SelectPersonnel
WHERE JobTitle not IN ( SELECT DISTINCT Title FROM Job )  