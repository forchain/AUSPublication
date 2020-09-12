INSERT INTO
       Job (Title, Display, [Order])
SELECT
       DISTINCT JobTitle,
       JobTitle,
       1
FROM
      ImportAuthor 
WHERE JobTitle not IN ( SELECT DISTINCT Title FROM Job )  