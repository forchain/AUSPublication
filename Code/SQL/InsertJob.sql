INSERT INTO
       Job (Title, Display, [Order])
SELECT
       DISTINCT JobTitle,
       JobTitle,
       1
FROM
       SelectPersonnel