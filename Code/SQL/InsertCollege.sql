INSERT INTO
    College ([Name])
SELECT
    DISTINCT CollegeName
FROM
    ImportAuthor
WHERE
    CollegeName not IN (
        SELECT DISTINCT
            [Name]
        FROM
            College
    );