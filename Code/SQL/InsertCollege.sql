INSERT INTO
    College ([Name])
SELECT
    DISTINCT CollegeName
FROM
    SelectPersonnel
WHERE
    CollegeName not IN (
        SELECT DISTINCT
            [Name]
        FROM
            College
    );