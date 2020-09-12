INSERT INTO
    College ([Name])
SELECT  top 1 "Other"
FROM College
WHERE not exists ( 
SELECT  *
FROM College
WHERE [name]='Others')  