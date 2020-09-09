SELECT
    *
FROM
    (SELECT  WosID
       ,AuthorNames
FROM Paper
WHERE AuthorNames like '*.*' or AuthorNames = '' ) AS ua,
    Author
WHERE
    ua.AuthorNames like '*' + Author.AbbrName + '*'