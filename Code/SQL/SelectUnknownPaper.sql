SELECT
    *
FROM
    Paper
WHERE
    AuthorNames = ''
    or (AuthorNames like '*.*');