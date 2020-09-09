SELECT *
FROM Paper
WHERE AuthorNames <> '' and (AuthorNames  not like '*.*' );
