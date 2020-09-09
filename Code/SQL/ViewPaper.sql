SELECT
    *
FROM
    SelectKnownPaper
Union
SELECT
    ID, WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, AuthorCount
FROM
    UnknownPaper