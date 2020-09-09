SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    1 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkAHCI-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    1 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkAHCI-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    1 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkAHCI-2020]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    2 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBHCI-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    2 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBHCI-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    2 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBHCI-2020]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    3 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBSCI-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    3 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBSCI-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    3 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkBSCI-2020]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    4 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkESCI-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    4 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkESCI-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    4 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkESCI-2020]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    5 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSCIE-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    5 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSCIE-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    5 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSCIE-2020]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2018 As [Year],
    6 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSSCI-2018]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2019 As [Year],
    6 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSSCI-2019]
UNION
SELECT
    [UT (Unique WOS ID)] As WoSID,
    DOI,
    [Article Title] As Title,
    2020 As [Year],
    6 As [Index],
    Addresses,
    SerializeAuthorNames(Addresses) As AuthorNames,
    CountAuthors(Addresses) As AuthorCount
FROM
    [LinkSSCI-2020]