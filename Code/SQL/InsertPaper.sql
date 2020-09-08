INSERT INTO
       Paper (
              WoSID,
              DOI,
              Title,
              [Year],
              [Index],
              Addresses,
              AuthorNames,
              AuthorCount,
              Author1,
              Author2,
              Author3,
              Author4,
              Author5,
              Author6,
              Author7,
              Author8,
              Author9
       )
SELECT
       RawPaper.[UT (Unique WOS ID)],
       RawPaper.DOI,
       RawPaper.[Article Title],
       [Year],
       [Index],
       RawPaper.Addresses,
       SerializeAuthorNames(RawPaper.Addresses),
       CountAuthors(RawPaper.Addresses),
       SelectAuthor(RawPaper.Addresses, 1, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 2, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 3, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 4, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 5, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 6, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 7, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 8, RawPaper.[Researcher Ids], RawPaper.ORCIDs),
       SelectAuthor(RawPaper.Addresses, 9, RawPaper.[Researcher Ids], RawPaper.ORCIDs)
FROM
       RawPaper;