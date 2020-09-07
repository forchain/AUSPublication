INSERT INTO
       Paper (
              WoSID,
              Title,
              [Year],
              [Index],
              Authors,
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
       RawPaper.[Article Title],
       [Year],
       [Index],
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