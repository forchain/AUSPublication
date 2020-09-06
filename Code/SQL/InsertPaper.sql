INSERT INTO
       Paper (
              WOSID,
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
       SelectAuthor(RawPaper.Addresses, 1),
       SelectAuthor(RawPaper.Addresses, 2),
       SelectAuthor(RawPaper.Addresses, 3),
       SelectAuthor(RawPaper.Addresses, 4),
       SelectAuthor(RawPaper.Addresses, 5),
       SelectAuthor(RawPaper.Addresses, 6),
       SelectAuthor(RawPaper.Addresses, 7),
       SelectAuthor(RawPaper.Addresses, 8),
       SelectAuthor(RawPaper.Addresses, 9)
FROM
       RawPaper;