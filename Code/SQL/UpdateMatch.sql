Update
    Match
SET
    PaperLastName = PaperLastName,
    PaperFirstName = PaperFirstName,
    PaperMiddleName = PaperMiddleName,
    PaperFirstInitial = PaperFirstInitial,
    PaperMiddleInitial = PaperMiddleInitial,
    AuthorID = AuthorID,
    AuthorCode = AuthorCode,
    AuthorFullName = AuthorFullName,
    AuthorLastName = AuthorLastName,
    AuthorFirstName = AuthorFirstName,
    AuthorMiddleName = AuthorMiddleName,
    AuthorFirstInitial = AuthorFirstInitial,
    AuthorMiddleInitial = AuthorMiddleInitial,
    FirstNameRequired = FirstNameRequired,
    FirstNameMatched = FirstNameMatched,
    MiddleNameRequired = MiddleNameRequired,
    MiddleNameMatched = MiddleNameMatched,
    MiddleInitialRequired = MiddleInitialRequired,
    MiddleInitialMatched = MiddleInitialMatched,
    Matched = Matched
WHERE
    ID = MatchID