CREATE TABLE Match (
    ID COUNTER PRIMARY KEY,

    ScoreID int,
    PaperID int,
    WoSID VARCHAR,
    PaperFullName VARCHAR, 
    PaperLastName VARCHAR, 
    PaperFirstName VARCHAR, 
    PaperMiddleName VARCHAR, 
    PaperFirstInitial VARCHAR, 
    PaperMiddleInitial VARCHAR,

    AuthorID int,
    AuthorCode VARCHAR,
    AuthorFullName VARCHAR,
    AuthorLastName VARCHAR, 
    AuthorFirstName VARCHAR, 
    AuthorMiddleName VARCHAR, 
    AuthorFirstInitial VARCHAR, 
    AuthorMiddleInitial VARCHAR,

    FirstNameRequired boolean,
    FirstNameMatched boolean,
    MiddleNameRequired boolean,
    MiddleNameMatched boolean,
    MiddleInitialRequired boolean,
    MiddleInitialMatched boolean,

    Points float,
    Condition float
);