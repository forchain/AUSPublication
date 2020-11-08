CREATE TABLE ImportScore (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    WoSID VARCHAR,
    [Index] BYTE,
    AuthorCount BYTE,
    [Year] int,

    FullName VARCHAR, 
    LastName VARCHAR, 
    FirstName VARCHAR, 
    MiddleName VARCHAR, 
    FirstInitial VARCHAR, 
    MiddleInitial VARCHAR
);