CREATE TABLE Score (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    WoSID VARCHAR,

    FullName VARCHAR, 
    LastName VARCHAR, 
    FirstName VARCHAR, 
    MiddleName VARCHAR, 
    FirstInitial VARCHAR, 
    MiddleInitial VARCHAR,

    AuthorID int,
    AuthorCode VARCHAR,
    AuthorName VARCHAR,

    AHCI float,
    BHCI float,
    BSCI float,
    ESCI float,
    SCIE float,
    SSCI float,

    Abs float,
    Weighted float
);