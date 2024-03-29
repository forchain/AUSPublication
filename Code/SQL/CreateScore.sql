CREATE TABLE Score (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    WoSID VARCHAR,
    [Index] byte,
    AuthorCount byte,
    [Year] int,

    FullName VARCHAR, 
    LastName VARCHAR, 
    FirstName VARCHAR, 
    MiddleName VARCHAR, 
    FirstInitial VARCHAR, 
    MiddleInitial VARCHAR,

    AuthorID int,
    AuthorCode VARCHAR,
    AuthorName VARCHAR,

    DepartmentID int,
    DepartmentName VARCHAR,

    JobID int,
    JobTitle VARCHAR,
    JobDisplay VARCHAR,
    JobOrder byte,

    AHCI float,
    BHCI float,
    BSCI float,
    ESCI float,
    SCIE float,
    SSCI float,

    Abs float,
    Weighted float
);