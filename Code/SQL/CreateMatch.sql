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

    DepartmentID int,
    DepartmentName VARCHAR,
    JobID int,
    JobTitle VARCHAR,
    JobDisplay VARCHAR,
    JobOrder VARCHAR,
    IsStudent bit,

    FirstNameCheck bit,
    MiddleNameCheck bit,
    MiddleInitialCheck bit,

    Matched bit
);