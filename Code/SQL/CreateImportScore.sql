CREATE TABLE ImportScore (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    WoSID VARCHAR,

    FullName VARCHAR, 
    LastName VARCHAR, 
    FirstName VARCHAR, 
    MiddleName VARCHAR, 
    FirstInitial VARCHAR, 
    MiddleInitial VARCHAR
);