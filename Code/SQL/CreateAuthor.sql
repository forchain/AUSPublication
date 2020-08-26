CREATE TABLE Author (
    ID COUNTER PRIMARY KEY,
    FirstName VARCHAR,
    LastName VARCHAR,
    PositionID int,
    DepartmentID int,
    FOREIGN KEY (DepartmentID) REFERENCES Department(ID),
    FOREIGN KEY (PositionID) REFERENCES Position(ID)
);