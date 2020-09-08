CREATE TABLE Department (
    [ID] INT,
    [NAME] VARCHAR,
    CollegeID INT,
    FOREIGN KEY (CollegeID) REFERENCES College(ID)
);