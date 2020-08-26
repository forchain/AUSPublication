CREATE TABLE [Weight] (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    PaperTitle VARCHAR,
    AuthorName VARCHAR,
    AuthorID int,
    FOREIGN KEY (PaperID) REFERENCES Paper(ID)
);