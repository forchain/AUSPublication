CREATE TABLE Paper (
    ID COUNTER PRIMARY KEY,
    WoSID VARCHAR,
    DOI VARCHAR,
    Title Memo,
    [Year] int,
    [Index] int,
    Addresses Memo,
    AuthorNames Memo,
    AuthorCount int
);