CREATE TABLE Paper (
    ID COUNTER PRIMARY KEY,
    WoSID VARCHAR,
    DOI VARCHAR,
    Title Memo,
    [Year] int,
    [Index] byte,
    Addresses Memo,
    AuthorNames Memo,
    AuthorCount int
);