
select * from Paper 
where ID in (select distinct PaperID from SelectUnknownAuthor)