

select * from [Weight] left join Author on Left([Weight].AuthorName, 1) = Left(Author.Name, 1)