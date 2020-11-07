SELECT  Code   AS AuthorCode
       ,d.NAME AS DepartmentName
       ,a.ID   AS AuthorID
       ,j.Title   AS JobTitle
       ,j.Display   AS JobDisplay
       ,j.Order   AS JobOrder
       ,*
FROM 
(Author AS a
	INNER JOIN Job AS j
	ON a.JobID = j.ID 
)
INNER JOIN Department AS d
ON a.DepartmentID = d.ID