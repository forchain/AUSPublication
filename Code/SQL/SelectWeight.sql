SELECT  w.ID               AS WeightID 
       ,PaperID 
       ,w.AuthorName       AS AuthorName 
       ,WoSID 
       ,DOI 
       ,p.Title            AS PaperTitle 
       ,[Year] 
       ,[Index] 
       ,Addresses 
       ,AuthorNames 
       ,AuthorCount 
       ,a.ID               AS AuthorID 
       ,FullName 
       ,AbbrName 
       ,IDOrDef(j.ID)      AS JobID 
       ,IDOrDef(d.ID)      AS DepartmentID 
       ,NameOrDef([NAME])  AS DepartmentName 
       ,IDOrDef(d.CollegeID) As CollegeID
       ,NameOrDef(j.Title) AS JobTitle 
       ,NameOrDef(Display)            AS JobDisplay 
       ,IDOrDef([Order])            AS JobOrder
FROM 
((( [Weight] AS w
	INNER JOIN ViewPaper AS p
	ON w.PaperID = p.ID )
	LEFT JOIN Author AS a
	ON a.AuthorName = w.AuthorName or w.AuthorName = a.AbbrName)
	LEFT JOIN Department AS d
	ON a.DepartmentID = d.ID 
)
LEFT JOIN Job AS j
ON a.JobID = j.ID