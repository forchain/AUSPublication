SELECT  PaperID
       ,COUNT(Not IsStudent) AS FacultyCount
FROM SelectResolvedMatch
GROUP BY  PaperID