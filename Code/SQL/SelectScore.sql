SELECT  *
       ,1 AS Abs
       ,CalcScore(w.a.ID,[Index],0,FacultyCount,AllCount) as Weighted
       ,CalcScore(w.a.ID,[Index],1,FacultyCount,AllCount) as AHCI
       ,CalcScore(w.a.ID,[Index],2,FacultyCount,AllCount) as BHCI
       ,CalcScore(w.a.ID,[Index],3,FacultyCount,AllCount) as BSCI
       ,CalcScore(w.a.ID,[Index],4,FacultyCount,AllCount) as ESCI
       ,CalcScore(w.a.ID,[Index],5,FacultyCount,AllCount) as SCIE
       ,CalcScore(w.a.ID,[Index],6,FacultyCount,AllCount) as SSCI
FROM SelectWeight AS w
INNER JOIN SelectFacultyCount AS f
ON w.PaperID = f.PaperID