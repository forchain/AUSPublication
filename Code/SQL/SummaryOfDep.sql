SELECT  DISTINCT  
       (GetFirstName(s.AuthorName)) AS FirstName 
       ,(GetLastName(s.AuthorName))  AS LastName 
       ,(s.AuthorName)               AS AuthorName 
       ,First(s.JobTitle)                 AS JobTitle 
       ,First(s.JobID)                    AS JobID 
       ,First(s.JobOrder)                 AS JobOrder 
       ,First(s.DepartmentID)             AS DepartmentID 
       ,First(s.DepartmentName)           AS DepartmentName 
       ,[AbsSum] + [WeightSum]            AS Total 
       ,COUNT(Abs)                        AS AbsSum 
       ,SUM(Weighted)                     AS WeightSum 
       ,SUM(SCIE)                         AS SCIEWeight 
       ,COUNT(IIf([SCIE] > 0,1,Null))     AS AbsSCIE 
       ,SUM(SSCI)                         AS SSCIWeight 
       ,SUM(AHCI)                         AS [AHCIWeight] 
       ,SUM(BSCI)                         AS [BSCIWeight] 
       ,SUM(BHCI)                         AS [BHCIWeight] 
       ,SUM(ESCI)                         AS ESCIWeight 
       ,COUNT(IIf([SSCI] > 0,1,Null))     AS AbsSSCI 
       ,COUNT(IIf([AHCI] > 0,1,Null))     AS [AbsAHCI] 
       ,COUNT(IIf([BSCI] > 0,1,Null))     AS [AbsBSCI] 
       ,COUNT(IIf([BHCI] > 0,1,Null))     AS [AbsBHCI] 
       ,COUNT(IIf([ESCI] > 0,1,Null))     AS AbsESCI
FROM SelectScore AS s
WHERE (DepartmentID) = [Forms]![DepForm]![Dep_Combo] 
AND [Year] = [Forms]![DepForm]![Year_Combo] 
GROUP BY  AuthorName;