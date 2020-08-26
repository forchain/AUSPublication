SELECT  SummaryView.DepartmentID
       ,Department.NAME AS DepartmentName
       ,SummaryView.AuthorID
       ,SummaryView.FirstName
       ,SummaryView.LastName
       ,SummaryView.AuthorName
       ,SummaryView.PositionName
       ,SummaryView.AbsSum
       ,SummaryView.WeightSum
       ,SummaryView.SCIEWeight
       ,SummaryView.AbsSCIE
       ,SummaryView.SSCIWeight
       ,SummaryView.[AHCIWeight]
       ,SummaryView.[BKCI-SWeight]
       ,SummaryView.[BKCI-SSHWeight]
       ,SummaryView.ESCIWeight
       ,SummaryView.AbsSSCI
       ,SummaryView.[AbsAHCI]
       ,SummaryView.[AbsBKCI-S]
       ,SummaryView.[AbsBKCI-SSH]
       ,SummaryView.AbsESCI
       ,Position.[Order]
FROM [Position]
INNER JOIN 
(SummaryView
	INNER JOIN Department
	ON SummaryView.DepartmentID = Department.ID
)
ON Position.ID = SummaryView.PositionID
WHERE (((SummaryView.DepartmentID)=[Forms]![DepForm]![Dep_Combo]))
ORDER BY Position.[Order]; 