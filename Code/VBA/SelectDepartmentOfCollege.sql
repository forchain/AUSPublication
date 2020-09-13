SELECT Department.ID, Department.Name, Department.CollegeID
FROM Department
WHERE (((Department.CollegeID)=[Forms]![FormDep]![College_Combo]));
