INSERT INTO Personnel ( FirstName, LastName, PositionID, DepartmentID )
SELECT DISTINCT RawAuthor.[First Name], RawAuthor.[Last Name], Position.ID, Department.ID
FROM (RawAuthor INNER JOIN Department ON RawAuthor.Department = Department.Name) INNER JOIN [Position] ON RawAuthor.[Position] = Position.Name;
