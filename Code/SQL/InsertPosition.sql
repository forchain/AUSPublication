INSERT INTO [Position] (Name, [Order])
SELECT  DISTINCT RawPosition.Position
       ,RawPosition.[Order]
FROM RawPosition
WHERE RawPosition.Position is not null; 