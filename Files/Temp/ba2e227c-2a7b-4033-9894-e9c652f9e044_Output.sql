IF NOT EXISTS(SELECT 1 FROM SYS.OBJECTS WHERE NAME = 'PlaceHolderTable')
CREATE TABLE PlaceHolderTable (
Column0 VARCHAR(MAX)
Column1 VARCHAR(MAX)
)
GO

GO

INSERT INTO PlaceHolderTable
('Column0','Column1')
VALUES

(
'Test Column 1',
'Test Column 2',

),(
'Data',
'New Data',

)