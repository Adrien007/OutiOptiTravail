--%%%C_TabBalise
SET dateformat ymd
--%%%C_TabBalise
DECLARE @Item1 int,
		@Cost1 int

BEGIN TRANSACTION;

--%%%C_TabBalise
--%%%Result_Attendu:1
--*************************************************************
--*****        Modify Item1                                ******
--*************************************************************
PRINT 'Modify Item1'

SELECT @Item1 = ID 
FROM [RandomSchema].GenericTable 
WHERE ItemName = 'Item1';

--%%%C_TabBalise
--%%%Result_Attendu:1
UPDATE [RandomSchema].GenericTable 
SET LastModified = getdate(),
    Description = REPLACE(Description, 'http://www.example-url.com', 'https://www.new-url.com')
WHERE ID = @Item1 AND Description LIKE '%http://www.example-url.com%';

--%%%C_TabBalise
--%%%Result_Attendu:0
SELECT * 
FROM [RandomSchema].GenericTable 
WHERE ID = @Item1 AND Description LIKE '%http://www.example-url.com%';


--%%%C_TabBalise
--%%%Result_Attendu:1
--*************************************************************
--*****        Modify Cost1                                ******
--*************************************************************
PRINT 'Modify Cost1'

SELECT @Cost1 = ID 
FROM [RandomSchema].GenericTable 
WHERE ItemName = 'Cost1';

--%%%C_TabBalise
--%%%Result_Attendu:1
UPDATE [RandomSchema].GenericTable 
SET LastModified = getdate(),
    URL = REPLACE(URL, 'https://www.example-site.com', 'https://www.new-example-site.com')
WHERE ID = @Cost1 AND URL LIKE '%https://www.example-site.com%';

--%%%C_TabBalise
--%%%Result_Attendu:0
SELECT * 
FROM [RandomSchema].GenericTable 
WHERE ID = @Cost1 AND URL LIKE '%https://www.example-site.com%';

--%%%C_TabBalise
ROLLBACK TRANSACTION;