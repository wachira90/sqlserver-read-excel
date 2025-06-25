# sqlserver-read-excel
sqlserver read excel

## DOWNLOAD INSTALL EXTENSION

``Microsoft Access Database Engine 2016 Redistributable``

```
https://www.microsoft.com/en-us/download/details.aspx?id=54920
```
![](img/img4.jpg)

## ALLOW EXTENSION

```
USE [master] 
GO 

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1 
GO 

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
GO 

EXEC sp_configure 'show advanced options', 1
RECONFIGURE WITH OVERRIDE
GO

EXEC sp_configure 'ad hoc distributed queries', 1
RECONFIGURE WITH OVERRIDE
GO
```

## COMMAND CONNECT

### OPENROWSET

```
SELECT * 
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0 Xml;Database=C:\sample\book1.xlsx;', Sheet1$);
```

### FOR SPACE SHEET

```
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0 Xml;Database=C:\sample\book1234.xlsx;', 'SELECT * FROM [Sheet kk$]');
```

### OPENDATASOURCE

```
SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=C:\sample\book1.xlsx;Extended Properties=EXCEL 12.0')...[Sheet1$];
```

![](img/img1.jpg)

### EXAMPLE

#### INSERT INTO DestinationTableName (with create new table)

```
USE [testdb]
GO
SELECT * INTO DestinationTableName FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0 Xml;Database=C:\sample\book1234.xlsx;', Sheet1$);
```

#### INSERT INTO DestinationTableName (with exist table)

```
INSERT INTO [amsprod].[dbo].[ams_incident_bk] (IncidentID)
SELECT TOP 10 [Incident ID]  FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0 Xml;Database=C:\sample\book1234.xlsx;', Sheet1$);
```

##### all field 

```
INSERT INTO [amsprod].[dbo].[ams_incident_bk] (*)
SELECT TOP 10 * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0 Xml;Database=C:\sample\book1234.xlsx;', Sheet1$);
```


#### FROM SERVER MAP DRIVE

```
SELECT * FROM OPENROWSET(
  'Microsoft.ACE.OLEDB.12.0',
  'Excel 12.0;Database=\\\\FileServer\\ExcelShare\\HRMSDATA.xlsx;HDR=YES;IMEX=1',
  'SELECT * FROM [EMPMASTER$]'
  )
```

## EXAMPLE EXCEL

![](img/img2.jpg)

## CREATE STORE PROCEDURE GET EXCEL

```
/**
FILE SYNTAX
book_2022_11.xlsx
book_2022_12.xlsx
book_2023_01.xlsx
BY https://github.com/wachira90/sqlserver-read-excel
**/

CREATE PROCEDURE GETEXCEL @SELDATE nvarchar(7)
AS
    DECLARE @GetDate AS VARCHAR(7)
    DECLARE @SQL AS VARCHAR(MAX)
    -- SET @GetDate = (SELECT STUFF(CONCAT(YEAR(GETDATE()) , MONTH(GETDATE() )),5,0,'_'))  //  2022_11
    SET @SQL = 'SELECT * FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', ''Excel 12.0 Xml;Database=C:\sample\book_' + CONVERT(VARCHAR(MAX),@SELDATE) + '.xlsx;'', Sheet1$);'
    EXEC(@SQL)
GO

/** RUN COMMAND **/
EXEC GETEXCEL '2022_11'

```
