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

## FIX. The error you're seeing in SQL Server 2022:

```
Msg 15281, Level 16, State 1, Procedure xp_cmdshell, Line 1
SQL Server blocked access to procedure 'sys.xp_cmdshell' of component 'xp_cmdshell' because this component is turned off as part of the security configuration for this server.
```

means that the `xp_cmdshell` feature is disabled for security reasons. If you **understand the risks** and **need to enable it**, here's how you can do it:

### ✅ Steps to Enable `xp_cmdshell` in SQL Server

1. **Open SQL Server Management Studio (SSMS)** and connect to your server.

2. **Run the following commands**:

```sql
-- Enable advanced options
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;

-- Enable xp_cmdshell
EXEC sp_configure 'xp_cmdshell', 1;
RECONFIGURE;
```

3. After running these commands, `xp_cmdshell` should be enabled and usable.

---

### ⚠️ Important Security Note

`xp_cmdshell` allows execution of command-line operations from SQL Server, which can be a **security risk** if misused. It's recommended to:

- Use it only when absolutely necessary.
- Restrict access to trusted users.
- Disable it when not in use:

```sql
-- Disable xp_cmdshell
EXEC sp_configure 'xp_cmdshell', 0;
RECONFIGURE;
```

