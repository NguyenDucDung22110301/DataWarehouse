/****** Object:  Database WideWorldImportersOrderFullfillmentDW    Script Date: 5/22/2025 1:34:35 PM ******/
/*
Kimball Group, The Microsoft Data Warehouse Toolkit
Generate a database from the datamodel worksheet, version: 4

You can use this Excel workbook as a data modeling tool during the logical design phase of your project.
As discussed in the book, it is in some ways preferable to a real data modeling tool during the inital design.
We expect you to move away from this spreadsheet and into a real modeling tool during the physical design phase.
The authors provide this macro so that the spreadsheet isn't a dead-end. You can 'import' into your
data modeling tool by generating a database using this script, then reverse-engineering that database into
your tool.

Uncomment the next lines if you want to drop and create the database
*/
CREATE DATABASE WideWorldImportersOrderFullfillmentDW
GO
ALTER DATABASE WideWorldImportersOrderFullfillmentDW
SET RECOVERY SIMPLE
GO
USE WideWorldImportersOrderFullfillmentDW
;
IF EXISTS (SELECT Name from sys.extended_properties where Name = 'Description')
    EXEC sys.sp_dropextendedproperty @name = 'Description'
EXEC sys.sp_addextendedproperty @name = 'Description', @value = 'Default description - you should change this.'
;





/* Drop table dbo.DimAudit */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimAudit') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimAudit 
;

/* Create table dbo.DimAudit */
CREATE TABLE dbo.DimAudit (
   [AuditKey]  int IDENTITY  NOT NULL
,  [ParentAuditKey]  int   NOT NULL
,  [TableName]  nvarchar(50)  DEFAULT 'Unknown' NOT NULL
,  [PkgName]  nvarchar(50)  DEFAULT 'Unknown' NOT NULL
,  [PkgGUID]  uniqueidentifier   NULL
,  [PkgVersionGUID]  uniqueidentifier   NULL
,  [PkgVersionMajor]  smallint   NULL
,  [PkgVersionMinor]  smallint   NULL
,  [ExecStartDT]  datetime  DEFAULT getdate() NOT NULL
,  [ExecStopDT]  datetime   NULL
,  [ExecutionInstanceGUID]  uniqueidentifier   NULL
,  [ExtractRowCnt]  bigint   NULL
,  [InsertRowCnt]  bigint   NULL
,  [UpdateRowCnt]  bigint   NULL
,  [ErrorRowCnt]  bigint   NULL
,  [TableInitialRowCnt]  bigint   NULL
,  [TableFinalRowCnt]  bigint   NULL
,  [TableMaxSurrogateKey]  bigint   NULL
,  [SuccessfulProcessingInd]  nchar(1)  DEFAULT 'N' NOT NULL
, CONSTRAINT [PK_dbo.DimAudit] PRIMARY KEY CLUSTERED 
( [AuditKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Audit', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimAudit
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Audit', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimAudit
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimAudit
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'Audit dimension tags each data row with the the process that added or updated it.', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimAudit
;

SET IDENTITY_INSERT dbo.DimAudit ON
;
INSERT INTO dbo.DimAudit (AuditKey, ParentAuditKey, TableName, PkgName, PkgGUID, PkgVersionGUID, PkgVersionMajor, PkgVersionMinor, ExecStartDT, ExecStopDT, ExecutionInstanceGUID, ExtractRowCnt, InsertRowCnt, UpdateRowCnt, ErrorRowCnt, TableInitialRowCnt, TableFinalRowCnt, TableMaxSurrogateKey, SuccessfulProcessingInd)
VALUES (-1, -1, 'Audit', 'None: Dummy row', NULL, NULL, NULL, NULL, '', '', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 'Y')
;
SET IDENTITY_INSERT dbo.DimAudit OFF
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'AuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'AuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ParentAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ParentAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TableName', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableName'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PkgName', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgName'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PkgGUID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgGUID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PkgVersionGUID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionGUID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PkgVersionMajor', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionMajor'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PkgVersionMinor', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionMinor'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ExecStartDT', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecStartDT'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ExecStopDT', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecStopDT'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ExecutionInstanceGUID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecutionInstanceGUID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ExtractRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExtractRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InsertRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'InsertRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UpdateRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'UpdateRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'ErrorRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ErrorRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TableInitialRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableInitialRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TableFinalRowCnt', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableFinalRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TableMaxSurrogateKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableMaxSurrogateKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'SuccessfulProcessingInd', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'SuccessfulProcessingInd'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'AuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Foreign key to self, to identify calling package execution', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ParentAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of the main table loaded by this package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableName'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of the SSIS package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgName'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Identifier for the package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgGUID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Identifier for the package version', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionGUID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Major version number for the package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionMajor'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Minor version number for the package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'PkgVersionMinor'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Date-time the package started executing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecStartDT'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Date-time the package finished executing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecStopDT'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Identifier for the execution of the package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExecutionInstanceGUID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of rows extracted from the source(s)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ExtractRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of rows inserted in the destination table', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'InsertRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of rows updated in the destination table', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'UpdateRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of error rows', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ErrorRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of rows in target table before we begin', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableInitialRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Count of rows in target table after package ends', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableFinalRowCnt'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Maximum surrogate key value in table (if we''re maintaining ourselves)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'TableMaxSurrogateKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Did the package finish executing successfully?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'SuccessfulProcessingInd'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'AuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ParentAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Y, N', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'SuccessfulProcessingInd'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'AuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimAudit', @level2type=N'COLUMN', @level2name=N'ParentAuditKey'; 
;





/* Drop table dbo.DimDate */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimDate') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimDate 
;

/* Create table dbo.DimDate */
CREATE TABLE dbo.DimDate (
   [TimeKey]  int   NOT NULL
,  [Date]  datetime   NOT NULL
,  [Month]  nvarchar(10)   NOT NULL
,  [MonthNumber]  int   NOT NULL
,  [Quarter]  nvarchar(10)   NOT NULL
,  [Year]  int   NOT NULL
,  [ISOWeekNumber]  tinyint   NOT NULL
, CONSTRAINT [PK_dbo.DimDate] PRIMARY KEY CLUSTERED 
( [TimeKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimDate
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Date', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimDate
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimDate
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'Date dimension contains one row for every day.', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimDate
;

INSERT INTO dbo.DimDate (TimeKey, Date, Month, MonthNumber, Quarter, Year, ISOWeekNumber)
VALUES (-1, '', '0', 1, '0', 0, 0)
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TimeKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'TimeKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DayOfWeek', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DayName', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DayOfMonth', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DayOfYear', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'WeekOfYear', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'TimeKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Full date as a SQL date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of the month', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Number of the month', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Quarter of the year', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Number of the year', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Week of year, 1..53', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'TimeKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/1/2016', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'January', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Q1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'2016', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'1', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'TimeKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Date'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Month', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Month'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Month Number', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'MonthNumber'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Quarter', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Quarter'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Year', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'Year'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'ISO Week Number', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimDate', @level2type=N'COLUMN', @level2name=N'ISOWeekNumber'; 
;





/* Drop table dbo.DimCustomers */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimCustomers') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimCustomers 
;

/* Create table dbo.DimCustomers */
CREATE TABLE dbo.DimCustomers (
   [CustomerKey]  int IDENTITY  NOT NULL
,  [CustomerID]  int   NOT NULL
,  [CustomerName]  nvarchar(40)   NOT NULL
,  [BillToCustomer]  nvarchar(30)   NOT NULL
,  [Category]  nvarchar(30)   NULL
,  [BuyingGroup]  nvarchar(30)   NOT NULL
,  [PrimaryContact]  nvarchar(30)  DEFAULT 'N/A' NOT NULL
,  [PostalCode]  varchar(20)   NOT NULL
,  [RowIsCurrent]  bit  DEFAULT 1 NOT NULL
,  [RowStartDate]  date  DEFAULT '12/31/1899' NOT NULL
,  [RowEndDate]  date  DEFAULT '12/31/9999' NOT NULL
,  [RowChangeReason]  nvarchar(200)   NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
, CONSTRAINT [PK_dbo.DimCustomers] PRIMARY KEY CLUSTERED 
( [CustomerKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCustomers
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Customers', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCustomers
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCustomers
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'Customers dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCustomers
;

SET IDENTITY_INSERT dbo.DimCustomers ON
;
INSERT INTO dbo.DimCustomers (CustomerKey, CustomerID, CustomerName, BillToCustomer, Category, BuyingGroup, PrimaryContact, PostalCode, RowIsCurrent, RowStartDate, RowEndDate, RowChangeReason, InsertAuditKey, UpdateAuditKey)
VALUES (-1, -1, 'No Customer', 'None', 'None', 'None', 'None', '-1', -1, '12/31/1899', '12/31/9999', 'N/A', -1, -1)
;
SET IDENTITY_INSERT dbo.DimCustomers OFF
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CustomerKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CustomerID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CustomerName', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'BillToCustomer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Category', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'BuyingGroup', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PrimaryContact', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'PostalCode', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Is Current', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Start Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row End Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Change Reason', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InsertAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UpdateAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Customer''s ID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Customer''s name and position', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Received name of Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Customer''s Category', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'The contact''s title at company', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Customer''s main contact', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Customer''s postal code', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Is this the current row for this member (Y/N)?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become valid for this member?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become invalid? (12/31/9999 if current row)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Why did the row change last?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process loaded this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process most recently updated this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3,…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Tailspin Toys (Head Office)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'CompanyA (Head Office)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Novelty Shop', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Novelty Shop', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Tailspin Toys', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'USA', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 0', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/24/2011', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/14/1998, 12/31/9999', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Customer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'CustomerKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'CustomerName', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'BillToCustomer', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Category', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'BuyingGroup', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'PrimaryContact', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'PostalCode', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'CustomerName'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BillToCustomer'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'Category'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'BuyingGroup'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PrimaryContact'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCustomers', @level2type=N'COLUMN', @level2name=N'PostalCode'; 
;





/* Drop table dbo.DimStateProvince */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimStateProvince') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimStateProvince 
;

/* Create table dbo.DimStateProvince */
CREATE TABLE dbo.DimStateProvince (
   [StateProvinceKey]  int IDENTITY  NOT NULL
,  [StateProvinceID]  int   NOT NULL
,  [Country]  nvarchar(40)   NOT NULL
,  [Continent]  nvarchar(30)   NULL
,  [SalesTerritory]  nvarchar(30)   NULL
,  [Region]  nvarchar(60)   NULL
,  [Subregion]  nvarchar(15)   NULL
,  [RowIsCurrent]  bit  DEFAULT 1 NOT NULL
,  [RowStartDate]  date  DEFAULT '12/31/1899' NOT NULL
,  [RowEndDate]  date  DEFAULT '12/31/9999' NOT NULL
,  [RowChangeReason]  nvarchar(200)   NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
, CONSTRAINT [PK_dbo.DimStateProvince] PRIMARY KEY CLUSTERED 
( [StateProvinceKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStateProvince
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvince', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStateProvince
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStateProvince
;

SET IDENTITY_INSERT dbo.DimStateProvince ON
;
INSERT INTO dbo.DimStateProvince (StateProvinceKey, StateProvinceID, Country, Continent, SalesTerritory, Region, Subregion, RowIsCurrent, RowStartDate, RowEndDate, RowChangeReason, InsertAuditKey, UpdateAuditKey)
VALUES (-1, -1, 'N/A', '', '', '', '', -1, '12/31/1899', '12/31/9999', 'N/A', -1, -1)
;
SET IDENTITY_INSERT dbo.DimStateProvince OFF
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvinceID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Country', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Continent', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'SalesTerritory', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Region', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Subregion', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Is Current', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Start Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row End Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Change Reason', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InsertAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UpdateAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'State Province ID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of country', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of continent', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of sales territory', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of region', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Name of subregion', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Is this the current row for this member (Y/N)?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become valid for this member?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become invalid? (12/31/9999 if current row)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Why did the row change last?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process loaded this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process most recently updated this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3,4...', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Lazada', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Nguyen Van A', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Sales Manager', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'123 Vo Van Ngan', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Vietnam', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 0', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/24/2011', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/14/1998, 12/31/9999', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'State Province Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Country', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Continent', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Sales Territory', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Region', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Subregion', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'StateProvinceID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Country'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Continent'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'SalesTerritory'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Region'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStateProvince', @level2type=N'COLUMN', @level2name=N'Subregion'; 
;





/* Drop table dbo.DimCity */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimCity') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimCity 
;

/* Create table dbo.DimCity */
CREATE TABLE dbo.DimCity (
   [CityKey]  int IDENTITY  NOT NULL
,  [CityID]  int   NOT NULL
,  [City]  nvarchar(40)   NOT NULL
,  [Location]  nvarchar(200)   NOT NULL
,  [StateProvinceKey]  int   NOT NULL
,  [RowIsCurrent]  bit  DEFAULT 1 NOT NULL
,  [RowStartDate]  date  DEFAULT '12/31/1899' NOT NULL
,  [RowEndDate]  date  DEFAULT '12/31/9999' NOT NULL
,  [RowChangeReason]  nvarchar(200)   NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
, CONSTRAINT [PK_dbo.DimCity] PRIMARY KEY CLUSTERED 
( [CityKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCity
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'City', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCity
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCity
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'City dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimCity
;

SET IDENTITY_INSERT dbo.DimCity ON
;
INSERT INTO dbo.DimCity (CityKey, CityID, City, Location, StateProvinceKey, RowIsCurrent, RowStartDate, RowEndDate, RowChangeReason, InsertAuditKey, UpdateAuditKey)
VALUES (-1, -1, 'N/A', '-1', -1, -1, '12/31/1899', '12/31/9999', 'N/A', -1, -1)
;
SET IDENTITY_INSERT dbo.DimCity OFF
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CityKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CityID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'City', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Location', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Is Current', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Start Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row End Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Change Reason', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InsertAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UpdateAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'ID of City', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'The name of supplying company', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Business key from source system (aka natural key)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'The name of each product', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Is this the current row for this member (Y/N)?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become valid for this member?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become invalid? (12/31/9999 if current row)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Why did the row change last?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process loaded this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process most recently updated this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Valhalla', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'POINT (-90.712611999999993 35.6323013)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3,…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 0', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/24/2011', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/14/1998, 12/31/9999', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'City', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'City', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'City', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'State Province', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'City Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'C', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Location', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'State Province Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'CityID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'City'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'Location'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimCity', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
;





/* Drop table dbo.DimStockItems */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimStockItems') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimStockItems 
;

/* Create table dbo.DimStockItems */
CREATE TABLE dbo.DimStockItems (
   [StockItemKey]  int IDENTITY  NOT NULL
,  [StockItemID]  int   NOT NULL
,  [StockItem]  nvarchar(200)   NOT NULL
,  [Color]  nvarchar(20)   NULL
,  [SellingPackage]  nvarchar(20)   NOT NULL
,  [BuyingPackage]  nvarchar(40)   NOT NULL
,  [Brand]  nvarchar(10)   NOT NULL
,  [Size]  nvarchar(20)   NOT NULL
,  [IsChillerStock]  nvarchar(20)  DEFAULT '1' NULL
,  [UnitPrice]  float   NULL
,  [RowIsCurrent]  bit  DEFAULT 1 NOT NULL
,  [RowStartDate]  date  DEFAULT '12/31/1899' NOT NULL
,  [RowEndDate]  date  DEFAULT '12/31/9999' NOT NULL
,  [RowChangeReason]  nvarchar(200)   NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
, CONSTRAINT [PK_dbo.DimStockItems] PRIMARY KEY CLUSTERED 
( [StockItemKey] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStockItems
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItems', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStockItems
exec sys.sp_addextendedproperty @name=N'Database Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStockItems
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'StockItems dimension', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=DimStockItems
;

SET IDENTITY_INSERT dbo.DimStockItems ON
;
INSERT INTO dbo.DimStockItems (StockItemKey, StockItemID, StockItem, Color, SellingPackage, BuyingPackage, Brand, Size, IsChillerStock, UnitPrice, RowIsCurrent, RowStartDate, RowEndDate, RowChangeReason, InsertAuditKey, UpdateAuditKey)
VALUES (-1, -1, '', '', '-1', 'Buying Package', '12/31/1899', '12/31/9999', 'N/A', -1, -1, '12/31/1899', '12/31/9999', 'N/A', -1, -1)
;
SET IDENTITY_INSERT dbo.DimStockItems OFF
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItemID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItem', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Color', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'SellingPackage', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'BuyingPackage', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Brand', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Size', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'IsChillerStock', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UnitPrice', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Is Current', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Start Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row End Date', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Row Change Reason', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InsertAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UpdateAuditKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Surrogate primary key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Stock Item ID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Stock Item Name', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Color of Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Selling Package unit', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Unit of Buying Package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Stock Item''s Name', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Stock Item''s Size', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Is Chiller Stock or not', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Price per unit of stock item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Is this the current row for this member (Y/N)?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become valid for this member?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'When did this row become invalid? (12/31/9999 if current row)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Why did the row change last?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process loaded this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'What process most recently updated this row?', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3…', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1,2,3,4...', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Animal with big feet slippers (Brown) XL', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Red', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Each', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Each', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'Northwind', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'XL', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 0', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'30', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 0', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/24/2011', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1/14/1998, 12/31/9999', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'Buying Package', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'SCD  Type', @value=N'n/a', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard SCD-2', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard Audit dim', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowIsCurrent'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowStartDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowEndDate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'RowChangeReason'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'Derived', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Stock Item', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Stock Item Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Stock Item Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Color', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'SellingPackage', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'BuyingPackage', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Brand', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Size', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Is Chiller Stock', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Unit Price', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItemID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'StockItem'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Color'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'SellingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'BuyingPackage'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Brand'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'nvarchar', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'Size'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'bit', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'IsChillerStock'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'float', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'DimStockItems', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
;





/* Drop table dbo.FactSales */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.FactSales') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.FactSales 
;

/* Create table dbo.FactSales */
CREATE TABLE dbo.FactSales (
   [SaleID]  int   NOT NULL
,  [CustomerKey]  int   NOT NULL
,  [CityKey]  int   NOT NULL
,  [StateProvinceKey]  int   NOT NULL
,  [StockItemKey]  int   NOT NULL
,  [WWIInvoiceID]  int   NOT NULL
,  [InvoiceDateKey]  int   NOT NULL
,  [DeliveryDateKey]  int   NOT NULL
,  [Quantity]  int   NOT NULL
,  [UnitPrice]  decimal   NOT NULL
,  [TotalExclTax]  decimal   NOT NULL
,  [TaxRate]  float   NOT NULL
,  [TaxAmount]  decimal   NOT NULL
,  [Profit]  decimal   NOT NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
,  [BKFactTable]  int   NULL
, CONSTRAINT [PK_dbo.FactSales] PRIMARY KEY NONCLUSTERED 
( [SaleID] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Fact', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactSales
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'FactSales', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactSales
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'Always create a table description, which becomes a table extended property.', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactSales
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'SaleID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CustomerKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CityKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'WWIInvoiceID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InvoiceDateKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InvoiceDateKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DeliveryDateKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'DeliveryDateKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Quantity', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'UnitPrice', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TotalExclTax', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TaxRate', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'TaxAmount', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Profit', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Insert Audit Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Update Audit Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'BK Fact Table', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Key to Audit dimension for row insertion', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Key to Audit dimension for row update', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'The natural key for the fact table, if any (eg order number)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard auditing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard auditing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimCustomers', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimCity', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimStateProvince', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimStockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'SaleID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'CustomerKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'CityKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'StockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'WWI Invoice ID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Quantity', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Unit Price', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Total Excluding Tax', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Tax Rate', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Tax Amount', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'Profit', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CustomerKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Quantity'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'decimal', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'UnitPrice'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'decimal', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TotalExclTax'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'float', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxRate'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'decimal', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'TaxAmount'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'decimal', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactSales', @level2type=N'COLUMN', @level2name=N'Profit'; 
;





/* Drop table dbo.FactOrderFullfill */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.FactOrderFullfill') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.FactOrderFullfill 
;

/* Create table dbo.FactOrderFullfill */
CREATE TABLE dbo.FactOrderFullfill (
   [SaleID]  int   NOT NULL
,  [CityKey]  int   NOT NULL
,  [StateProvinceKey]  int   NOT NULL
,  [StockItemKey]  int   NOT NULL
,  [WWIInvoiceID]  int   NOT NULL
,  [InvoiceDateKey]  int   NOT NULL
,  [DeliveryDateKey]  int   NOT NULL
,  [DeliveryTime]  int   NOT NULL
,  [InsertAuditKey]  int  DEFAULT -1 NOT NULL
,  [UpdateAuditKey]  int  DEFAULT -1 NOT NULL
,  [BKFactTable]  int   NULL
, CONSTRAINT [PK_dbo.FactOrderFullfill] PRIMARY KEY NONCLUSTERED 
( [SaleID] )
) ON [PRIMARY]
;

--Table extended properties...
exec sys.sp_addextendedproperty @name=N'Table Type', @value=N'Fact', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactOrderFullfill
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'FactOrderFullfill', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactOrderFullfill
exec sys.sp_addextendedproperty @name=N'Table Description', @value=N'Always create a table description, which becomes a table extended property.', @level0type=N'SCHEMA', @level0name=dbo, @level1type=N'TABLE', @level1name=FactOrderFullfill
;

--Column extended properties
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'SaleID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'CityKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'StockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'WWIInvoiceID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'InvoiceDateKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InvoiceDateKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DeliveryDateKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'DeliveryDateKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'DeliveryTime', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'DeliveryTime'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Insert Audit Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'Update Audit Key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Name', @value=N'BK Fact Table', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Key to Audit dimension for row insertion', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'Key to Audit dimension for row update', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Description', @value=N'The natural key for the fact table, if any (eg order number)', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Example Values', @value=N'1, 2, 3', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'key', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Display Folder', @value=N'Exclude from cube', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'BKFactTable'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard auditing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'InsertAuditKey'; 
exec sys.sp_addextendedproperty @name=N'ETL Rules', @value=N'Standard auditing', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'UpdateAuditKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source System', @value=N'WideWorldImporters', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Schema', @value=N'dbo', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimCity', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimStateProvince', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'DimStockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Table', @value=N'Sale', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'SaleID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'CityKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'StateProvinceKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'StockItemKey', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Field Name', @value=N'WWI Invoice ID', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'SaleID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'CityKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StateProvinceKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'StockItemKey'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'WWIInvoiceID'; 
exec sys.sp_addextendedproperty @name=N'Source Datatype', @value=N'int', @level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'FactOrderFullfill', @level2type=N'COLUMN', @level2name=N'DeliveryTime'; 
;
ALTER TABLE dbo.DimAudit ADD CONSTRAINT
   FK_dbo_DimAudit_ParentAuditKey FOREIGN KEY
   (
   ParentAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimCustomers ADD CONSTRAINT
   FK_dbo_DimCustomers_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimCustomers ADD CONSTRAINT
   FK_dbo_DimCustomers_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimStateProvince ADD CONSTRAINT
   FK_dbo_DimStateProvince_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimStateProvince ADD CONSTRAINT
   FK_dbo_DimStateProvince_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimCity ADD CONSTRAINT
   FK_dbo_DimCity_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimCity ADD CONSTRAINT
   FK_dbo_DimCity_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimStockItems ADD CONSTRAINT
   FK_dbo_DimStockItems_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.DimStockItems ADD CONSTRAINT
   FK_dbo_DimStockItems_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_CustomerKey FOREIGN KEY
   (
   CustomerKey
   ) REFERENCES DimCustomers
   ( CustomerKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_CityKey FOREIGN KEY
   (
   CityKey
   ) REFERENCES DimCity
   ( CityKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_StateProvinceKey FOREIGN KEY
   (
   StateProvinceKey
   ) REFERENCES DimStateProvince
   ( StateProvinceKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_StockItemKey FOREIGN KEY
   (
   StockItemKey
   ) REFERENCES DimStockItems
   ( StockItemKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_InvoiceDateKey FOREIGN KEY
   (
   InvoiceDateKey
   ) REFERENCES DimDate
   ( Timekey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_DeliveryDateKey FOREIGN KEY
   (
   DeliveryDateKey
   ) REFERENCES DimDate
   ( Timekey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_CityKey FOREIGN KEY
   (
   CityKey
   ) REFERENCES DimCity
   ( CityKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_StateProvinceKey FOREIGN KEY
   (
   StateProvinceKey
   ) REFERENCES DimStateProvince
   ( StateProvinceKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_StockItemKey FOREIGN KEY
   (
   StockItemKey
   ) REFERENCES DimStockItems
   ( StockItemKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_InvoiceDateKey FOREIGN KEY
   (
   InvoiceDateKey
   ) REFERENCES DimDate
   ( Timekey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_DeliveryDateKey FOREIGN KEY
   (
   DeliveryDateKey
   ) REFERENCES DimDate
   ( Timekey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_InsertAuditKey FOREIGN KEY
   (
   InsertAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
ALTER TABLE dbo.FactOrderFullfill ADD CONSTRAINT
   FK_dbo_FactOrderFullfill_UpdateAuditKey FOREIGN KEY
   (
   UpdateAuditKey
   ) REFERENCES DimAudit
   ( AuditKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
 
