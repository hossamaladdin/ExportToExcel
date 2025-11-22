# ExportToExcel

[![SQL Server](https://img.shields.io/badge/SQL%20Server-2019%2B-blue.svg)](https://www.microsoft.com/en-us/sql-server/sql-server-downloads)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)

## Overview

A powerful SQL Server stored procedure that integrates T-SQL and PowerShell to export tables or data to Excel/CSV and deliver via email or SharePoint upload. Ideal for DBAs automating report distribution from remote servers.

## Features

- **Export Formats**: CSV or Excel (using ImportExcel PowerShell module)
- **Data Sources**: Supports user tables, system tables, or global temp tables (##TableName)
- **Delivery Options**:
  - Email attachment (using Database Mail)
  - SharePoint/OneDrive upload (using Microsoft.Online.SharePoint.PowerShell)
- **Automation**: Single procedure call handles export, formatting, and delivery
- **Flexible Output**: Customizable file paths, names, and Export-Excel parameters
- **Permissions**: Automatically enables/disables xp_cmdshell during execution

## Prerequisites

- SQL Server 2019 or later (tested with BCP utility)
- PowerShell 5.1 or later
- `ImportExcel` PowerShell module (auto-installs if missing)
- `Microsoft.Online.SharePoint.PowerShell` (for SharePoint uploads)
- Database Mail configured (for email)
- Appropriate permissions:
  - Execute permissions on xp_cmdshell
  - Access to BCP utility
  - File system access to temp directories
  - SharePoint/OneDrive credentials if uploading

## Installation

1. Execute `ExporttoExcel_1.0.sql` to create the stored procedure
2. For SharePoint integration, ensure PowerShell modules are installed and accessible
3. Configure Database Mail for email features
4. Test with the provided example script

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| @databasename | varchar(20) | NULL | Database name (omit for tempdb tables) |
| @schemaname | varchar(20) | NULL | Schema name (omit for temp tables) |
| @tablename | varchar(20) | Required | Table name (or ##globaltemp for temp tables) |
| @filepath | varchar(100) | NULL | Output directory (defaults to %TEMP%) |
| @filename | varchar(100) | NULL | Output filename (defaults to table name) |
| @fullname | varchar(100) | Output | Full path of exported file |
| @CSVOnly | bit | 0 | Export to CSV only (no Excel conversion) |
| @params | varchar(100) | '' | Export-Excel parameters (e.g., '-AutoSize -TableName') |
| @overwrite | bit | 1 | Overwrite existing files |
| @deleteold | bit | 0 | Clean old .xlsx files in temp folder |
| @uploadonly | bit | 0 | Upload existing file without re-exporting |
| @uploadfile | bit | 0 | Upload to SharePoint (1 = yes) |
| @SharePointUrl | varchar(100) | '' | SharePoint site URL |
| @Library | varchar(100) | '' | SharePoint library/folder name |
| @Username | varchar(100) | '' | SharePoint username |
| @Password | varchar(100) | '' | SharePoint password |
| @AttachToMail | bit | 0 | Attach file to email (1 = yes) |
| @MailList | varchar(100) | '' | Email recipients (semicolon-separated) |
| @MailSubject | varchar(100) | 'Subject' | Email subject |
| @MailBody | varchar(100) | 'PFA' | Email body |

## Usage Examples

### Basic Excel Export
```sql
EXEC ExportToExcel @tablename = 'MyTable'
```

### CSV Export to Custom Location
```sql
EXEC ExportToExcel
   @tablename = 'MyTable',
   @filepath = 'C:\Exports\',
   @filename = 'Report.csv',
   @CSVOnly = 1
```

### Export Temp Table and Email
```sql
-- Create temp table
SELECT column1, column2 INTO ##TempExport FROM MyTable

EXEC ExportToExcel
   @tablename = '##TempExport',
   @AttachToMail = 1,
   @MailList = 'recipient@domain.com',
   @MailSubject = 'Monthly Report',
   @MailBody = 'Please find attached the monthly report.'
```

### Export and Upload to SharePoint
```sql
EXEC ExportToExcel
   @tablename = 'MyTable',
   @uploadfile = 1,
   @SharePointUrl = 'https://company.sharepoint.com/sites/reports',
   @Library = 'Monthly',
   @Username = 'user@company.com',
   @Password = 'MyPassword'
```

See `ExporttoExcel_1.0_Example.sql` for a complete working example.

## Security Notes

- The procedure temporarily enables xp_cmdshell during execution and disables it afterward
- Ensure proper permissions are granted to the executing user
- Store SharePoint passwords securely and consider using secure strings

## Acknowledgments

Created by Hossam Alaa (hossam.alaaddin@gmail.com)
- Inspired by Doug Finke's [ImportExcel](https://github.com/dfinke/ImportExcel) module
