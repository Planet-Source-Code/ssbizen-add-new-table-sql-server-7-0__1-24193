<div align="center">

## Add New Table \(SQL Server 7\.0\)


</div>

### Description

This code lets you add a new table to existing database in SQL Server programmatically. Useful when developing a database application for off-line users. Simply send an executable through e-mail, and let it run once.
 
### More Info
 
Set DB name, DB file name, User login, Password(if required), name of new table, and fields. Added table can be easily removed from SQL Server Enterprise Manager.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SSBizen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ssbizen.md)
**Level**          |Advanced
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ssbizen-add-new-table-sql-server-7-0__1-24193/archive/master.zip)

### API Declarations

```
Microsoft SQLDMO Object Library(SQLDMO.RLL)
Microsoft ADO 2.5
```


### Source Code

```
Option Explicit
Const DATABASE = "*" 'Enter name of the database here
Const DBFILE_LOC = "C:\MSSQL7\DATA\*_DATA.mdf" 'Physical path
Const USER = "*" 'User name for login
Const PASSWORD = "*" 'Password
Const TABLE = "*" 'Name of the new table
Const COLUMN1 = "*" 'Field#1 name
Const COLUMN2 = "*" 'Field#2 name
Sub Main()
Dim oSQLServer As SQLDMO.SQLServer, oDatabase As SQLDMO.DATABASE
Dim tblNewTable As New SQLDMO.TABLE
Dim colNewColumn1 As New SQLDMO.Column, colNewColumn2 As New SQLDMO.Column
On Error GoTo Errors
 Set oSQLServer = New SQLDMO.SQLServer
 oSQLServer.Connect , "sa" 'Use USER/PASSWORD if neccessary
 Set oDatabase = oSQLServer.Databases(DATABASE)
 'Populate the Column objects to define
 'the table columns.
 colNewColumn1.Name = COLUMN1
 colNewColumn1.Datatype = "decimal"
 colNewColumn1.Length = 5
 colNewColumn1.NumericPrecision = 3
 colNewColumn1.NumericScale = 0
 colNewColumn1.AllowNulls = False
 colNewColumn2.Name = COLUMN2
 colNewColumn2.Datatype = "datetime"
 colNewColumn2.Length = 8
 colNewColumn2.AllowNulls = True
 'Name the table, then set desired properties
 'to control eventual table construction
 tblNewTable.Name = TABLE
 tblNewTable.FileGroup = "PRIMARY"
 'Add column objects to the Columns collection
 tblNewTable.Columns.Add colNewColumn1
 tblNewTable.Columns.Add colNewColumn2
 'Create the table by adding the
 'Table object to its containing collection.
 oDatabase.Tables.Add tblNewTable
 Exit Sub
Errors:
 ErrorHandler ("Main")
End Sub
Sub ErrorHandler(ByVal strLocation As String)
 If Err.Number <> 0 Then
 MsgBox "Error #: " & Str(Err.Number) & vbCrLf & _
 "Description: " & Err.Description & vbCrLf & _
 "Source: " & Err.Source, _
 vbCritical + vbSystemModal, "CreateTable: " & strLocation
 End If
End Sub
```

