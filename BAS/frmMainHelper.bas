Attribute VB_Name = "frmMainHelper"
'------------------------------------------------------------------------------
'Purpose  : frmMain handler methods
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Option Explicit
DefLng A-Z
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Public Sub MainSetup(ByVal frm As frmMain)
'------------------------------------------------------------------------------
'Purpose  : Initial Control setup
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 16.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

Call LVWTableSourceHeaderSetup(frm.lvwTableSource)
Call LVWColumnSourceHeaderSetup(frm.lvwColumnsSource)
Call LVWTableTargetHeaderSetup(frm.lvwTableTarget)

End Sub
'==============================================================================

Private Sub LVWTableSourceHeaderSetup(ByVal lvw As ListView)

With lvw.ColumnHeaders
   .Clear
   
   .Add , , "Table"
   .Add , , "Records", , lvwColumnRight
End With

lvw.ListItems.Clear

End Sub

Private Sub LVWColumnSourceHeaderSetup(ByVal lvw As ListView)

With lvw.ColumnHeaders
   .Clear
   
   .Add , , "Columns"
   .Add , , "Data type (source)"
   .Add , , "Size", , lvwColumnRight
   .Add , , "Precision", , lvwColumnRight
   .Add , , "AllowNULL", , lvwColumnCenter
End With

lvw.ListItems.Clear

End Sub

Private Sub LVWTableTargetHeaderSetup(ByVal lvw As ListView)

With lvw.ColumnHeaders
   .Clear
   
   .Add , , "Table"
   .Add , , "Records", , lvwColumnRight
End With

lvw.ListItems.Clear

End Sub
'==============================================================================

Public Function MainTransferStart(ByVal frm As frmMain, ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Handles data transfer startup, i.e. establishing DB connections
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

On Error GoTo MainTransferStartErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:MainTransferStart->"

' *** Establishing the database connections

' Try the (SQL) source server first ...
StatusMsg "Establishing database connection to source DBMS ..."

If Not OpenSource(oDB) Then
   Exit Function
End If

' ... now it's up to the (MS Access) target database
StatusMsg "Establishing database connection to target DBMS ..."

If Not OpenTarget(oDB) Then
   Exit Function
End If

' *** Clear control contents
ClearControls frm

' *** Verify source data
' Source tables - SQL
If VerifyTableSource(frmMain, frmMain.lvwTableSource, oDB) = False Then
   Exit Function
End If

MainTransferStart = True

'------------------
MainTransferStartErrExit:

On Error GoTo 0

Exit Function

'------------------
MainTransferStartErrHandler:

Call oDB.FireEvent(etOnDBEvent, _
   eBADbEvent.dbeErrOtherUnown, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description)

Err.Clear
Resume MainTransferStartErrExit

End Function
'==============================================================================

Public Function MainCopyData(ByVal frm As frmMain, ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Copy data from source to target
'
'Prereq.  : -
'Parameter: frm   - Instance of frmMain
'           oDB   - Current cDB object
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 09.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim i As Long
Dim sTable As String
Dim oLI As ListItem, oLITarget As ListItem

On Error GoTo MainCopyDataErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:MainCopyData->"

frm.lvwTableTarget.ListItems.Clear

' We've verified the tables' existance before, now run through the columns

For i = 1 To frm.lvwTableSource.ListItems.Count

   Set oLI = frm.lvwTableSource.ListItems(i)
   Set frm.lvwTableSource.SelectedItem = oLI
   sTable = oLI.Text

   If VerifyColumnSource(frmMain, sTable, frmMain.lvwColumnsSource, oDB) = False Then
      
      Exit Function
      
   Else
   
      If DropTableTarget(frm, oDB, sTable) = True Then
      ' We got rid off the table, recreate it ...
      
         StatusMsg "Creating target table " & sTable & " ..."
         
         If CreateTable(oDB.CnTarget, oDB, oDB.DBTables.DBTableGetByName(sTable)) = True Then
         ' Table creation succeeded, now actually copy the data over
         
            Set oLITarget = frm.lvwTableTarget.ListItems.Add(, sTable, sTable)
            Set frm.lvwTableTarget.SelectedItem = oLITarget
            
            StatusMsg "Copying data of table " & sTable & " ..."
            
            If CopyData(frm, oDB, sTable) = True Then
               
               oLITarget.ListSubItems(1).Text = Format$(frm.LastRecCount, "#,###,###,##0")
            
            Else
               
               oLITarget.ListSubItems(1).Text = "(failed)"
            
            End If
            
            ListViewAdjustColumnWidth frm.lvwTableTarget, , True, True
            
            Sleep 0: DoEvents
            If frm.AppState = asStopRequest Then
               Exit For
            End If
         
         End If
      
      End If
   
   End If

Next i

'------------------
MainCopyDataErrExit:

On Error GoTo 0
Exit Function

'------------------
MainCopyDataErrHandler:

oDB.FireEvent etOnDBEvent, eBADbEvent.dbeErrOtherUnown, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description

Err.Clear
Resume MainCopyDataErrExit

End Function
'==============================================================================

Public Function MainDBTestConnection(ByVal sConnection As String, ByRef sErrMsg As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Test if a DB connection could be established for the provided ADO
'           connection string
'
'Prereq.  : -
'Parameter: sConnection - ADO Connection string to test
'           sErrMsg     - (ByRef!) ADO error message, if connection fails.
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 22.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim cn As ADODB.Connection
Dim oErr As Object

On Error Resume Next

Set cn = New ADODB.Connection

cn.ConnectionString = sConnection
Call cn.Open

If Err Then
   If cn.State <> adStateClosed Then
      Call cn.Close
   End If
   sErrMsg = "Error: " & CStr(Err.Number) & ", " & Err.Description
   MainDBTestConnection = False
ElseIf cn.Errors.Count > 0 Then
   For Each oErr In cn.Errors
      sErrMsg = sErrMsg & "Error: " & CStr(oErr.Number) & ", " & oErr.Description & vbNewLine
   Next oErr
   Call cn.Close
   MainDBTestConnection = False
   Call cn.Errors.Clear
Else
   Call cn.Close
   MainDBTestConnection = True
End If

On Error GoTo 0

Set cn = Nothing

End Function
'==============================================================================

Private Function OpenSource(ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Opens the ADO source connection
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim sTemp As String

On Error GoTo OpenSourceErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:OpenSource->"

With oDB

   Set .CnSource = New ADODB.Connection
   .CnSource.ConnectionString = .ConnectionSource
   Call .CnSource.Open

End With

OpenSource = True

'------------------
OpenSourceErrExit:

On Error GoTo 0
Exit Function

'------------------
OpenSourceErrHandler:

sTemp = ADOGetConnectionErrStr(oDB.CnSource)

If Len(sTemp) > 0 Then
   sTemp = PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description & vbNewLine & sTemp
Else
   sTemp = PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
End If

Call oDB.FireEvent(etOnDBEvent, _
   eBADbEvent.dbeConnectionSourceFailed, sTemp)

Err.Clear
Resume OpenSourceErrExit

End Function
'==============================================================================

Private Function VerifyTableSource(ByVal frm As frmMain, ByVal lvw As ListView, ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Verifies that all necessary tables exist in the source database
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable
Dim sTemp As String

VerifyTableSource = True

For Each oTable In oDB.DBTables.DBTables
   sTemp = oTable.TblName
   StatusMsg "Verifying source table existance " & sTemp & " ..."
   If DBADOUtil.DBADOTableExistsCN(oDB.CnSource, sTemp) = teTableExists Then
      Call lvw.ListItems.Add(, sTemp, sTemp)
   Else
      Call lvw.ListItems.Add(, sTemp, sTemp & " (missing)")
      VerifyTableSource = VerifyTableSource And False
   End If
Next oTable

ListViewAdjustColumnWidth lvw, , True, True

End Function

Private Function VerifyColumnSource(ByVal frm As frmMain, ByVal sTable As String, _
   ByVal lvwColumns As ListView, ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Verifies that all necessary columns exist in the source database
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: 20.10.2015
'           - Allow for fetching all columns
'------------------------------------------------------------------------------
Dim oTable As cDBTable
Dim oCol As cDBColumn
Dim sTemp As String
Dim i As Long
Dim oLI As ListItem
Dim oMap As cDBMapping

Dim eType As ADOX.DataTypeEnum, lSize As Long, lPrecision As Long

On Error GoTo VerifyColumnSourceErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:VerifyColumnSource->"

VerifyColumnSource = True

lvwColumns.ListItems.Clear

' We've verified the existance of the table before, so pick
' the table names from the listbox.

Set oTable = oDB.DBTables.DBTableGetByName(sTable)

' Retrieve all columns?
If oTable.DBColumns.CopyAllColumns = True Then

   Dim cat As ADOX.Catalog
   Dim tbl As ADOX.Table
   Dim colADO As ADOX.Column
   
   Set cat = New ADOX.Catalog
   Set cat.ActiveConnection = oDB.CnSource
   
   For Each tbl In cat.Tables
   
      If tbl.Name = sTable Then
         
         For Each colADO In tbl.Columns
         
         sTemp = colADO.Name
         StatusMsg "Table: " & sTable & ": adding source column information " & sTemp & " ..."
         
         ' Add column name and data type to ListView
         ' Name
         Set oLI = lvwColumns.ListItems.Add(, sTemp, sTemp)
            
            Set oCol = New cDBColumn
            
            With oCol
               .ColName = colADO.Name
               Set .ADOXColumnSource = colADO
               If colADO.Attributes = (colADO.Attributes Or adColNullable) Then
                  .AllowNullStr = "True"
               End If
               .Precision = colADO.Precision
               .Size = colADO.DefinedSize
               If oDB.DBMappings.HasDBMappingOfType(colADO.Type) = True Then
                  Set oMap = oDB.DBMappings.GetDBMappingForType(colADO.Type)
                  .TypeTarget = oMap.TypeTarget
               Else
                  .TypeTarget = colADO.Type
               End If
            End With
            
            Call oDB.DBTables.DBTableAddCol(oTable, oCol)
            
            ' Column properties (source)
            sTemp = modApp.GetFieldTypeEx(oCol.ADOXColumnSource, eType, lSize, lPrecision)
            
            ' Data type (source)
            sTemp = sTemp & "(" & CStr(eType) & ")"
            oLI.ListSubItems.Add , , sTemp
         
            ' Size
            oLI.ListSubItems.Add , , Format$(lSize, "#,###,##0")
            ' Precision
            oLI.ListSubItems.Add , , Format$(lPrecision, "##0")
            ' AllowNULL
            oLI.ListSubItems.Add , , IIf(oCol.AllowNull = True, "True", "False")
         
         Next colADO
         
      End If
   
   Next tbl
   
   Set tbl = Nothing
   Set cat.ActiveConnection = Nothing
   Set cat = Nothing

Else     '// If oTable.DBColumns.CopyAllColumns = True

   For Each oCol In oTable.DBColumns.DBColumns
      
      sTemp = oCol.ColName
      StatusMsg "Table: " & sTable & ": verifying source column existance " & sTemp & " ..."
      
      ' Add column name and data type to ListView
      ' Name
      Set oLI = lvwColumns.ListItems.Add(, sTemp, sTemp)
      
      If DBADOUtil.DBADOColumnExistsCN(oDB.CnSource, sTable, sTemp) = ceColExists Then
         ' Set source column' ADOX Column
         Set oCol.ADOXColumnSource = DBADOUtil.DBADOColumnGetADOXColCN(oDB.CnSource, sTable, sTemp)
         Call oDB.DBTables.DBTableAddADOXCol(oTable, oCol)
         
         ' Column properties (source)
         sTemp = modApp.GetFieldTypeEx(oCol.ADOXColumnSource, eType, lSize, lPrecision)
         
         ' Data type (source)
         sTemp = sTemp & "(" & CStr(eType) & ")"
         oLI.ListSubItems.Add , , sTemp
         
         ' Size
         oLI.ListSubItems.Add , , Format$(lSize, "#,###,##0")
         ' Precision
         oLI.ListSubItems.Add , , Format$(lPrecision, "##0")
         ' AllowNULL
         oLI.ListSubItems.Add , , IIf(oCol.AllowNull = True, "True", "False")
      
      Else
         
         oLI.ListSubItems.Add , , "(missing)"
         
         VerifyColumnSource = VerifyColumnSource And False
         oDB.FireEvent etOnDBEvent, eBADbEvent.dbeColSourceMissing, sTemp
      End If
      
   Next oCol

End If   '// If oTable.DBColumns.CopyAllColumns = True

ListViewAdjustColumnWidth lvwColumns, , True, True

' After we're successfully verified the source columns, iterate through the default data type mappings
' and set each target column's data type accordingly.
If VerifyColumnSource = True Then

   oDB.FireEvent etOnDBEvent, eBADbEvent.dbeColDoMapping, sTable

End If

'------------------
VerifyColumnSourceErrExit:

On Error GoTo 0
Exit Function

'------------------
VerifyColumnSourceErrHandler:
VerifyColumnSource = False

oDB.FireEvent etOnDBEvent, eBADbEvent.dbeErrOtherUnown, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description

Err.Clear
Resume VerifyColumnSourceErrExit

End Function

Private Function OpenTarget(ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Opens the ADO target connection
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim sTemp As String

On Error GoTo OpenTargetErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:OpenTarget->"

With oDB

   Set .CnTarget = New ADODB.Connection
   .CnTarget.ConnectionString = .ConnectionTarget
   Call .CnTarget.Open

End With

OpenTarget = True

'------------------
OpenTargetErrExit:

On Error GoTo 0
Exit Function

'------------------
OpenTargetErrHandler:

sTemp = ADOGetConnectionErrStr(oDB.CnTarget)

If Len(sTemp) > 0 Then
   sTemp = PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description & vbNewLine & sTemp
Else
   sTemp = PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
End If

Call oDB.FireEvent(etOnDBEvent, _
   eBADbEvent.dbeConnectionTargetFailed, sTemp)

Err.Clear
Resume OpenTargetErrExit

End Function
'==============================================================================

Private Function VerifyTableTarget(ByVal frm As frmMain, ByVal lst As ListBox, ByVal oDB As cDB) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Verifies that all necessary tables exist in the SQL database
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable
Dim sTemp As String

VerifyTableTarget = True

For Each oTable In oDB.DBTables.DBTables
   sTemp = oTable.TblName
   StatusMsg "Verifying target table existance " & sTemp & " ..."
   If DBADOUtil.DBADOTableExistsCN(oDB.CnTarget, sTemp) = teTableExists Then
      lst.AddItem sTemp
   Else
      lst.AddItem sTemp & " (missing)"
      VerifyTableTarget = VerifyTableTarget And False
   End If
Next oTable

End Function
'==============================================================================

Private Function DropTableTarget(ByVal frm As frmMain, ByVal oDB As cDB, ByVal sTable As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Drops the passed table from the database
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 12.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable

On Error GoTo DropTableTargetErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:DropTableTarget->"

For Each oTable In oDB.DBTables.DBTables

   If oTable.TblName = sTable Then
      
      If DBADOTableExistsCN(oDB.CnTarget, oTable.TblName) = False Then
      ' Table doesn't even exist in target database, DROP TABLE was "successful"
         DropTableTarget = True
         Exit Function
      Else
      ' Table does exist -> drop it.
         Call oDB.CnTarget.Execute("DROP TABLE " & sTable, , adExecuteNoRecords)
         Exit For
      End If
      
   End If

Next oTable

' If reaching thus far, table either didn't exist in the first place or was successfully dropped
DropTableTarget = True

'------------------
DropTableTargetErrExit:

On Error GoTo 0
Exit Function

'------------------
DropTableTargetErrHandler:

Err.Clear
Resume DropTableTargetErrExit

End Function
'==============================================================================

Private Function CreateTable(ByVal cn As ADODB.Connection, ByVal oDB As cDB, ByVal oTable As cDBTable) As Boolean
'------------------------------------------------------------------------------
'Purpose  : (Re-)Create a table in a database according to the XML configuration
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 13.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table
Dim col As ADOX.Column

Dim oDBCol As cDBColumn
Dim sColType As String

Dim sMsg As String

On Error GoTo CreateTableErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:CreateTable->"

Set cat = New ADOX.Catalog
Set cat.ActiveConnection = cn

Set tbl = New ADOX.Table

With tbl

   .Name = oTable.TblName
   
   ' Add the necessary columns to the table
   For Each oDBCol In oTable.DBColumns.DBColumns
   
         StatusMsg "Creating target column " & oDBCol.ColName & " ..."
         
         Set col = New ADOX.Column
         col.Name = oDBCol.ColName
         col.Type = oDBCol.GetTypeTarget
         
         If (oDBCol.Size = 0) And (oDBCol.Precision = 0) Then
         ' Integer types
         ElseIf (oDBCol.Size > 0) And (oDBCol.Precision = 0) Then
         ' String types
            col.DefinedSize = oDBCol.Size
         ElseIf (oDBCol.Size = 0) And (oDBCol.Precision > 0) Then
         ' Float types
            If DBADOColumnTypeIsInteger(col.Type) = False Then
               col.Precision = oDBCol.Precision
            End If
         End If
         
         ' Allow null?
         If oDBCol.AllowNull = True Then
         ' From XML definition
            col.Attributes = adColNullable
         End If
         
         ' Regardless of AllowNull, the following target column types should not be NULL
         ' adBoolean can never be NULL
         Select Case oDBCol.TypeTarget
         Case adoBoolean
            col.Attributes = 0
         End Select
         
         sColType = modApp.GetFieldType(col)
         .Columns.Append col
   
   Next oDBCol
   
End With

sMsg = GetDBColumnsPropertiesString(tbl)

' After adding the columns, finish by adding the table to the DB
cat.Tables.Append tbl

sMsg = vbNullString

' Add index, if necessary
For Each oDBCol In oTable.DBColumns.DBColumns

   If oDBCol.IsIndex = True Then
      If AddIndex(oDB, cat, tbl, oDBCol) = False Then
         CreateTable = False
         Exit Function
      End If
   End If

Next oDBCol

Set cat.ActiveConnection = Nothing
Set cat = Nothing
Set tbl = Nothing
Set col = Nothing

CreateTable = True

'------------------
CreateTableErrExit:

On Error GoTo 0
Exit Function

'------------------
CreateTableErrHandler:

CreateTable = False

If Len(sMsg) > 0 Then
   sMsg = Err.Description & vbNewLine & sMsg
Else
   sMsg = Err.Description
End If

Call oDB.FireEvent(etOnDBEvent, eBADbEvent.dbeTableCreateFailed, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & sMsg)

Err.Clear
Resume CreateTableErrExit

End Function
'==============================================================================

Private Function AddIndex(ByVal oDB As cDB, ByVal cat As ADOX.Catalog, ByVal tbl As ADOX.Table, ByVal oDBCol As cDBColumn) As Boolean

Dim i As Long
Dim col As ADOX.Column
Dim idx As ADOX.Index
Dim bolFound As Boolean

On Error GoTo AddIndexErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:AddIndex->"

' Retrieve the index' column fomr the table definition
For i = 0 To tbl.Columns.Count - 1
   Set col = tbl.Columns(i)
   If col.Name = oDBCol.ColName Then
      bolFound = True
      Exit For
   End If
Next i

If bolFound = True Then

   Set idx = New ADOX.Index
   
   With idx
      .Name = oDBCol.ColName
      ' Primary index?
      .PrimaryKey = oDBCol.IsPrimary
      If .PrimaryKey = False Then
         .Unique = oDBCol.IsUnique
         If .Unique = False Then
            ' AllowNull
            If oDBCol.AllowNull = True Then
               .IndexNulls = adIndexNullsIgnore
            End If
         End If
      End If
      
      .Columns.Append col.Name
   End With
   
   tbl.Indexes.Append idx
   
End If

AddIndex = True

'------------------
AddIndexErrExit:

On Error GoTo 0
Exit Function

'------------------
AddIndexErrHandler:

AddIndex = False

Call oDB.FireEvent(etOnDBEvent, eBADbEvent.dbeIndexCreateFailed, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description)

Err.Clear
Resume AddIndexErrExit

End Function
'==============================================================================

Private Function CopyData(ByVal frm As frmMain, ByVal oDB As cDB, ByVal sTable As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Copy data from source DB/table to target DB/table
'
'Prereq.  : -
'Parameter: frm      - Calling form for providing visual feedback
'           oDB      - Root DB handling object
'           sTable   - Table name to prcess
'Returns  : -
'Note     : http://forum.winbatch.com/index.php?topic=736.0
'           https://support.microsoft.com/en-us/kb/200427
'
'   Author: Knuth Konrad 14.10.2015
'   Source: -
'  Changed: 22.10.2015
'           - Handle BIT (SQL) / YesNo (Access) values where the former allows
'           NULL, the later not.
'           30.10.2015
'           - Enclose column names in brackets ([]) to escape potential
'           naming conflicts
'           11.01.2016
'           - FIX: Handling of SQL BOOL -> Access Booelan
'------------------------------------------------------------------------------
Dim sSQLSource As String, sSQLTarget As String
Dim sSQLCount As String
Dim lRecCount As Long, lCount As Long, lColCount As Long, i As Long
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

Dim sParam As String

Dim rs As ADODB.Recordset

On Error GoTo CopyDataErrHandler

Const PROCEDURE_NAME As String = "frmMainHelper:CopyData->"

sParam = "n/a"

' Create the query SQL statement
Dim oDBTable As cDBTable, oDBCol As cDBColumn

Set cmd = New ADODB.Command
With cmd
   Set .ActiveConnection = oDB.CnTarget
   .CommandType = adCmdText
   .Prepared = True
End With

Set oDBTable = oDB.DBTables.DBTableGetByName(sTable)
sSQLSource = vbNullString: sSQLTarget = vbNullString

lColCount = oDBTable.DBColumnsCount
For i = 1 To lColCount
   
   Set oDBCol = oDBTable.DBColumns.DBColumns(i)
   
   Set prm = New ADODB.Parameter
   
   With prm
      .Name = oDBCol.ColName
      sParam = .Name
      .Direction = adParamInput
      .Type = oDBCol.GetTypeTarget
      If oDBCol.Size > 0 Then
         .Size = oDBCol.Size
      End If
      If oDBCol.Precision > 0 Then
         .Precision = oDBCol.Precision
      End If
   End With
   
   cmd.Parameters.Append prm
   
   sSQLSource = sSQLSource & "[" & oDBCol.ColName & "]"
   sSQLTarget = sSQLTarget & "@" & oDBCol.ColName

   If i < lColCount Then
       sSQLSource = sSQLSource & ", "
       sSQLTarget = sSQLTarget & ", "
   End If

Next i

sParam = "n/a"

' SQL Record count statement - for informational purpose only
sSQLCount = "SELECT COUNT(*) As RecCount FROM " & sTable
If Len(oDBTable.Query) > 0 Then
   sSQLCount = sSQLCount & " WHERE " & oDBTable.Query
End If
sSQLCount = sSQLCount & ";"

' SQL INSERT for copying data
sSQLTarget = "INSERT INTO " & sTable & " (" & sSQLSource & ") VALUES(" & sSQLTarget & ");"

' SQL - selecting the source records
sSQLSource = "SELECT " & sSQLSource & " FROM " & sTable
If Len(oDBTable.Query) > 0 Then
   sSQLSource = sSQLSource & " WHERE " & oDBTable.Query
End If
sSQLSource = sSQLSource & ";"

Set rs = New ADODB.Recordset
Call rs.Open(sSQLCount, oDB.CnSource, adOpenForwardOnly, adLockReadOnly)

If Not rs Is Nothing Then

   ' Get number of records in source
   lRecCount = 0: lCount = 0: cmd.CommandText = sSQLTarget
   
   rs.MoveFirst
   lRecCount = rs.Fields("RecCount").Value
   Call rs.Close
   
   Call oDB.FireEvent(etOnDBEvent, eBADbEvent.dbeRecordCount, lRecCount)
   
   ' Retrieve the actual data
   Call rs.Open(sSQLSource, oDB.CnSource, adOpenForwardOnly, adLockReadOnly)
   rs.MoveFirst
   
   Do
   
      lCount = lCount + 1
      oDB.FireEvent etOnDBEvent, eBADbEvent.dbeRecordAdded, lCount
      
      ' Set the target column's value
      For Each prm In cmd.Parameters
      
' { --- DEBUG --- 20.10.2015
'      Debug.Print prm.Name
'      Debug.Print rs.Fields(prm.Name).Type
' } --- DEBUG --- 20.10.2015
         
         Select Case rs.Fields(prm.Name).Type
         Case ADODB.DataTypeEnum.adBoolean
         ' While SQL's "bool" (= bit) allows NULL values, Access's YesNo doesn't
            prm.Type = rs.Fields(prm.Name).Type
            If Not IsNull(rs.Fields(prm.Name).Value) Then
               cmd.Parameters(prm.Name).Value = CBool(rs.Fields(prm.Name).Value)
            Else
               cmd.Parameters(prm.Name).Value = False
            End If
         Case Else
            cmd.Parameters(prm.Name).Value = rs.Fields(prm.Name).Value
         End Select
      Next prm
      
      sParam = GetDBParametersString(cmd)
' { --- DEBUG --- 20.10.2015
'      Debug.Print sSQLTarget
'      Debug.Print sParam
' } --- DEBUG --- 20.10.2015
     
      Call cmd.Execute(, , adExecuteNoRecords)
      
      rs.MoveNext
      
      ' Check for user cancellation
      If lCount Mod 100 = 0 Then
         Sleep 0: DoEvents
         If frm.AppState = asStopRequest Then
            Exit Do
         End If
      End If
   
   Loop Until rs.EOF

End If

rs.Close
Set rs = Nothing

Set cmd = Nothing

CopyData = True

'------------------
CopyDataErrExit:

On Error GoTo 0
Exit Function

'------------------
CopyDataErrHandler:

CopyData = False

oDB.FireEvent etOnDBEvent, eBADbEvent.dbeErrOtherUnown, _
   PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description & ". Param: " & sParam

Err.Clear
Resume CopyDataErrExit

End Function
'==============================================================================

Private Sub ClearControls(ByVal frm As frmMain)
'------------------------------------------------------------------------------
'Purpose  : Clears the contents of the GUI controls (Listboxes etc.)
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 09.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

With frm

   .lvwColumnsSource.ListItems.Clear
   .lvwTableSource.ListItems.Clear
   .lvwTableTarget.ListItems.Clear

End With

End Sub
'==============================================================================


