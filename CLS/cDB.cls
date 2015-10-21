VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : XML Config parent object
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 07.10.2015
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
Public Enum eBAEventType
   etOnXMLParse
   etOnDataChange
   etOnDBEvent
End Enum

Public Enum eBAParseEvent
   xmlErrOtherUnown = -1
   xmlSuccess = 0
   xmlNoContent
   xmlMissingConnectionSource
   xmlMissingConnectionTarget
   xmlMissingTablesNode
   xmlMissingMappingNode
   xmlTableDefAlreadyExists
End Enum

Public Enum eBADataChange
   dcConnectionSource
   dcConnectionTarget
End Enum

Public Enum eBADbEvent
   dbeErrOtherUnown = -1
   dbeSuccess
   dbeConnectionSourceFailed
   dbeConnectionTargetFailed
   dbeTblSourceMissing
   dbeColSourceMissing
   dbeColDoMapping
   dbeTableCreateFailed
   dbeIndexCreateFailed
   dbeRecordCount
   dbeRecordAdded
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Private msConnectionSource As String
Private msConnectionTarget As String

Private moDBMappings As cDBMappings
Private WithEvents moDBTables As cDBTables
Attribute moDBTables.VB_VarHelpID = -1

Private WithEvents moCnSource As ADODB.Connection
Attribute moCnSource.VB_VarHelpID = -1
Private WithEvents moCnTarget As ADODB.Connection
Attribute moCnTarget.VB_VarHelpID = -1
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
Public Event OnXMLParse(ByVal eEvent As eBAParseEvent, ByVal vntValue As Variant)
Public Event OnDataChange(ByVal eEvent As eBADataChange, ByVal vntValue As Variant)
Public Event OnDBEvent(ByVal eEvent As eBADbEvent, ByVal vntValue As Variant)
'==============================================================================

Friend Property Get ConnectionSource() As String
   ConnectionSource = msConnectionSource
End Property

Friend Property Let ConnectionSource(ByVal sValue As String)
   msConnectionSource = sValue
End Property

Friend Property Get ConnectionTarget() As String
   ConnectionTarget = msConnectionTarget
End Property

Friend Property Let ConnectionTarget(ByVal sValue As String)
   msConnectionTarget = sValue
End Property

Friend Property Get DBTables() As cDBTables
   Set DBTables = moDBTables
End Property

Friend Property Set DBTables(ByVal oValue As cDBTables)
   Set moDBTables = oValue
End Property

Friend Property Get DBMappings() As cDBMappings
   Set DBMappings = moDBMappings
End Property

Friend Property Set DBMappings(ByVal oValue As cDBMappings)
   Set moDBMappings = oValue
End Property

Friend Property Get CnSource() As ADODB.Connection
   Set CnSource = moCnSource
End Property

Friend Property Set CnSource(ByVal oValue As ADODB.Connection)
   Set moCnSource = oValue
End Property

Friend Property Get CnTarget() As ADODB.Connection
   Set CnTarget = moCnTarget
End Property

Friend Property Set CnTarget(ByVal oValue As ADODB.Connection)
   Set moCnTarget = oValue
End Property
'==============================================================================

Friend Sub ColumnAddSourceType(ByVal oDBTable As cDBTable, ByVal oDBCol As cDBColumn, ByVal eType As eADOXDataTypeEnum)
'------------------------------------------------------------------------------
'Purpose  : Removes a Column from the specified Table's column collection and adds it
'           back again with new SourceType value.
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 13.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable, oCol As cDBColumn, i As Long

With Me
   For Each oTable In .DBTables.DBTables
      If oTable.TblName = oDBTable.TblName Then
         For i = 1 To oTable.DBColumns.DBColumns.Count
         Set oCol = oTable.DBColumns.DBColumns(i)
            If oCol.ColName = oDBCol.ColName Then
               oCol.TypeSource = eType
               Exit Sub
            End If
         Next i
      End If
   Next oTable
End With

End Sub

Friend Sub ColumnAddADOXCol(ByVal oDBTable As cDBTable, ByVal oDBCol As cDBColumn)
'------------------------------------------------------------------------------
'Purpose  : Removes a Column from the specified Table's column collection and adds it
'           back again with new ADOXCol object.
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 14.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable, oCol As cDBColumn, i As Long

With Me
   For Each oTable In .DBTables.DBTables
      If oTable.TblName = oDBTable.TblName Then
         For i = 1 To oTable.DBColumns.DBColumns.Count
         Set oCol = oTable.DBColumns.DBColumns(i)
            If oCol.ColName = oDBCol.ColName Then
               Set oCol.ADOXColumnSource = oDBCol.ADOXColumnSource
               Exit Sub
            End If
         Next i
      End If
   Next oTable
End With

End Sub

Friend Sub ColumnAddCol(ByVal oDBTable As cDBTable, ByVal oDBCol As cDBColumn)
'------------------------------------------------------------------------------
'Purpose  : Adds a column to the table's column collection
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 20.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable

With Me
   For Each oTable In .DBTables.DBTables
      If oTable.TblName = oDBTable.TblName Then
         oTable.DBColumns.DBColumns.Add oDBCol
         Exit For
      End If
   Next oTable
End With

End Sub

Friend Sub FireEvent(ByVal eEventType As eBAEventType, ByVal lEvent As Long, vntValue As Variant)
'------------------------------------------------------------------------------
'Purpose  : Allow other objects to raise events
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 08.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

Select Case eEventType

Case etOnXMLParse
   RaiseEvent OnXMLParse(lEvent, vntValue)
Case etOnDataChange
   RaiseEvent OnDataChange(lEvent, vntValue)
Case etOnDBEvent
   RaiseEvent OnDBEvent(lEvent, vntValue)
End Select

End Sub

Friend Function ParseXML(ByVal sXMLFile As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Root XML parsing method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 07.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim sXPath As String, sParseErr As String

Dim oXMLNode As MSXML2.IXMLDOMNode, oXMLChildNode As MSXML2.IXMLDOMNode, oXMLDoc As MSXML2.DOMDocument
Dim oXMLErr As MSXML2.IXMLDOMParseError
'Dim oXMLNodeList As MSXML2.IXMLDOMNodeList

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDB:ParseXML->"

Set oXMLDoc = New MSXML2.DOMDocument60
oXMLDoc.async = False
Call oXMLDoc.Load(sXMLFile)

If oXMLDoc.parseError.errorCode <> 0 Then
   Set oXMLErr = oXMLDoc.parseError
   With oXMLErr
      sParseErr = "XML Parse Error" & vbNewLine & _
         "errCode: " & .errorCode & vbNewLine & _
         "filePos: " & .filepos & vbNewLine & _
         "line   : " & .Line & vbNewLine & _
         "linePos: " & .linepos & vbNewLine & _
         "reason : " & .reason & vbNewLine & _
         "srcText: " & .srcText & vbNewLine
   End With
   RaiseEvent OnXMLParse(xmlErrOtherUnown, sParseErr)
   Set oXMLErr = Nothing
   Set oXMLDoc = Nothing
   Exit Function
End If

' Parse ConnectionStrings first
sXPath = "./Settings"

Set oXMLNode = oXMLDoc.selectSingleNode(sXPath)

If oXMLNode Is Nothing Then
' Document contains no data -> exit

   RaiseEvent OnXMLParse(xmlNoContent, vbNullString)
   Set oXMLErr = Nothing
   Set oXMLDoc = Nothing
   Exit Function

Else     '// If oXMLNode Is Nothing
' Parse the XML into our objects

   ' Parse attributes first ...
'   If Not ParseXMLAttr(oXMLNode) = True Then
'      XMLErrCount = XMLErrCount + 1
'   End If
   
   ' ... and (sub) nodes afterwards
   ' Settings contains as ChildNodes
   ' - DBSettings
   ' - Tables
   ' - Mappings
   If oXMLNode.hasChildNodes Then
   
      ' *** Node DBSettings ***
      ' Source
      sXPath = "./DBSettings/DBSetting[@Name='Source']"
   
      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)
      
      If oXMLChildNode Is Nothing Then
      ' Document contains no data -> exit
      
         RaiseEvent OnXMLParse(xmlMissingConnectionSource, vbNullString)
         Set oXMLErr = Nothing
         Set oXMLDoc = Nothing
         Exit Function
      
      Else
      
         With Me
            .ConnectionSource = modXMLUtil.GetXMLNodeValueFromNode(oXMLChildNode, "Connection", vbNullString)
            If Len(.ConnectionSource) < 1 Then
               RaiseEvent OnXMLParse(xmlMissingConnectionSource, vbNullString)
               Set oXMLErr = Nothing
               Set oXMLDoc = Nothing
               Exit Function
            Else
               RaiseEvent OnDataChange(dcConnectionSource, .ConnectionSource)
            End If
         End With
      
      End If
      
      ' Target
      sXPath = "./DBSettings/DBSetting[@Name='Target']"
   
      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)
      
      If oXMLChildNode Is Nothing Then
      ' Document contains no data -> exit
      
         RaiseEvent OnXMLParse(xmlMissingConnectionTarget, vbNullString)
         Set oXMLErr = Nothing
         Set oXMLDoc = Nothing
         Exit Function
      
      Else
      
         With Me
            .ConnectionTarget = modXMLUtil.GetXMLNodeValueFromNode(oXMLChildNode, "Connection", vbNullString)
            If Len(.ConnectionTarget) < 1 Then
               RaiseEvent OnXMLParse(xmlMissingConnectionTarget, vbNullString)
               Set oXMLErr = Nothing
               Set oXMLDoc = Nothing
               Exit Function
            Else
               RaiseEvent OnDataChange(dcConnectionTarget, .ConnectionTarget)
            End If
         End With
      
      End If
      
      
      ' *** Node Tables ***
      sXPath = "./Tables"
   
      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)
      
      If oXMLChildNode Is Nothing Then
      ' Document contains no data -> exit
      
         RaiseEvent OnXMLParse(xmlMissingTablesNode, vbNullString)
         Set oXMLErr = Nothing
         Set oXMLDoc = Nothing
         Exit Function
      
      Else
      
         If Me.DBTables.ParseXML(oXMLChildNode) > 0 Then
      
            If Me.DBTables.HasDBTables = False Then
               RaiseEvent OnXMLParse(xmlMissingTablesNode, vbNullString)
            End If
            Set oXMLErr = Nothing
            Set oXMLDoc = Nothing
            Exit Function
         
         End If
      
      End If

      ' *** Node Mappings ***
      sXPath = "./Mappings"
   
      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)
      
      If Not oXMLChildNode Is Nothing Then
      ' Mappings are optional ...
      
         If Me.DBMappings.ParseXML(oXMLChildNode) > 0 Then
         ' ... but if a Mappings node is present, at least one Mapping node should bepresent
      
            RaiseEvent OnXMLParse(xmlMissingMappingNode, vbNullString)
            Set oXMLErr = Nothing
            Set oXMLDoc = Nothing
            Exit Function
         
         End If
      
      End If
      
   Else     '// If oXMLNode.hasChildNodes
   
   ' Document contains no data -> exit
   
      RaiseEvent OnXMLParse(xmlNoContent, vbNullString)
      Set oXMLErr = Nothing
      Set oXMLDoc = Nothing
      Exit Function
   
   End If   '// If oXMLNode.hasChildNodes
   
End If   '// If oXMLNode Is Nothing

ParseXML = True

'------------------
ParseXMLErrExit:

On Error GoTo 0

Exit Function

'------------------
ParseXMLErrHandler:

RaiseEvent OnXMLParse(xmlErrOtherUnown, PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description)

Err.Clear
Resume ParseXMLErrExit

End Function
'==============================================================================

Private Function ParseXMLAttr(ByVal oNode As IXMLDOMNode) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Auslesen der Attribute dieses Nodes
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

On Error GoTo ParseXMLAttrErrHandler

Const PROCEDURE_NAME As String = "cDBTable:ParseXMLAttr->"

With Me

'   .ColName = GetXMLAttributeValueFromNode(oNode, "Name")
'   .IsIndex = CBool(GetXMLAttributeValueFromNode(oNode, "IsIndex", "False"))
'   .Precision = CLng(GetXMLAttributeValueFromNode(oNode, "Precision", "0"))
'   .Size = CLng(GetXMLAttributeValueFromNode(oNode, "Size", "0"))
'   .TypeSource = CLng(GetXMLAttributeValueFromNode(oNode, "TypeSource", "0"))
'   .TypeTarget = CLng(GetXMLAttributeValueFromNode(oNode, "TypeTarget", "0"))

End With

ParseXMLAttr = True

'------------------
ParseXMLAttrErrExit:
On Error GoTo 0

Exit Function

'------------------
ParseXMLAttrErrHandler:

Err.Clear
Resume ParseXMLAttrErrExit

End Function
'==============================================================================

Private Sub Class_Initialize()

   With Me
      Set .DBMappings = New cDBMappings
      Set .DBTables = New cDBTables
   End With

End Sub

Private Sub Class_Terminate()

   With Me
      Set .DBMappings = Nothing
      Set .DBTables = Nothing
   
      ' *** Kill potentially open datase connections
      ' No messin' here ...
      On Error Resume Next
   
      If Not .CnTarget Is Nothing Then
         If Not .CnTarget.State = adStateClosed Then
            Call .CnTarget.Close
         End If
         Set .CnTarget = Nothing
      End If
   
      If Not .CnSource Is Nothing Then
         If Not .CnSource.State = adStateClosed Then
            Call .CnSource.Close
         End If
         Set .CnSource = Nothing
      End If
   
   End With
   
   On Error GoTo 0

End Sub
'==============================================================================

Private Sub moDBTables_AddColADOXCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)

Call ColumnAddADOXCol(oTable, oCol)

End Sub

Private Sub moDBTables_AddCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)

Call ColumnAddCol(oTable, oCol)

End Sub

Private Sub moDBTables_AddColSourceType(ByVal oTable As cDBTable, ByVal oCol As cDBColumn, ByVal eColumnType As eADOXDataTypeEnum)

Call ColumnAddSourceType(oTable, oCol, eColumnType)

End Sub

Private Sub moDBTables_OnXMLParse(ByVal eEvent As Long, ByVal vntValue As Variant)

RaiseEvent OnXMLParse(eEvent, vntValue)

End Sub

