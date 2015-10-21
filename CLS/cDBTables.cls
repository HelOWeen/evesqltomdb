VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Database tables
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
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
Private Enum eBAParseEvent
   xmlErrOtherUnown = -1
   xmlSuccess = 0
   xmlNoContent
   xmlMissingConnectionSource
   xmlMissingConnectionTarget
   xmlMissingTablesNode
   xmlMissingMappingNode
   xmlTableDefAlreadyExists
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Private mcolDBTables As Collection
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
Public Event OnXMLParse(ByVal eEvent As Long, ByVal vntValue As Variant)
Public Event AddColSourceType(ByVal oTable As cDBTable, ByVal oCol As cDBColumn, _
   ByVal eColumnType As eADOXDataTypeEnum)
Public Event AddColADOXCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)
Public Event AddCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)
'==============================================================================

Friend Property Get DBTables() As Collection
   Set DBTables = mcolDBTables
End Property

Friend Property Set DBTables(ByVal colValue As Collection)
   Set mcolDBTables = colValue
End Property

Friend Property Get DBTablesCount() As Long
   DBTablesCount = DBTables.Count
End Property

Friend Property Get HasDBTables() As Boolean
   HasDBTables = CBool(DBTables.Count > 0)
End Property
'==============================================================================

Friend Function DBTableExistsByName(ByVal sTable As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Find out if the table sTable exists in the DBTables collection
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 09.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable

For Each oTable In DBTables
   If LCase$(oTable.TblName) = LCase$(sTable) Then
      DBTableExistsByName = True
      Exit Function
   End If
Next oTable

End Function

Friend Function DBTableGetByName(ByVal sTable As String) As cDBTable
'------------------------------------------------------------------------------
'Purpose  : Retrieve a table object from the collection by the table's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 09.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oTable As cDBTable

For Each oTable In DBTables
   If LCase$(oTable.TblName) = LCase$(sTable) Then
      Set DBTableGetByName = oTable
      Exit Function
   End If
Next oTable

End Function

Friend Sub DBTableAddColSourceType(ByVal oTable As cDBTable, ByVal oCol As cDBColumn, ByVal eType As eADOXDataTypeEnum)

RaiseEvent AddColSourceType(oTable, oCol, eType)

End Sub

Friend Sub DBTableAddADOXCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)

RaiseEvent AddColADOXCol(oTable, oCol)

End Sub

Friend Sub DBTableAddCol(ByVal oTable As cDBTable, ByVal oCol As cDBColumn)

RaiseEvent AddCol(oTable, oCol)

End Sub

Friend Function ParseXML(ByVal oXMLNodeParent As MSXML2.IXMLDOMNode) As Long
'------------------------------------------------------------------------------
'Purpose  : Parses the configuration XML into its objects
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
Dim lError As Long, i As Long
Dim sXPath As String
Dim oXMLNode As MSXML2.IXMLDOMNode, oXMLChildNode As MSXML2.IXMLDOMNode
Dim oXMLNodeList As MSXML2.IXMLDOMNodeList

Dim oTable As cDBTable

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDBTables:ParseXML->"

Set oXMLNode = oXMLNodeParent

If oXMLNode Is Nothing Then
' Dokument enthält keine Daten -> raus hier

   Exit Function

Else
' Parsen der Daten in die einzelnen Objekt

   ' Erst die Attributwerte dieser Node auslesen ...
   If Not ParseXMLAttr(oXMLNode) = True Then
      lError = lError + 1
   End If
   
   ' ... dann alle anderen Nodes in Objekte wandeln.
   ' Tables enthält als ChildNodes
   ' - Table
   If oXMLNode.hasChildNodes Then

      ' *** Table ***
      sXPath = "./Table"
      
      Set oXMLNodeList = oXMLNode.selectNodes(sXPath)

      If Not oXMLNodeList Is Nothing Then
         
         For i = 0 To oXMLNodeList.Length - 1
         
            Set oXMLChildNode = oXMLNodeList.Item(i)
            
            Set oTable = New cDBTable
            lError = lError + oTable.ParseXML(oXMLChildNode)
            
            If Not Me.DBTableExistsByName(oTable.TblName) Then
               DBTables.Add oTable
            Else
               RaiseEvent OnXMLParse(eBAParseEvent.xmlTableDefAlreadyExists, oTable.TblName)
               lError = lError + 1
               Exit For
            End If
         
         Next i

      End If

   End If   ' //If oXMLNode.hasChildNodes
         
End If

'------------------
ParseXMLErrExit:

ParseXML = lError
On Error GoTo 0

Exit Function

'------------------
ParseXMLErrHandler:

RaiseEvent OnXMLParse(eBAParseEvent.xmlErrOtherUnown, PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description)
Err.Clear

lError = lError + 1
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

Const PROCEDURE_NAME As String = "cDBTables:ParseXMLAttr->"

With Me

'   .PointOfTurnaround = GetXMLAttributeValueFromNode(oNode, "PointOfTurnaround")
'   TemTextLog "PointOfTurnaround: " & .PointOfTurnaround
'
'   .IssuedBy = GetXMLAttributeValueFromNode(oNode, "IssuedBy")
'   TemTextLog "IssuedBy: " & .IssuedBy
'
'   .SpecificFareType = GetXMLAttributeValueFromNode(oNode, "SpecificFareType")
'   TemTextLog "SpecificFareType: " & .SpecificFareType

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
   Set DBTables = New Collection
End Sub

Private Sub Class_Terminate()
   Set DBTables = Nothing
End Sub

