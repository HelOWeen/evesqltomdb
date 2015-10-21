VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Represents a database table
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
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Private msTblName As String
Private msQuery As String

Private moDBColumns As cDBColumns
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
'==============================================================================

Friend Property Get TblName() As String
   TblName = msTblName
End Property

Friend Property Let TblName(ByVal sValue As String)
   msTblName = sValue
End Property

Friend Property Get Query() As String
   Query = msQuery
End Property

Friend Property Let Query(ByVal sValue As String)
   msQuery = sValue
End Property

Friend Property Get DBColumns() As cDBColumns
   Set DBColumns = moDBColumns
End Property

Friend Property Set DBColumns(ByVal oValue As cDBColumns)
   Set moDBColumns = oValue
End Property

Friend Property Get DBColumnsCount() As Long
   DBColumnsCount = DBColumns.DBColumnsCount
End Property

Friend Property Get HasDBColumns() As Boolean
   HasDBColumns = DBColumns.HasDBColumns
End Property
'==============================================================================

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

Dim oColumns As cDBColumns

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
   ' - Columns
   If oXMLNode.hasChildNodes Then

      ' *** Columns ***
      ' //Table[@Name='invCategories']/Columns
      sXPath = "//Table[@Name='" & Me.TblName & "']/Columns"
      
      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)

      If Not oXMLChildNode Is Nothing Then
         
         lError = lError + DBColumns.ParseXML(oXMLChildNode)

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

Const PROCEDURE_NAME As String = "cDBTable:ParseXMLAttr->"

With Me

   .Query = GetXMLAttributeValueFromNode(oNode, "WhereClause", vbNullString)
   .TblName = GetXMLAttributeValueFromNode(oNode, "Name")

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
      Set .DBColumns = New cDBColumns
   End With
   
End Sub

Private Sub Class_Terminate()
   
   With Me
      Set .DBColumns = Nothing
   End With
   
End Sub
'==============================================================================
