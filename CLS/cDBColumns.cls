VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Database columns mapping
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
Private mbolCopyAllColumns As Boolean

Private mcolDBColumns As Collection
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
'==============================================================================

Friend Property Get DBColumns() As Collection
   Set DBColumns = mcolDBColumns
End Property

Friend Property Set DBColumns(ByVal colValue As Collection)
   Set mcolDBColumns = colValue
End Property

Friend Property Get DBColumnsCount() As Long
   DBColumnsCount = DBColumns.Count
End Property

Friend Property Get HasDBColumns() As Boolean
   HasDBColumns = CBool(DBColumnsCount > 0)
End Property

Friend Property Get CopyAllColumns() As Boolean
   CopyAllColumns = mbolCopyAllColumns
End Property

Friend Property Let CopyAllColumns(ByVal bolValue As Boolean)
   mbolCopyAllColumns = bolValue
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

Dim oColumn As cDBColumn

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDBColumns:ParseXML->"

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
   ' Columns enthält als ChildNodes
   ' - Column
   If oXMLNode.hasChildNodes Then

      ' *** Column ***
      sXPath = "./Column"
      
      Set oXMLNodeList = oXMLNode.selectNodes(sXPath)

      If Not oXMLNodeList Is Nothing Then
         
         For i = 0 To oXMLNodeList.Length - 1
         
            Set oXMLChildNode = oXMLNodeList.Item(i)
            
            Set oColumn = New cDBColumn
            lError = lError + oColumn.ParseXML(oXMLChildNode)
            DBColumns.Add oColumn
         
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

Const PROCEDURE_NAME As String = "cDBColumns:ParseXMLAttr->"

With Me

   .CopyAllColumns = CBool(GetXMLAttributeValueFromNode(oNode, "CopyAllColumns", "false"))

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
      Set .DBColumns = New Collection
   End With
   
End Sub

Private Sub Class_Terminate()
   
   With Me
      Set .DBColumns = Nothing
   End With
   
End Sub
'==============================================================================
