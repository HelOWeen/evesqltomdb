VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBMappings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Data type mapping
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 16.10.2015
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
Private mcolDBMappings As Collection
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
'==============================================================================

Friend Property Get DBMappings() As Collection
   Set DBMappings = mcolDBMappings
End Property

Friend Property Set DBMappings(ByVal colValue As Collection)
   Set mcolDBMappings = colValue
End Property

Friend Property Get DBMappingsCount() As Long
   DBMappingsCount = DBMappings.Count
End Property

Friend Property Get HasDBMappings() As Boolean
   HasDBMappings = CBool(DBMappingsCount > 0)
End Property
'==============================================================================

Friend Function HasDBMappingOfType(ByVal eType As eADOXDataTypeEnum) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Find out if a general data type mapping for a certain (source) data
'           type exists
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 16.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oMap As cDBMapping

On Error GoTo HasDBMappingOfTypeErrHandler

Const PROCEDURE_NAME As String = "cDBMappings:HasDBMappingOfType->"
' gobjLog.AppTrace PROCEDURE_NAME,
' TraceIt PROCEDURE_NAME,

With Me

   If .HasDBMappings = False Then
      HasDBMappingOfType = False
      Exit Function
   End If
   
   For Each oMap In .DBMappings
      If oMap.TypeSource = eType Then
         HasDBMappingOfType = True
         Exit Function
      End If
   Next oMap
   
End With

'------------------
HasDBMappingOfTypeErrExit:

On Error GoTo 0
Exit Function

'------------------
HasDBMappingOfTypeErrHandler:
' gobjLog.AppLog ApplicationError, PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
' LogIt PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
Err.Clear
Resume HasDBMappingOfTypeErrExit

End Function

Friend Function GetDBMappingForType(ByVal eType As eADOXDataTypeEnum) As cDBMapping
'------------------------------------------------------------------------------
'Purpose  : Find out if a general data type mapping for a certain (source) data
'           type exists
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 16.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim oMap As cDBMapping

On Error GoTo GetDBMappingForTypeErrHandler

Const PROCEDURE_NAME As String = "cDBMappings:GetDBMappingForType->"
' gobjLog.AppTrace PROCEDURE_NAME,
' TraceIt PROCEDURE_NAME,

With Me

   If .HasDBMappings = False Then
      Set GetDBMappingForType = Nothing
      Exit Function
   End If
   
   For Each oMap In .DBMappings
      If oMap.TypeSource = eType Then
         Set GetDBMappingForType = oMap
         Exit Function
      End If
   Next oMap
   
End With

'------------------
GetDBMappingForTypeErrExit:

On Error GoTo 0
Exit Function

'------------------
GetDBMappingForTypeErrHandler:
' gobjLog.AppLog ApplicationError, PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
' LogIt PROCEDURE_NAME & CStr(Err.Number) & ", " & CStr(Err.Source) & ", " & Err.Description
Err.Clear
Resume GetDBMappingForTypeErrExit

End Function

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

Dim oMap As cDBMapping

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDBMappings:ParseXML->"

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
   ' Mappings enthält als ChildNodes
   ' - Mapping
   If oXMLNode.hasChildNodes Then

      ' *** Mapping ***
      sXPath = "./Mapping"
      
      Set oXMLNodeList = oXMLNode.selectNodes(sXPath)

      If Not oXMLNodeList Is Nothing Then
         
         For i = 0 To oXMLNodeList.Length - 1
         
            Set oXMLChildNode = oXMLNodeList.Item(i)
            
            Set oMap = New cDBMapping
            lError = lError + oMap.ParseXML(oXMLChildNode)
            DBMappings.Add oMap
         
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

Const PROCEDURE_NAME As String = "cDBMappings:ParseXMLAttr->"

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
   
   With Me
      Set .DBMappings = New Collection
   End With
   
End Sub

Private Sub Class_Terminate()
   
   With Me
      Set .DBMappings = Nothing
   End With
   
End Sub
'==============================================================================

