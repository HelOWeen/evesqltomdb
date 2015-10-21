VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Represents a general data type mapping
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
Private meTypeTarget As eADOXDataTypeEnum
Private meTypeSource As eADOXDataTypeEnum
Private mlSize As Long
Private mlPrecision As Long
Private mbolAllowNull As Boolean

Public ADOXColumnSource As ADOX.Column
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
'==============================================================================

Friend Property Get TypeSource() As eADOXDataTypeEnum
   TypeSource = meTypeSource
End Property

Friend Property Let TypeSource(ByVal eValue As eADOXDataTypeEnum)
   meTypeSource = eValue
End Property

Friend Property Get TypeTarget() As eADOXDataTypeEnum
   TypeTarget = meTypeTarget
End Property

Friend Property Let TypeTarget(ByVal eValue As eADOXDataTypeEnum)
   meTypeTarget = eValue
End Property

Friend Property Get AllowNull() As Boolean
   AllowNull = mbolAllowNull
End Property

Friend Property Let AllowNull(ByVal bolValue As Boolean)
   mbolAllowNull = bolValue
End Property

Friend Property Get Precision() As Long
   Precision = mlPrecision
End Property

Friend Property Let Precision(ByVal lValue As Long)
   mlPrecision = lValue
End Property

Friend Property Get Size() As Long
   Size = mlSize
End Property

Friend Property Let Size(ByVal lValue As Long)
   mlSize = lValue
End Property

Friend Property Get GetTypeTarget() As eADOXDataTypeEnum

   With Me
      If .TypeTarget = adoAsSource Then
         GetTypeTarget = GetTypeSource()
      Else
         GetTypeTarget = IIf(.TypeTarget <> adoError, .TypeTarget, .GetTypeSource)
      End If
   End With
   
End Property

Friend Property Get GetTypeSource() As eADOXDataTypeEnum

   With Me
      If .TypeSource = adoAsSource Then
         GetTypeSource = .ADOXColumnSource.Type
      Else
         If Not .ADOXColumnSource Is Nothing Then
            GetTypeSource = IIf(.TypeSource <> adoError, .TypeSource, .ADOXColumnSource.Type)
         Else
            GetTypeSource = .TypeSource
         End If
      End If
   End With
   
End Property

Friend Function ParseXML(ByVal oXMLNodeParent As MSXML2.IXMLDOMNode) As Long
'------------------------------------------------------------------------------
'Purpose  : Parses the configuration XML into its objects
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim lError As Long, i As Long
Dim sXPath As String
Dim oXMLNode As MSXML2.IXMLDOMNode, oXMLChildNode As MSXML2.IXMLDOMNode
Dim oXMLNodeList As MSXML2.IXMLDOMNodeList

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDBMapping:ParseXML->"

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
'   If oXMLNode.hasChildNodes Then
'
'      ' *** Columns ***
'      sXPath = "./Columns"
'
'      Set oXMLChildNode = oXMLNode.selectSingleNode(sXPath)
'
'      If Not oXMLChildNode Is Nothing Then
'
'         Set oColumns = New cDBColumns
'         lError = lError + oColumns.ParseXML(oXMLChildNode)
'
'      End If
'
'   End If   ' //If oXMLNode.hasChildNodes
         
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
Dim sTemp As String

On Error GoTo ParseXMLAttrErrHandler

Const PROCEDURE_NAME As String = "cDBMapping:ParseXMLAttr->"

With Me

   .AllowNull = CBool(GetXMLAttributeValueFromNode(oNode, "AllowNull", "False"))
   .Size = CLng(Val(GetXMLAttributeValueFromNode(oNode, "Size", "-1")))
   .Precision = CLng(Val(GetXMLAttributeValueFromNode(oNode, "Precision", "-1")))
   sTemp = GetXMLAttributeValueFromNode(oNode, "TypeTarget", CStr(eADOXDataTypeEnum.adoAsSource))
   If Len(sTemp) < 1 Then
      .TypeTarget = eADOXDataTypeEnum.adoAsSource
   Else
      .TypeTarget = CLng(Val(sTemp))
   End If
   sTemp = GetXMLAttributeValueFromNode(oNode, "TypeSource", CStr(eADOXDataTypeEnum.adoAsSource))
   If Len(sTemp) < 1 Then
      .TypeSource = eADOXDataTypeEnum.adoAsSource
   Else
      .TypeSource = CLng(Val(sTemp))
   End If

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

   ' Set defaults
   With Me
      .TypeTarget = eADOXDataTypeEnum.adoAsSource
      .TypeSource = eADOXDataTypeEnum.adoAsSource
   End With
   
End Sub

