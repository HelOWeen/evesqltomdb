VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'Purpose  : Represents a database column mapping
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
' Mirrored ADO data types enum
Public Enum eADOXDataTypeEnum
   adoAsSource = -1  ' Use same data type as source in target
   adoBigInt = 20
   adoBinary = 128
   adoBoolean = 11
   adoBSTR = 8
   adoChapter = 136
   adoChar = 129
   adoCurrency = 6
   adoDate = 7
   adoDBDate = 133
   adoDBTime = 134
   adoDBTimeStamp = 135
   adoDecimal = 14
   adoDouble = 5
   adoEmpty = 0
   adoError = 10
   adoFileTime = 64
   adoGUID = 72
   adoIDispatch = 9
   adoInteger = 3
   adoIUnknown = 13
   adoLongVarBinary = 205
   adoLongVarChar = 201
   adoLongVarWChar = 203
   adoNumeric = 131
   adoPropVariant = 138
   adoSingle = 4
   adoSmallInt = 2
   adoTinyInt = 16
   adoUnsignedBigInt = 21
   adoUnsignedInt = 19
   adoUnsignedSmallInt = 18
   adoUnsignedTinyInt = 17
   adoUserDefined = 132
   adoVarBinary = 204
   adoVarChar = 200
   adoVariant = 12
   adoVarNumeric = 139
   adoVarWChar = 202
   adoWChar = 130
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Private msColName As String
Private meTypeTarget As eADOXDataTypeEnum
Private meTypeSource As eADOXDataTypeEnum
Private mbolIsIndex As Boolean
Private mbolIsPrimary As Boolean
Private mbolIsUnique As Boolean
Private mlSize As Long
Private mlPrecision As Long
Private msAllowNullStr As String

Public ADOXColumnSource As ADOX.Column
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
'==============================================================================

Friend Property Get ColName() As String
   ColName = msColName
End Property

Friend Property Let ColName(ByVal sValue As String)
   msColName = sValue
End Property

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

Friend Property Get IsIndex() As Boolean
   IsIndex = mbolIsIndex
End Property

Friend Property Let IsIndex(ByVal bolValue As Boolean)
   mbolIsIndex = bolValue
End Property

Friend Property Get IsPrimary() As Boolean
   IsPrimary = mbolIsPrimary
End Property

Friend Property Let IsPrimary(ByVal bolValue As Boolean)
   mbolIsPrimary = bolValue
End Property

Friend Property Get IsUnique() As Boolean
   IsUnique = mbolIsUnique
End Property

Friend Property Let IsUnique(ByVal bolValue As Boolean)
   mbolIsUnique = bolValue
End Property

Friend Property Get Size() As Long
   Size = mlSize
End Property

Friend Property Let Size(ByVal lValue As Long)
   mlSize = lValue
End Property

Friend Property Get Precision() As Long
   Precision = mlPrecision
End Property

Friend Property Let Precision(ByVal lValue As Long)
   mlPrecision = lValue
End Property

Friend Property Get AllowNullStr() As String
   AllowNullStr = msAllowNullStr
End Property

Friend Property Let AllowNullStr(ByVal sValue As String)
   msAllowNullStr = sValue
End Property

Friend Property Get AllowNull() As Boolean

   ' Assume False ...
   AllowNull = False
   
   If Len(AllowNullStr) > 0 Then
   ' Defined in XML
      AllowNull = CBool(AllowNullStr)
   Else
   ' Not set in XML, cehck out the source column's setting
      If Not Me.ADOXColumnSource Is Nothing Then
         With Me.ADOXColumnSource
            If .Attributes = (.Attributes Or adColNullable) Then
               AllowNull = True
            End If
         End With
      End If
   End If

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
         GetTypeSource = IIf(.TypeSource <> adoError, .TypeSource, .ADOXColumnSource.Type)
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
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim lError As Long, i As Long
Dim sXPath As String
Dim oXMLNode As MSXML2.IXMLDOMNode, oXMLChildNode As MSXML2.IXMLDOMNode
Dim oXMLNodeList As MSXML2.IXMLDOMNodeList

On Error GoTo ParseXMLErrHandler

DoEvents: Sleep 0
Const PROCEDURE_NAME As String = "cDBColumn:ParseXML->"

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

Const PROCEDURE_NAME As String = "cDBColumn:ParseXMLAttr->"

With Me

   .AllowNullStr = GetXMLAttributeValueFromNode(oNode, "AllowNull", vbNullString)
   .ColName = GetXMLAttributeValueFromNode(oNode, "Name")
   .IsIndex = CBool(GetXMLAttributeValueFromNode(oNode, "IsIndex", "False"))
   .IsPrimary = CBool(GetXMLAttributeValueFromNode(oNode, "IsPrimary", "False"))
   .IsUnique = CBool(GetXMLAttributeValueFromNode(oNode, "IsUnique", "False"))
   .Precision = CLng(Val(GetXMLAttributeValueFromNode(oNode, "Precision", "0")))
   .Size = CLng(Val(GetXMLAttributeValueFromNode(oNode, "Size", "0")))
   sTemp = GetXMLAttributeValueFromNode(oNode, "TypeTarget", CStr(eADOXDataTypeEnum.adoAsSource))
   If Len(sTemp) < 1 Then
      .TypeTarget = eADOXDataTypeEnum.adoAsSource
   Else
      .TypeTarget = CLng(Val(sTemp))
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
