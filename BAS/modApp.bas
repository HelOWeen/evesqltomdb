Attribute VB_Name = "modApp"
'------------------------------------------------------------------------------
'Purpose  : Global util functions and Sub Main()
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
'$PROBHIDE OPTIMIZATION
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

Public Sub Main()
'------------------------------------------------------------------------------
'Purpose  : Startprozedur der Anwendung
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Dim eInit As eInitError
Dim sFile As String


' InitCommonControlsVB

'** Auswerten der Kommandozeile
Set gobjCmd = New cCmdLine
gobjCmd.Init Command$

Set gobjApp = New cApplication

sFile = Trim$(gobjCmd.GetValueByName(LCase$(CMD_CONFIG)))

If Len(sFile) > 0 Then
   
   If FileExist(sFile) = False Then
      
      MsgBox "Configuration file not found: " & sFile & ". Application will now terminate.", vbCritical Or vbOKOnly, "Initialisation failed"

      Set gobjApp = Nothing
      Exit Sub
      
   Else
   
      gobjApp.LastXML = sFile
      
   End If
   
End If


'** Erste Initialisierung der Anwendungsparameter
eInit = gobjApp.Init()
   
' ** Anwendungsfenster laden
Screen.MousePointer = vbHourglass
DoEvents
Load frmMain

PCtrl.SetWindowMinBox frmMain, False

Screen.MousePointer = vbNormal

With frmMain
   Set .moApp = gobjApp
   .txtXML.Text = gobjApp.LastXML
   Call .cmdApply_Click
   .Show
End With

End Sub
'==============================================================================

Public Sub CleanUp()
'------------------------------------------------------------------------------
'Funktion : Aufräumen bei Programmende
'
'Vorauss. : -
'Parameter: -
'Notiz    : -
'
'    Autor: Knuth Konrad 29.01.2002
' geändert: -
'------------------------------------------------------------------------------
On Error Resume Next

Const PROCEDURE_NAME As String = "modApp:CleanUp->"

Set gobjCmd = Nothing
Set gobjApp = Nothing

End Sub
'==============================================================================

Public Sub StatusMsg(ByVal sMsg As String)
'------------------------------------------------------------------------------
'Funktion : Gibt eine Meldung in einer Statusbar aus
'
'Aufgerufen von: -
'Vorauss. : -
'Parameter: sMsg           -  Anzuzeigender Text
'           stbStatusbar   -  Statusbar, falls abweichend von frmMain.stbMain
'           bolBeep        -  Nach Anzeige einen Beep ausgeben
'Notiz    : -
'
'    Autor: Knuth Konrad
' erstellt: 25.05.2000
' geändert: 07.06.2000
'           - Neue Parameter stbStatusBar und bolBeep
'------------------------------------------------------------------------------

With frmMain
   .stbMain.SimpleText = sMsg
   .stbMain.Refresh
   
   .txtTask.Text = sMsg
   .txtTask.Refresh

   .timMain.Enabled = True
End With

End Sub
'==============================================================================

Public Function EnQuote(ByVal sText As String, Optional ByVal sQuote As String = """") As String

EnQuote = sQuote & sText & sQuote

End Function
'==============================================================================

Public Sub DebugWrite2File(ByVal sContent As String, ByVal sFilename As String)

Dim hFile As Long

On Error GoTo DebugWrite2FileErrHandler

If gobjCmd.GetValueByName("debug") = True Then

   hFile = FreeFile
   
   Open sFilename For Output As #hFile
   Print #hFile, sContent
   Close #hFile

End If

'------------------
DebugWrite2FileErrExit:
On Error GoTo 0
Exit Sub

'------------------
DebugWrite2FileErrHandler:
Err.Clear
Resume DebugWrite2FileErrExit

End Sub
'==============================================================================

Public Function GetFieldType(ByVal oCol As ADOX.Column, _
   Optional ByVal bolAddTypeValue As Boolean = True) As String

Dim sTemp As String, eType As ADOX.DataTypeEnum
Dim bolIsNumeric As Boolean
Dim oProp As ADOX.Property

'sTemp = "Column type: "

Select Case oCol.Type
Case adBigInt
'20 Indicates an eight-byte signed integer (DBTYPE_I8).
   sTemp = sTemp & "adBigInt"
   bolIsNumeric = True
Case adBinary
' 128 Indicates a binary value (DBTYPE_BYTES).
   sTemp = sTemp & "adBinary"
Case adBoolean
' 11 Indicates a boolean value (DBTYPE_BOOL).
   sTemp = sTemp & "adBoolean"
Case adBSTR
' 8 Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
   sTemp = sTemp & "adBSTR"
Case adChapter
' 136 Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
   sTemp = sTemp & "adChapter"
Case adChar
' 129 Indicates a string value (DBTYPE_STR).
   sTemp = sTemp & "adChar"
Case adCurrency
' 6 Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
   sTemp = sTemp & "adCurrency"
   bolIsNumeric = True
Case adDate
'7 Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
   sTemp = sTemp & "adDate"
Case adDBDate
'133 Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
   sTemp = sTemp & "adDBDate"
Case adDBTime
'134 Indicates a time value (hhmmss) (DBTYPE_DBTIME).
   sTemp = sTemp & "adDBTime"
Case adDBTimeStamp
' 135 Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
   sTemp = sTemp & "adDBTimeStamp"
Case adDecimal
' 14 Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
   sTemp = sTemp & "adDecimal"
   bolIsNumeric = True
Case adDouble
' 5 Indicates a double-precision floating-point value (DBTYPE_R8).
   sTemp = sTemp & "adDouble"
   bolIsNumeric = True
Case adEmpty
' 0 Specifies no value (DBTYPE_EMPTY).
   sTemp = sTemp & "adEmpty"
Case adError
' 10 Indicates a 32-bit error code (DBTYPE_ERROR).
   sTemp = sTemp & "adError"
Case adFileTime
' 64 Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
   sTemp = sTemp & "adFileTime"
Case adGUID
' 72 Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
   sTemp = sTemp & "adGUID"
Case adIDispatch
' 9 Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH).
   'Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = sTemp & "adIDispatch"
Case adInteger
'3 Indicates a four-byte signed integer (DBTYPE_I4).
   sTemp = sTemp & "adInteger"
   bolIsNumeric = True
Case adIUnknown
' 13 Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN).
' Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = sTemp & "adIUnknown"
Case adLongVarBinary
' 205 Indicates a long binary value.
   sTemp = sTemp & "adLongVarBinary"
Case adLongVarChar
' 201 Indicates a long string value.
   sTemp = sTemp & "adLongVarChar"
Case adLongVarWChar
' 203 Indicates a long null-terminated Unicode string value.
   sTemp = sTemp & "adLongVarWChar"
Case adNumeric
' 131 Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
   sTemp = sTemp & "adNumeric"
   bolIsNumeric = True
Case adPropVariant
' 138 Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
   sTemp = sTemp & "adPropVariant"
Case adSingle
' 4 Indicates a single-precision floating-point value (DBTYPE_R4).
   sTemp = sTemp & "adSingle"
   bolIsNumeric = True
Case adSmallInt
' 2 Indicates a two-byte signed integer (DBTYPE_I2).
   sTemp = sTemp & "adSmallInt"
   bolIsNumeric = True
Case adTinyInt
' 16 Indicates a one-byte signed integer (DBTYPE_I1).
   sTemp = sTemp & "adTinyInt"
   bolIsNumeric = True
Case adUnsignedBigInt
'21 Indicates an eight-byte unsigned integer (DBTYPE_UI8).
   sTemp = sTemp & "adUnsignedBigInt"
   bolIsNumeric = True
Case adUnsignedInt
' 19 Indicates a four-byte unsigned integer (DBTYPE_UI4).
   sTemp = sTemp & "adUnsignedInt"
   bolIsNumeric = True
Case adUnsignedSmallInt
' 18 Indicates a two-byte unsigned integer (DBTYPE_UI2).
   sTemp = sTemp & "adUnsignedSmallInt"
   bolIsNumeric = True
Case adUnsignedTinyInt
' 17 Indicates a one-byte unsigned integer (DBTYPE_UI1).
   sTemp = sTemp & "adUnsignedTinyInt"
   bolIsNumeric = True
Case adUserDefined
' 132 Indicates a user-defined variable (DBTYPE_UDT).
   sTemp = sTemp & "adUserDefined"
Case adVarBinary
' 204 Indicates a binary value.
   sTemp = sTemp & "adVarBinary"
Case adVarChar
' 200 Indicates a string value.
   sTemp = sTemp & "adVarChar"
Case adVariant
' 12 Indicates an Automation Variant (DBTYPE_VARIANT).
' Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = sTemp & "adVariant"
 Case adVarNumeric
 ' 139 Indicates a numeric value.
   sTemp = sTemp & "adVarNumeric"
   bolIsNumeric = True
Case adVarWChar
' 202 Indicates a null-terminated Unicode character string.
   sTemp = sTemp & "adVarWChar"
Case adWChar
' 130 Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
   sTemp = sTemp & "adWChar"
Case Else
   sTemp = sTemp & "Other/Unknwon"
End Select

If bolAddTypeValue = True Then
   sTemp = sTemp & "(" & CStr(oCol.Type) & ")"
End If

sTemp = sTemp & ", Size: " & Format$(oCol.DefinedSize, "#,###,##0")

If bolIsNumeric Then
   sTemp = sTemp & ", Precision: " & CStr(oCol.Precision)
End If

GetFieldType = sTemp

End Function
'==============================================================================

Public Function GetFieldTypeEx(ByVal oCol As ADOX.Column, _
   ByRef eType As ADOX.DataTypeEnum, ByRef lSize As Long, ByRef lPrecision As Long) As String

Dim sTemp As String
' , eType As ADOX.DataTypeEnum
Dim oProp As ADOX.Property

eType = oCol.Type
lSize = oCol.DefinedSize
lPrecision = oCol.Precision

Select Case oCol.Type
Case adBigInt
'20 Indicates an eight-byte signed integer (DBTYPE_I8).
   sTemp = "adBigInt"
Case adBinary
' 128 Indicates a binary value (DBTYPE_BYTES).
   sTemp = "adBinary"
Case adBoolean
' 11 Indicates a boolean value (DBTYPE_BOOL).
   sTemp = "adBoolean"
Case adBSTR
' 8 Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
   sTemp = "adBSTR"
Case adChapter
' 136 Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
   sTemp = "adChapter"
Case adChar
' 129 Indicates a string value (DBTYPE_STR).
   sTemp = "adChar"
Case adCurrency
' 6 Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
   sTemp = "adCurrency"
Case adDate
'7 Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
   sTemp = "adDate"
Case adDBDate
'133 Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
   sTemp = "adDBDate"
Case adDBTime
'134 Indicates a time value (hhmmss) (DBTYPE_DBTIME).
   sTemp = "adDBTime"
Case adDBTimeStamp
' 135 Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
   sTemp = "adDBTimeStamp"
Case adDecimal
' 14 Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
   sTemp = "adDecimal"
Case adDouble
' 5 Indicates a double-precision floating-point value (DBTYPE_R8).
   sTemp = "adDouble"
Case adEmpty
' 0 Specifies no value (DBTYPE_EMPTY).
   sTemp = "adEmpty"
Case adError
' 10 Indicates a 32-bit error code (DBTYPE_ERROR).
   sTemp = "adError"
Case adFileTime
' 64 Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
   sTemp = "adFileTime"
Case adGUID
' 72 Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
   sTemp = "adGUID"
Case adIDispatch
' 9 Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH).
   'Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = "adIDispatch"
Case adInteger
'3 Indicates a four-byte signed integer (DBTYPE_I4).
   sTemp = "adInteger"
Case adIUnknown
' 13 Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN).
' Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = "adIUnknown"
Case adLongVarBinary
' 205 Indicates a long binary value.
   sTemp = "adLongVarBinary"
Case adLongVarChar
' 201 Indicates a long string value.
   sTemp = "adLongVarChar"
Case adLongVarWChar
' 203 Indicates a long null-terminated Unicode string value.
   sTemp = "adLongVarWChar"
Case adNumeric
' 131 Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
   sTemp = "adNumeric"
Case adPropVariant
' 138 Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
   sTemp = "adPropVariant"
Case adSingle
' 4 Indicates a single-precision floating-point value (DBTYPE_R4).
   sTemp = "adSingle"
Case adSmallInt
' 2 Indicates a two-byte signed integer (DBTYPE_I2).
   sTemp = "adSmallInt"
Case adTinyInt
' 16 Indicates a one-byte signed integer (DBTYPE_I1).
   sTemp = "adTinyInt"
Case adUnsignedBigInt
'21 Indicates an eight-byte unsigned integer (DBTYPE_UI8).
   sTemp = "adUnsignedBigInt"
Case adUnsignedInt
' 19 Indicates a four-byte unsigned integer (DBTYPE_UI4).
   sTemp = "adUnsignedInt"
Case adUnsignedSmallInt
' 18 Indicates a two-byte unsigned integer (DBTYPE_UI2).
   sTemp = "adUnsignedSmallInt"
Case adUnsignedTinyInt
' 17 Indicates a one-byte unsigned integer (DBTYPE_UI1).
   sTemp = "adUnsignedTinyInt"
Case adUserDefined
' 132 Indicates a user-defined variable (DBTYPE_UDT).
   sTemp = "adUserDefined"
Case adVarBinary
' 204 Indicates a binary value.
   sTemp = "adVarBinary"
Case adVarChar
' 200 Indicates a string value.
   sTemp = "adVarChar"
Case adVariant
' 12 Indicates an Automation Variant (DBTYPE_VARIANT).
' Note   This data type is currently not supported by ADO. Usage may cause unpredictable results.
   sTemp = "adVariant"
 Case adVarNumeric
 ' 139 Indicates a numeric value.
   sTemp = "adVarNumeric"
Case adVarWChar
' 202 Indicates a null-terminated Unicode character string.
   sTemp = "adVarWChar"
Case adWChar
' 130 Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
   sTemp = "adWChar"
Case Else
   sTemp = "Other/Unknwon"
End Select

GetFieldTypeEx = sTemp

End Function
'==============================================================================



