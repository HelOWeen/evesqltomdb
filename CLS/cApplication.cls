VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'Purpose  : Application settings/configuration
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
' INI Sections
Private Const INISC_GENERAL As String = "General"
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Public Enum eInitError
   InitSuccess = 0
   DBOpenFailed
   DBFileNotFound
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Private msINIFile As String      ' INI inkl. komplettem Pfad
Private msAppPath As String      ' Path to <User>\AppData

Private msLastXML As String      ' XML config file
Private mbolAutoStart As Boolean ' Start data copying automatically
'------------------------------------------------------------------------------
'*** Events ***
'------------------------------------------------------------------------------
Public Event OnXMLFileChange(ByVal sXMLFile As String, ByVal bolExists As Boolean)
'==============================================================================

Friend Property Get AppPath() As String
   AppPath = NormalizePath(msAppPath)
End Property

Friend Property Let AppPath(ByVal sValue As String)
   msAppPath = sValue
End Property

Friend Property Get INIFile() As String
   INIFile = msINIFile
End Property

Friend Property Let INIFile(ByVal sValue As String)
   msINIFile = sValue
End Property

Friend Property Get LastXML() As String
   LastXML = msLastXML
End Property

Friend Property Let LastXML(ByVal sValue As String)
   
   msLastXML = sValue
   
   If FileExist(sValue) = True Then
      RaiseEvent OnXMLFileChange(sValue, True)
   Else
      RaiseEvent OnXMLFileChange(sValue, False)
   End If
   
End Property

Friend Property Get AutoStart() As Boolean
   AutoStart = mbolAutoStart
End Property

Friend Property Let AutoStart(ByVal bolValue As Boolean)
   mbolAutoStart = bolValue
End Property

Friend Property Get RegSectionWindowPosition() As String
   RegSectionWindowPosition = REG_SECTION & "\WindowPosition"
End Property

Friend Function Init(Optional ByVal sConfigXML As String = vbNullString) As eInitError
'------------------------------------------------------------------------------
'Purpose  : Initializierung des cApplication-Objekts
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: 05.11.2015
'           - Pass CLI config via parameter
'------------------------------------------------------------------------------
Dim lRetval As Long
Dim sTemp As String

' ** Last config XML file used
sTemp = Space$(MAX_PATH)
lRetval = GetPrivateProfileString(INISC_GENERAL, "CfgXML", LastXML, sTemp, Len(sTemp), INIFile)
If lRetval Then
   LastXML = Left$(sTemp, lRetval)
End If

' Config passed by command line?
If Len(sConfigXML) > 0 Then
   LastXML = sConfigXML
End If

Init = InitSuccess

End Function
'==============================================================================

Private Sub SaveSettings()
'------------------------------------------------------------------------------
'Purpose  : Programmeinstellungen speichern
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 06.10.2015
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

Call WritePrivateProfileString(INISC_GENERAL, "CfgXML", LastXML, INIFile)

End Sub
'==============================================================================

Private Sub Class_Initialize()

' *** Datenordner kreieren/zuweisen
CreateNestedDirs NormalizePath(PFile.GetSpecialFolderPath(spfAppdata)) & PATH_APP
AppPath = NormalizePath(PFile.GetSpecialFolderPath(spfAppdata)) & PATH_APP

' INI für programmweite Einstellungen
INIFile = NormalizePath(AppPath) & FILE_INIFILE

End Sub
'--------------------------

Private Sub Class_Terminate()

SaveSettings

End Sub
'==============================================================================
