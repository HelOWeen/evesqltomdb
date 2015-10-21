Attribute VB_Name = "modGlobal"
' $PROBHIDE VB.NET

'------------------------------------------------------------------------------
'Purpose  : Projektweite Variabeln- und Prozedurdeklarationen
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
'*** Win32 API Konstanten
Public Const MAX_PATH As Long = 260
Public Const SXH_PROXY_SET_PROXY As Long = 2

' Diverse
Public Const PATH_APP As String = "EVESqlToMdb"
Public Const REG_SECTION As String = "Software\BasicAware\EVESqlToMdb"

' Anwendungspfade/-dateien
Public Const FILE_INIFILE As String = "EVESqlToMdb.ini"

' Kommandozeilenparameter
Public Const CMD_CONFIG As String = "Config"             ' INI per CLI übergeben
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'** Win32 API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub InitCommonControls Lib "comctl32.dll" ()
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Public gobjApp As cApplication   ' Application Object - Applikationseinstellungen
Public gobjCmd As cCmdLine       ' Kommandozeilenparameter
'==============================================================================

