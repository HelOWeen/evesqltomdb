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
' *** Win32 API
Public Const MAX_PATH As Long = 260
Public Const SXH_PROXY_SET_PROXY As Long = 2

' *** Application
' Misc.
Public Const PATH_APP As String = "EVESqlToMdb"
Public Const REG_SECTION As String = "Software\BasicAware\EVESqlToMdb"

' Application path / files
Public Const FILE_INIFILE As String = "EVESqlToMdb.ini"

' Command line parameters
Public Const CMD_CONFIG As String = "Config"             ' XML passed via CLI
Public Const CMD_AUTOSTART As String = "AutoStart"       ' Automatically start copying
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
' *** Win32 API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub InitCommonControls Lib "comctl32.dll" ()
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
Public gobjApp As cApplication   ' Application Object - application settings
Public gobjCmd As cCmdLine       ' Command line parameter processing
'==============================================================================

