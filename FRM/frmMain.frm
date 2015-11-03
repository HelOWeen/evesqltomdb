VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "EVESqlToMdb"
   ClientHeight    =   8535
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   16275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer timMain 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   8280
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      Caption         =   "XML mapping file"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15975
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   315
         Left            =   14760
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdXMLBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   14040
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtXML 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   13575
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Data export / import"
      Height          =   5175
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   15975
      Begin MSComctlLib.ListView lvwTableSource 
         Height          =   3015
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5318
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Start data transfer"
         Height          =   375
         Left            =   14040
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar prgMain 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4680
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtTask 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   13575
      End
      Begin MSComctlLib.ListView lvwColumnsSource 
         Height          =   3015
         Left            =   4560
         TabIndex        =   18
         Top             =   1320
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5318
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwTableTarget 
         Height          =   3015
         Left            =   11640
         TabIndex        =   20
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5318
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblRecords 
         AutoSize        =   -1  'True
         Caption         =   "lblrecords"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   4440
         Width           =   675
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "Contained columns"
         Height          =   195
         Index           =   5
         Left            =   4560
         TabIndex        =   17
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "Copied tables"
         Height          =   195
         Index           =   4
         Left            =   11640
         TabIndex        =   19
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "Outstanding tables"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "Current / last task"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Database settings"
      Height          =   1815
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   15975
      Begin VB.CommandButton cmdTestTarget 
         Caption         =   "Test"
         Height          =   375
         Left            =   14520
         TabIndex        =   10
         ToolTipText     =   "Test target ADO connection string"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdTestSource 
         Caption         =   "Test"
         Height          =   375
         Left            =   14520
         TabIndex        =   7
         ToolTipText     =   "Test source ADO connection string"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtConnectionTarget 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   14055
      End
      Begin VB.TextBox txtConnectionSource 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   14055
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "MS Access database - Connection string"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2880
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "SQL Server - Connection string"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2190
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Public Enum eAppState
   asIdle
   asIsRunning
   asStopRequest
End Enum

Private meAppState As eAppState

Public LastRecCount As Long

Public WithEvents moApp As cApplication
Attribute moApp.VB_VarHelpID = -1
Public WithEvents moDB As cDB
Attribute moDB.VB_VarHelpID = -1
'==============================================================================

Public Property Get AppState() As eAppState
   AppState = meAppState
End Property

Friend Property Let AppState(ByVal eValue As eAppState)
   meAppState = eValue
   ' Debug.Print "AppState: " & CStr(eValue)
End Property
'==============================================================================

Private Sub DoColMapping(ByVal sTable As String)

Dim oMap As cDBMapping, oCol As cDBColumn, oTable As cDBTable

Set oTable = moDB.DBTables.DBTableGetByName(sTable)

If Not oTable Is Nothing Then

   For Each oCol In oTable.DBColumns.DBColumns
   
      If oCol.TypeTarget = adoAsSource Then
      ' See if this is handled by a general mapping
      
         If moDB.DBMappings.HasDBMappingOfType(oCol.GetTypeSource) Then
            Set oMap = moDB.DBMappings.GetDBMappingForType(oCol.GetTypeSource)
            oCol.TypeTarget = oMap.GetTypeTarget
            'oCol.AllowNull = oMap.AllowNull
            If oMap.Precision <> -1 Then
               oCol.Precision = oMap.Precision
            End If
            If oMap.Size <> -1 Then
               oCol.Size = oMap.Size
            End If
         End If
         
      End If
   
   Next oCol

End If

End Sub
'==============================================================================

Public Sub cmdApply_Click()

gobjApp.LastXML = Me.txtXML.Text

End Sub

Public Sub cmdOK_Click()

Dim dblTimeStart As Double, dblTimeEnd As Double

If AppState = asIsRunning Then
   Exit Sub
End If

Screen.MousePointer = vbHourglass

If gobjApp.AutoStart = False Then
   ' Reinitiate the config.
   Me.cmdApply_Click
Else
' Don't reinitialize with param AutoStart, but set AutoStart to False
' to prevent error loops
   gobjApp.AutoStart = False
End If

AppState = asIsRunning
prgMain.Value = 0

dblTimeStart = Timer

If MainTransferStart(Me, moDB) = True Then
' Initial checks went well, start with actual data transfer

   Call MainCopyData(Me, moDB)
   
End If

Screen.MousePointer = vbNormal

dblTimeEnd = Timer

If AppState = asStopRequest Then
   StatusMsg "Operation cancelled by user."
Else
   StatusMsg "Done. Duration: " & Format$(dblTimeEnd - dblTimeStart, "#,##0.00") & " seconds."
End If

AppState = asIdle

End Sub

Private Sub cmdTestSource_Click()

Dim sMsg As String, bolResult As Boolean

StatusMsg "Testing source database connection ..."
Screen.MousePointer = vbHourglass

bolResult = MainDBTestConnection(Me.txtConnectionSource.Text, sMsg)

Screen.MousePointer = vbNormal

If bolResult = True Then
   StatusMsg "Connection test successfully completed."
   MsgBox "Test successfully completed.", vbInformation Or vbOKOnly, "Connection test"
Else
   StatusMsg "Connection test failed."
   MsgBox sMsg, vbCritical Or vbOKOnly, "Connection test"
End If

End Sub

Private Sub cmdTestTarget_Click()

Dim sMsg As String, bolResult As Boolean

StatusMsg "Testing target database connection ..."
Screen.MousePointer = vbHourglass

bolResult = MainDBTestConnection(Me.txtConnectionTarget.Text, sMsg)

Screen.MousePointer = vbNormal

If bolResult = True Then
   StatusMsg "Connection test successfully completed."
   MsgBox "Test successfully completed.", vbInformation Or vbOKOnly, "Connection test"
Else
   StatusMsg "Connection test failed."
   MsgBox sMsg, vbCritical Or vbOKOnly, "Connection test"
End If


End Sub

Private Sub cmdXMLBrowse_Click()

Dim oCmDlg As New saCommonDialog
Dim sFile As String

sFile = Me.txtXML.Text

With oCmDlg
   If .VBGetOpenFileName(sFile, "XML mapping file", True, False, False, False, "XML files (*.xml)|*.xml", , gobjApp.AppPath, "Select mapping file") = True Then
      Me.txtXML.Text = sFile
      gobjApp.LastXML = sFile
   End If
End With

End Sub
'==============================================================================

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

' Bei laufender Anwendung Stop signalisieren
If KeyCode = vbKeyEscape Then
   If AppState = asIsRunning Then
      AppState = asStopRequest
      DoEvents
   End If
End If

'If (KeyCode = vbKeyF4 And Shift = 2) Or (KeyCode = vbKeyW And Shift = 2) Or KeyCode = vbKeyEscape Then
''Formular bei CTRL+F4 oder CTRL+W oder ESC schließen
'   Me.Hide
'   Unload Me
'End If

End Sub

Private Sub Form_Load()

Dim oWP As cWindowPosition

lblRecords.Caption = vbNullString

Set oWP = New cWindowPosition
With oWP
   .RegSection = gobjApp.RegSectionWindowPosition
   .RestorePosition Me
   '.SavePosition Me
End With

Call MainSetup(Me)

AppState = asIdle

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim oWP As cWindowPosition
'Call gobjApp.SetMainSize(Me)
Set oWP = New cWindowPosition
With oWP
   .RegSection = gobjApp.RegSectionWindowPosition
   '.RestorePosition Me
   .SavePosition Me
End With

If AppState = asIsRunning Then
   AppState = asStopRequest
   Me.Show
   Cancel = True
ElseIf AppState = asIdle Then
   CleanUp
End If

End Sub

Private Sub mnuFileExit_Click()

Me.Hide
Unload Me

End Sub

Private Sub mnuHelpAbout_Click()

Load frmAbout
frmAbout.Show vbModal, Me

End Sub

Private Sub timMain_Timer()

With frmMain.stbMain
   .SimpleText = vbNullString
   .Refresh
End With

timMain.Enabled = False

End Sub
'==============================================================================

Private Sub moApp_OnXMLFileChange(ByVal sXMLFile As String, ByVal bolExists As Boolean)

If bolExists = True Then
   ' ToDo: XML laden, Textboxen mit ConnectionStrings füllen
   Set moDB = New cDB
   With moDB
      Call .ParseXML(sXMLFile)
   End With
End If

End Sub
'==============================================================================

Private Sub moDB_OnDataChange(ByVal eEvent As eBADataChange, ByVal vntValue As Variant)

Select Case eEvent

Case dcConnectionSource
   Me.txtConnectionSource.Text = CStr(vntValue)

Case dcConnectionTarget
   Me.txtConnectionTarget.Text = CStr(vntValue)

End Select

End Sub

Private Sub moDB_OnXMLParse(ByVal eEvent As eBAParseEvent, ByVal vntValue As Variant)

Dim sMsg As String

Select Case eEvent

Case eBAParseEvent.xmlErrOtherUnown
   sMsg = "Unknown error"
Case eBAParseEvent.xmlSuccess

Case eBAParseEvent.xmlNoContent
   sMsg = "XML has no content"
Case eBAParseEvent.xmlMissingConnectionSource
   sMsg = "Missing source connection string"
Case eBAParseEvent.xmlMissingConnectionTarget
   sMsg = "Missing target connection string"
Case eBAParseEvent.xmlMissingTablesNode
   sMsg = "Missing Tables node"
Case eBAParseEvent.xmlTableDefAlreadyExists
   sMsg = "Duplicate table definition:"
End Select

If (eEvent <> xmlSuccess) Then
   MsgBox sMsg & vbNewLine & CStr(vntValue), vbCritical Or vbOKOnly, "Error parsing XML"
End If

End Sub

Private Sub moDB_OnDBEvent(ByVal eEvent As eBADbEvent, ByVal vntValue As Variant)

Dim sMsg As String
Dim oLI As ListItem, lValue As Long

Select Case eEvent

Case dbeErrOtherUnown
   sMsg = "Other/Unknown error"
Case dbeSuccess

Case dbeConnectionSourceFailed
   sMsg = "Connection to source database failed"
Case dbeConnectionTargetFailed
   sMsg = "Connection to target database faile"
Case dbeTblSourceMissing
   sMsg = "Source table missing"

Case dbeColDoMapping
   DoColMapping CStr(vntValue)

Case dbeColSourceMissing
   sMsg = "Source column missing"
Case dbeTableCreateFailed
   sMsg = "Target table creation failed"
Case dbeIndexCreateFailed
   sMsg = "Target index creation failed"

Case dbeRecordCount
   lblRecords.Caption = "Number of records in source table: " & Format$(CLng(vntValue), "#,###,###,##0")
   LastRecCount = CLng(vntValue)
   
   If Not lvwTableSource.SelectedItem Is Nothing Then
      Set oLI = lvwTableSource.SelectedItem
      Call oLI.ListSubItems.Add(, , Format$(CLng(vntValue), "#,###,###,##0"))
   End If

Case dbeRecordAdded
   Dim dblPercent As Double, oSubLI As ListSubItem
   lValue = CLng(vntValue)
   dblPercent = Percent(lValue, LastRecCount)
   With prgMain
      If dblPercent > 100 Then
         dblPercent = 100
      End If
      .Value = dblPercent
   End With

   If Not lvwTableTarget.SelectedItem Is Nothing Then
      Set oLI = lvwTableTarget.SelectedItem
      If oLI.ListSubItems.Count < 1 Then
         oLI.ListSubItems.Add
      End If
      Set oSubLI = oLI.ListSubItems(1)
      oSubLI.Text = Format$(lValue, "#,###,###,##0")
   End If

End Select

sMsg = sMsg & vbNewLine

Select Case eEvent
Case eBADbEvent.dbeSuccess, eBADbEvent.dbeRecordAdded, eBADbEvent.dbeRecordCount, eBADbEvent.dbeColDoMapping
Case Else
   Dim frmMsg As frmExtTextbox
   Set frmMsg = New frmExtTextbox
   Load frmMsg
   Set frmMsg.ParentTextbox = Me.txtTask
   frmMsg.Caption = "An error occured"
   txtTask.Text = sMsg & CStr(vntValue)
   Screen.MousePointer = vbNormal
   CenterFrmOnFrm frmMsg, Me
   frmMsg.Show vbModal, Me
   Set frmMsg = Nothing
   'MsgBox sMsg & CStr(vntValue), vbCritical Or vbOKOnly, "An error occured"
End Select

End Sub
'==============================================================================

