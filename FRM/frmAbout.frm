VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info zu meiner Anwendung"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAbout 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   360
         Picture         =   "frmAbout.frx":0000
         ScaleHeight     =   337.12
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   337.12
         TabIndex        =   3
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version"
         Height          =   225
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         Caption         =   "Name der Anwendung"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1170
         TabIndex        =   5
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label lblDescription 
         Caption         =   "EVE Online Financial Tool"
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   3885
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2565
      Width           =   1260
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warnung: "
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   255
      TabIndex        =   1
      Top             =   2430
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z
'==============================================================================

Private Sub cmdOK_Click()

Unload Me

End Sub
'==============================================================================

Private Sub Form_Load()

Dim oWP As cWindowPosition

On Error Resume Next

Me.Caption = "Info about " & App.Title

lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

lblDescription = "EVESqlToMdb. Copy/Export CCP's SDE to another database." & vbNewLine & _
   "CCP hf is not affiliated with EWAViewer nor is it in any way responsible for the functionality (or errors) of this program."

lblTitle.Caption = App.Title

picIcon.Picture = frmMain.Icon
Me.Icon = frmMain.Icon

lblDisclaimer.Caption = "This application: Copyright " & Chr$(169) & " 2015, " & CStr(Year(Now)) & " by Hel O'Ween" & vbNewLine & _
   "EVE Online is a trademark of CCP Games. All EVE related material used in this application is copyrighted by CCP hf."

Set oWP = New cWindowPosition
With oWP
   .RegSection = gobjApp.RegSectionWindowPosition
   .RestorePosition Me
   '.SavePosition Me
End With

On Error GoTo 0

End Sub
'==============================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim oWP As cWindowPosition

Set oWP = New cWindowPosition
With oWP
   .RegSection = gobjApp.RegSectionWindowPosition
   .SavePosition Me
End With

End Sub
'==============================================================================

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 1) And (Shift = 2) Then
   MsgBox "Lead Developer: Knuth Konrad (Konrad@softAware.de)", vbOKOnly Or vbInformation, "Progammer's info"
End If

End Sub
'==============================================================================
