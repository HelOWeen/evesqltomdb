VERSION 5.00
Begin VB.Form frmExtTextbox 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtContent 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmExtTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Const FRM_MIN_HEIGHT As Long = 3600
Const FRM_MIN_WIDTH As Long = 4800

Dim WithEvents moTxtBox As TextBox
Attribute moTxtBox.VB_VarHelpID = -1
'==============================================================================

Public Property Get ParentTextbox() As TextBox

   Set ParentTextbox = moTxtBox
   
End Property

Public Property Set ParentTextbox(ByVal oValue As TextBox)

   Set moTxtBox = oValue
   Me.txtContent.Text = ParentTextbox.Text
   
End Property
'==============================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not ParentTextbox Is Nothing Then
   ParentTextbox.Text = Me.txtContent.Text
End If

End Sub
'==============================================================================

Private Sub Form_Resize()

If Me.WindowState <> vbMinimized Then

   With Me
      If .Height < FRM_MIN_HEIGHT Then
         .Height = FRM_MIN_HEIGHT
      End If
      If .Width < FRM_MIN_WIDTH Then
         .Width = FRM_MIN_WIDTH
      End If
   End With
   
   With Me.txtContent
      .Move Me.ScaleLeft + 120, Me.ScaleTop + 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
   End With

End If

End Sub
'==============================================================================

Private Sub moTxtBox_Change()

Me.txtContent.Text = ParentTextbox.Text

End Sub
'==============================================================================
