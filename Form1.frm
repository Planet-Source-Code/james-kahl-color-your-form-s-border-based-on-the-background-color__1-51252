VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Bakground"
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo ColorErr
    
    'intialize the common dialog
    With cdlColor
        .CancelError = True
        .ShowColor
        Me.BackColor = .Color
        DrawBorder Me, Me.BackColor
    End With
    Exit Sub
    
ColorErr:
    If Err.Number <> 32755 Then
        MsgBox Err.Number & ": " & Err.Description, , "Error choosing color"
    End If
End Sub

Private Sub Form_Paint()
    'don't draw the border when we first show the form
    If Me.BackColor = vbButtonFace Then Exit Sub
    DrawBorder Me, Me.BackColor
End Sub
