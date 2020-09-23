VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Windows System Color Alterer"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.ColorSelector ColorSelector1 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change!"
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "A Window's Background"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "Highlighted Text Background"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "Windows Title Bar"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Written by Daniel..."
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Select your color here ->>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdChange_Click()
Dim RetVal
Dim SelectedColor

SelectedColor = QBColor(ColorSelector1.SelectedColor)

If Option1(0).Value = True Then
    RetVal = SetSysColors(1, COLOR_ACTIVECAPTION, SelectedColor)
ElseIf Option1(1).Value = True Then
    RetVal = SetSysColors(1, COLOR_HIGHLIGHT, SelectedColor)
ElseIf Option1(2).Value = True Then
    RetVal = SetSysColors(1, COLOR_WINDOW, SelectedColor)
End If
End Sub

Private Sub Form_Load()
MsgBox "Select what system colors you want to change and click the 'Change' button.", vbInformation, "Just a note..."

Option1(0).Value = True

End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
    Option1(1).Value = False
    Option1(2).Value = False
ElseIf Option1(1).Value = True Then
    Option1(0).Value = False
    Option1(2).Value = False
ElseIf Option1(2).Value = True Then
    Option1(0).Value = False
    Option1(1).Value = False
End If


End Sub
