VERSION 5.00
Begin VB.Form FormTest 
   Caption         =   "Color Palette Test"
   ClientHeight    =   2970
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4410
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   2430
      TabIndex        =   2
      Top             =   2265
      Width           =   1470
   End
   Begin VB.TextBox TextTest 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      Left            =   510
      TabIndex        =   0
      Text            =   "Color Palette Test"
      Top             =   270
      Width           =   3390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Test"
      Height          =   390
      Left            =   2430
      TabIndex        =   1
      Top             =   1590
      Width           =   1470
   End
End
Attribute VB_Name = "FormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

' Just add frmColorPalette and ModuleColorPalette to your project.
' The parameter passed to the Palette is the name of the control where you want apply
' the ForeColor or BackColor.
' It can be any control that has "ForeColor" and "BackColor" properties

ShowPalette TextTest

End Sub
Private Sub Command3_Click()

    Unload Me

End Sub
Private Sub Form_Unload(Cancel As Integer)

' unloads the ColorPalette if loaded.
    If frmColorPalette Is Nothing = False Then Unload frmColorPalette
    End

End Sub
