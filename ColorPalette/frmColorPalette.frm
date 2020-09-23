VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorPalette 
   BorderStyle     =   0  'None
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   Icon            =   "frmColorPalette.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmColorPalette.frx":000C
   ScaleHeight     =   5325
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox PictClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1545
      Picture         =   "frmColorPalette.frx":0E24
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   58
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox PictMaxMin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      Picture         =   "frmColorPalette.frx":13A8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   57
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton ComApplyColor 
      Appearance      =   0  'Flat
      Caption         =   "&Apply"
      Height          =   330
      Left            =   165
      TabIndex        =   56
      Top             =   4260
      Width           =   1605
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "&Background"
      Height          =   225
      Index           =   1
      Left            =   165
      TabIndex        =   55
      Top             =   3555
      Width           =   1590
   End
   Begin VB.OptionButton OptionColor 
      Alignment       =   1  'Right Justify
      Caption         =   "&Foreground"
      Height          =   225
      Index           =   0
      Left            =   165
      TabIndex        =   54
      Top             =   3300
      Value           =   -1  'True
      Width           =   1590
   End
   Begin VB.TextBox TxtGreen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      TabIndex        =   51
      Text            =   "0"
      Top             =   2415
      Width           =   360
   End
   Begin VB.PictureBox PicGreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   150
      Left            =   150
      ScaleHeight     =   150
      ScaleWidth      =   1230
      TabIndex        =   50
      Top             =   2475
      Width           =   1230
      Begin VB.PictureBox LabBlackGreen 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   495
         ScaleHeight     =   255
         ScaleWidth      =   1185
         TabIndex        =   64
         Top             =   -45
         Width           =   1215
      End
      Begin VB.PictureBox LabFillGreen 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         ScaleHeight     =   255
         ScaleWidth      =   195
         TabIndex        =   63
         Top             =   -30
         Width           =   225
      End
   End
   Begin VB.PictureBox PicBlue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   150
      Left            =   150
      ScaleHeight     =   150
      ScaleWidth      =   1230
      TabIndex        =   52
      Top             =   2760
      Width           =   1230
      Begin VB.PictureBox LabFillBlue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         ScaleHeight     =   255
         ScaleWidth      =   195
         TabIndex        =   62
         Top             =   -30
         Width           =   225
      End
      Begin VB.PictureBox LabBlackBlue 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   510
         ScaleHeight     =   255
         ScaleWidth      =   1185
         TabIndex        =   61
         Top             =   -30
         Width           =   1215
      End
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   6
      Left            =   1365
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   7
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   7
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   10
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   10
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   11
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   12
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   13
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   14
      Left            =   1365
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   15
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   15
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   16
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   16
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   17
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   17
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   18
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   18
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   19
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   19
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   20
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   20
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   21
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   21
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   22
      Left            =   1365
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   22
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   23
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   23
      Top             =   1155
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   24
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   24
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   25
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   26
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   27
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   27
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   28
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   28
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   29
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   29
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   30
      Left            =   1365
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   30
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   31
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   31
      Top             =   1380
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   32
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   32
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   33
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   33
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   34
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   34
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   35
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   35
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   36
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   36
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   37
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   37
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   38
      Left            =   1365
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   38
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   39
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   39
      Top             =   1605
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   40
      Left            =   165
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   40
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   41
      Left            =   360
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   41
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   42
      Left            =   570
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   42
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   43
      Left            =   765
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   43
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   44
      Left            =   975
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   44
      Top             =   1845
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   45
      Left            =   1170
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   45
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   46
      Left            =   1380
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   46
      Top             =   1830
      Width           =   180
   End
   Begin VB.PictureBox MiniCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   47
      Left            =   1575
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   47
      Top             =   1845
      Width           =   180
   End
   Begin VB.TextBox TxtBlue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      TabIndex        =   53
      Text            =   "0"
      Top             =   2715
      Width           =   360
   End
   Begin VB.TextBox TxtRed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      TabIndex        =   49
      Text            =   "255"
      Top             =   2115
      Width           =   360
   End
   Begin VB.PictureBox PicRed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   150
      Left            =   150
      ScaleHeight     =   150
      ScaleWidth      =   1230
      TabIndex        =   48
      Top             =   2175
      Width           =   1230
      Begin VB.PictureBox LabBlackRed 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   525
         ScaleHeight     =   255
         ScaleWidth      =   1185
         TabIndex        =   60
         Top             =   -30
         Width           =   1215
      End
      Begin VB.PictureBox LabFillRed 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         ScaleHeight     =   255
         ScaleWidth      =   195
         TabIndex        =   59
         Top             =   -30
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList Tools 
      Left            =   780
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":192C
            Key             =   "MaxiN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":1EC0
            Key             =   "MaxiD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":2454
            Key             =   "MaxiO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":29E8
            Key             =   "MiniN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":2F7C
            Key             =   "MiniD"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":3510
            Key             =   "MiniO"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":3AA4
            Key             =   "CloseN"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":4038
            Key             =   "CloseD"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":45CC
            Key             =   "CloseO"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":4B60
            Key             =   "FrmMax"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColorPalette.frx":85F4
            Key             =   "FrmMin"
         EndProperty
      EndProperty
   End
   Begin VB.Label LabColorPalette 
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   67
      Top             =   165
      Width           =   975
   End
   Begin VB.Label LabColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   165
      TabIndex        =   66
      Top             =   3885
      Width           =   1590
   End
   Begin VB.Label LabApplyColorIn 
      BackStyle       =   0  'Transparent
      Caption         =   "Apply color in:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   165
      TabIndex        =   65
      Top             =   3060
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   8
      Left            =   1185
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   7
      Left            =   900
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   6
      Left            =   615
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   5
      Left            =   330
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   4
      Left            =   750
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   3
      Left            =   465
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   2
      Left            =   1035
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   1
      Left            =   1320
      Top             =   2085
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   0
      Left            =   180
      Top             =   2085
      Width           =   15
   End
   Begin VB.Label LabColorPalette 
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   68
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Franco Gomes - September 2002.
Private SourceControl As Object
Private CurColor As Long
Private Blue As Double
Private Green As Double
Private Red As Double
Private Sep As Long
Private MiW As Long
Private MiV As Long
Private MxV As Long
Private MxW As Long
Private CurW As Long

'The following declarations are related with the displacement of the Palette form.
Private MouseDn As Long
Private MouseDnX As Long
Private MouseDnY As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Function GetScrollValue(PosX As Single, MinW As Long, MaxW As Long, MinV As Long, MaxV As Long) As Long

' getting the value (between 0 and 255) relatively to the cursor position on the scroll.
    CurW = MinW + PosX - 1
    If CurW < MinW Then CurW = MinW
    If CurW > MaxW + MinW Then CurW = MaxW + MinW
    GetScrollValue = CInt((MaxV * (CurW - MinW)) / MaxW)
    If GetScrollValue < MinV Then GetScrollValue = MinV
    If GetScrollValue > MaxV Then GetScrollValue = MaxV

End Function
Private Sub SetScrollWidth(Obj1 As PictureBox, Obj2 As PictureBox, CurV As Long, Desl As Long, MinW As Long, MaxW As Long, MaxV As Long)

' setting the scroll position (adjusting the widths of the 2 Picture Boxes)
    CurW = CInt((MaxW * CurV) / MaxV) + MinW
    If CurW < MinW Then CurW = MinW
    If CurW > MaxW + MinW Then CurW = MaxW + MinW
    Obj1.Width = CurW
    Obj2.Left = CurW - MinW + Desl

End Sub
Private Sub CalculateRGB(CurColor As Long)
' getting the color separation in RGB
    
    Blue = Int(CurColor / 65536)
    Green = Int((CurColor - (Blue * 65536)) / 256)
    Red = CurColor - ((Blue * 65536) + (Green * 256))
    
        
End Sub
Private Sub ApplyColor()
' The color is aplied as the choice made in OptionColor()
    If OptionColorValue = True Then
        SourceControl.BackColor = LabColor.BackColor
        Else:
        SourceControl.ForeColor = LabColor.BackColor
    End If

End Sub
Private Sub ComApplyColor_Click()

    ApplyColor

End Sub
Private Sub LabColorPalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' transfering mouse status to Form
    Form_MouseDown Button, Shift, X, Y

End Sub
Private Sub LabColorPalette_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub
Private Sub LabColorPalette_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, X, Y

End Sub
Private Sub Form_Load()

    Me.Width = 2000
    Me.Height = 4725
    MouseDn = 0
    Me.Picture = Tools.ListImages("FrmMax").Picture ' Set the background picture
    PictClose.Picture = Tools.ListImages("CloseN").Picture ' Set the Close Button picture
    PictMaxMin.Picture = Tools.ListImages("MiniN").Picture ' Set the Maximize/minimize button picture

End Sub
Public Sub IniPalette(SrcCtrl As Object)

' This is the public procedure call before show the palette form
Dim Ca As Long
Dim Cb As Long
Dim Cc As Long
Dim Cd As Long
Dim Tp As Long
Dim Lf As Long
Dim Tq As Long

    OptionColor(1).Value = OptionColorValue ' Setting the Public variable that keeps the last user option.
    
    Set SourceControl = SrcCtrl
    If OptionColorValue = True Then
        CurColor = SourceControl.BackColor
        Else: CurColor = SourceControl.ForeColor
    End If
    
    If CurColor < 0 Then CurColor = 0
    If CurColor > &HFFFFFF Then CurColor = &HFFFFFF
    
    ' The following code serves to organize all the objects in form.
    Tp = 705
    Lf = 164
    Tq = Tp
    Cc = 0
    For Cd = 0 To 5
        Cb = Lf
        For Ca = 0 To 7
            MiniCor(Cc).Left = Cb
            MiniCor(Cc).Top = Tq
            Cc = Cc + 1
            Cb = Cb + 210
        Next Ca
        Tq = Tq + 210
    Next Cd
    Ca = MiniCor(7).Left + MiniCor(7).Width
    Cb = MiniCor(7).Left + MiniCor(7).Width - MiniCor(0).Left
    TxtRed.Left = Ca - TxtRed.Width
    TxtGreen.Left = Ca - TxtGreen.Width
    TxtBlue.Left = Ca - TxtBlue.Width
    CalculateRGB CurColor
    MiW = 60
    MiV = 0
    MxV = 255
    Sep = 60
    MxW = PicRed.Width - MiW - Sep / 4
    LabFillRed.Left = -MiW
    LabFillGreen.Left = -MiW
    LabFillBlue.Left = -MiW
    LabBlackRed.Width = PicRed.Width
    LabBlackGreen.Width = PicRed.Width
    LabBlackBlue.Width = PicRed.Width
    LabColor.Left = Lf
    LabColor.Width = (TxtRed.Left + TxtRed.Width) - Lf
    OptionColor(0).Left = Lf
    OptionColor(1).Left = Lf
    OptionColor(0).Width = LabColor.Width
    OptionColor(1).Width = LabColor.Width
    ComApplyColor.Left = Lf
    ComApplyColor.Width = LabColor.Width
    SetBars
    SetTexts
    LabColor.BackColor = CurColor

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseDn = 1
    MouseDnX = X
    MouseDnY = Y

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Z As POINTAPI

    If Button = 0 Then ' Placing the appropriate pictures for the top buttons (Close and Max/Min)
        If X > PictClose.Left - 45 And X < PictClose.Left + PictClose.Width + 45 And Y > PictClose.Top - 45 And Y < PictClose.Top + PictClose.Height + 45 Then
            PictClose.Picture = Tools.ListImages("CloseO").Picture
            Else
            PictClose.Picture = Tools.ListImages("CloseN").Picture
        End If
        If X > PictMaxMin.Left - 45 And X < PictMaxMin.Left + PictMaxMin.Width + 45 And Y > PictMaxMin.Top - 45 And Y < PictMaxMin.Top + PictMaxMin.Height + 45 Then
            If frmColorPalette.Height < 4725 Then
                PictMaxMin.Picture = Tools.ListImages("MaxiO").Picture
                Else
                PictMaxMin.Picture = Tools.ListImages("MiniO").Picture
            End If
            Else
            If frmColorPalette.Height < 4725 Then
                PictMaxMin.Picture = Tools.ListImages("MaxiN").Picture
                Else
                PictMaxMin.Picture = Tools.ListImages("MiniN").Picture
            End If
        End If
    End If
    
    If MouseDn <> 1 Then Exit Sub
    
    ' Moving the form
    Call GetCursorPos(Z)

    Me.Top = (Z.Y * 15) - MouseDnY
    Me.Left = (Z.X * 15) - MouseDnX

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseDn = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmColorPalette = Nothing

End Sub
Private Sub SetBars()

'Setting the 3 scroll bars
    SetScrollWidth frmColorPalette.LabFillRed, frmColorPalette.LabBlackRed, CInt(Red), Sep, MiW, MxW, MxV
    SetScrollWidth frmColorPalette.LabFillGreen, frmColorPalette.LabBlackGreen, CInt(Green), Sep, MiW, MxW, MxV
    SetScrollWidth frmColorPalette.LabFillBlue, frmColorPalette.LabBlackBlue, CInt(Blue), Sep, MiW, MxW, MxV

End Sub
Private Sub SetTexts()
'Setting the 3 TextBoxs
    TxtRed.Text = CStr(Red)
    TxtGreen.Text = CStr(Green)
    TxtBlue.Text = CStr(Blue)

End Sub
Private Sub SetColors()

    CurColor = RGB(Red, Green, Blue)
    LabColor.BackColor = CurColor

End Sub
Private Sub MiniCor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' One of the small colored squares is clicked
    CurColor = MiniCor(Index).BackColor
    CalculateRGB CurColor
    SetBars
    SetTexts
    SetColors

End Sub
Private Sub OptionColor_Click(Index As Integer)

OptionColorValue = OptionColor(1).Value

End Sub
Private Sub PicBlue_Click()

    TxtBlue.SelStart = 0
    TxtBlue.SelLength = Len(TxtBlue.Text) + 1

End Sub
Private Sub PicBlue_KeyDown(KeyCode As Integer, Shift As Integer)
' We can move the scroll bars with the navigation keys. The Shif Key multiply the displacement.
Dim Mov As Long
If Shift = 1 Then
    Mov = 10
    Else: Mov = 1
End If

    If KeyCode = 39 And Blue < MxV Then Blue = Blue + Mov
    If KeyCode = 37 And Blue > MiV Then Blue = Blue - Mov
    If KeyCode = 36 Then Blue = MiV
    If KeyCode = 35 Then Blue = MxV
    TxtBlue.Text = CStr(Blue)

End Sub
Private Sub PicBlue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicBlue.BackColor = 0

End Sub
Private Sub PicBlue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Blue = GetScrollValue(X, MiW, MxW, MiV, MxV)
        SetTexts
    End If

End Sub
Private Sub PicBlue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicBlue.BackColor = &HFFFF&

End Sub
Private Sub PicGreen_KeyDown(KeyCode As Integer, Shift As Integer)

Dim Mov As Long
If Shift = 1 Then
    Mov = 10
    Else: Mov = 1
End If

    If KeyCode = 39 And Green < MxV Then Green = Green + Mov
    If KeyCode = 37 And Green > MiV Then Green = Green - Mov
    If KeyCode = 36 Then Green = MiV
    If KeyCode = 35 Then Green = MxV
    TxtGreen.Text = CStr(Green)

End Sub
Private Sub PicGreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicGreen.BackColor = 0

End Sub
Private Sub PicGreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Green = GetScrollValue(X, MiW, MxW, MiV, MxV)
        SetTexts
    End If

End Sub
Private Sub PicGreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicGreen.BackColor = &HFFFF&

End Sub
Private Sub PicRed_KeyDown(KeyCode As Integer, Shift As Integer)

Dim Mov As Long
If Shift = 1 Then
    Mov = 10
    Else: Mov = 1
End If

    If KeyCode = 39 And Red < MxV Then Red = Red + Mov
    If KeyCode = 37 And Red > MiV Then Red = Red - Mov
    If KeyCode = 36 Then Red = MiV
    If KeyCode = 35 Then Red = MxV
    TxtRed.Text = CStr(Red)

End Sub
Private Sub PicRed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicRed.BackColor = 0

End Sub
Private Sub PicRed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Red = GetScrollValue(X, MiW, MxW, MiV, MxV)
        SetTexts
    End If

End Sub
Private Sub PicRed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicRed.BackColor = &HFFFF&

End Sub
Private Sub TxtBlue_Change()

' some input verificatios
    If IsNumeric(TxtBlue.Text) = False Then TxtBlue.Text = "0"
    If CInt(TxtBlue.Text) > 255 Then TxtBlue.Text = "255"
    If CInt(TxtBlue.Text) < 0 Then TxtBlue.Text = "0"
    Blue = CInt(TxtBlue.Text)
    SetBars
    SetColors
    LabColor.BackColor = CurColor

End Sub
Private Sub TxtBlue_GotFocus()

    TxtBlue.SelStart = 0
    TxtBlue.SelLength = Len(TxtBlue.Text) + 1

End Sub
Private Sub TxtGreen_Change()

    If IsNumeric(TxtGreen.Text) = False Then TxtGreen.Text = "0"
    If CInt(TxtGreen.Text) > 255 Then TxtGreen.Text = "255"
    If CInt(TxtGreen.Text) < 0 Then TxtGreen.Text = "0"
    Green = CInt(TxtGreen.Text)
    SetBars
    SetColors
    LabColor.BackColor = CurColor

End Sub
Private Sub TxtGreen_Click()

    TxtGreen.SelStart = 0
    TxtGreen.SelLength = Len(TxtGreen.Text) + 1

End Sub
Private Sub TxtGreen_GotFocus()

    TxtGreen.SelStart = 0
    TxtGreen.SelLength = Len(TxtGreen.Text) + 1

End Sub
Private Sub TxtRed_Change()

    If IsNumeric(TxtRed.Text) = False Then TxtRed.Text = "0"
    If CInt(TxtRed.Text) > 255 Then TxtRed.Text = "255"
    If CInt(TxtRed.Text) < 0 Then TxtRed.Text = "0"
    Red = CInt(TxtRed.Text)
    SetBars
    SetColors
    LabColor.BackColor = CurColor

End Sub
Private Sub TxtRed_Click()

    TxtRed.SelStart = 0
    TxtRed.SelLength = Len(TxtRed.Text) + 1

End Sub
Private Sub TxtRed_GotFocus()

    TxtRed.SelStart = 0
    TxtRed.SelLength = Len(TxtRed.Text) + 1

End Sub
Private Sub PictClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PictClose.Picture = Tools.ListImages("CloseD").Picture

End Sub
Private Sub Pictclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Unload Me

End Sub
Private Sub PictMaxMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If frmColorPalette.Height < 4725 Then
        PictMaxMin.Picture = Tools.ListImages("MaxiD").Picture
        Else
        PictMaxMin.Picture = Tools.ListImages("MiniD").Picture
    End If

End Sub
Private Sub PictMaxMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Placing the apropriate pictures (stored in the ImageList) to the Palette form and Max/Min top button.
    If frmColorPalette.Height < 4725 Then
        frmColorPalette.Height = 4725
        PictMaxMin.Picture = Tools.ListImages("MiniN").Picture
        frmColorPalette.Picture = Tools.ListImages("FrmMax").Picture
        Else
        frmColorPalette.Height = 555
        PictMaxMin.Picture = Tools.ListImages("MaxiN").Picture
        frmColorPalette.Picture = Tools.ListImages("FrmMin").Picture
    End If

End Sub
