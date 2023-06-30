VERSION 5.00
Begin VB.Form DRS2 
   BackColor       =   &H00404040&
   Caption         =   "DRS2 - By D.C"
   ClientHeight    =   6660
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Roll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   4800
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Go 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label N6 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label N5 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label N4 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label N3 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label N2 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label N1 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image imgDie 
      Height          =   1575
      Left            =   4560
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label L6 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label L5 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label L4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label L3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label L2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label L1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dice Rolling Simulation 2"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu End 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "DRS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Go_Click()
Roll.Enabled = True
End Sub

Private Sub Roll_Timer()
Dim Die As Integer
Randomize
Die = Int(Rnd * 6) + 1
Select Case Die
Case 1
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\one.bmp")
L1.Caption = Val(L1.Caption) + 1
Case 2
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\two.bmp")
L2.Caption = Val(L2.Caption) + 1
Case 3
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\three.bmp")
L3.Caption = Val(L3.Caption) + 1
Case 4
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\four.bmp")
L4.Caption = Val(L4.Caption) + 1
Case 5
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\five.bmp")
L5.Caption = Val(L5.Caption) + 1
Case 6
imgDie.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\six.bmp")
L6.Caption = Val(L6.Caption) + 1
End Select
End Sub

Private Sub Stop_Click()
Roll.Enabled = False
End Sub

Private Sub Reset_Click()
L1.Caption = 0
L2.Caption = 0
L3.Caption = 0
L4.Caption = 0
L5.Caption = 0
L6.Caption = 0
End Sub

Private Sub End_Click()
End
End Sub





