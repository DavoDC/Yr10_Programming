VERSION 5.00
Begin VB.Form RandNG 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Number Generator  - By David Charkey"
   ClientHeight    =   5850
   ClientLeft      =   4935
   ClientTop       =   3360
   ClientWidth     =   8385
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8385
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton Generator 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Caption         =   "Random Number Generator"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8475
   End
   Begin VB.Label Result 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   8475
   End
   Begin VB.Menu Help 
      Caption         =   "Information"
   End
   Begin VB.Menu Close 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "RandNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
End
End Sub

Private Sub Generator_Click()
Dim RNG As Integer
Randomize
RNG = Int((1000 * Rnd) + 1)
Result.Caption = RNG
End Sub

Private Sub Help_Click()
MsgBox "Generates random numbers between 1 and 1000.                                                     Created by David Charkey (2015)"
End Sub


Private Sub Reset_Click()
Result.Caption = "0"
End Sub
