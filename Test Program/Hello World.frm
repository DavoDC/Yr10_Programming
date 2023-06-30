VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hello World"
   ClientHeight    =   5475
   ClientLeft      =   4305
   ClientTop       =   2400
   ClientWidth     =   13440
   DrawMode        =   1  'Blackness
   FillColor       =   &H00800000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   13440
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton Erase 
      Caption         =   "Clear Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Close Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6360
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Display 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Name 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Card Maker 5000"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Erase_Click()
Display.Text = "."
End Sub

Private Sub Name_Click()

End Sub

Private Sub Quit_Click(Index As Integer)
End
End Sub

Private Sub Text1_Change()

End Sub
