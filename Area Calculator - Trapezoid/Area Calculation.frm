VERSION 5.00
Begin VB.Form AreaCalc 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Area Calculator - By David Charkey "
   ClientHeight    =   7635
   ClientLeft      =   3825
   ClientTop       =   2040
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12855
   Begin VB.CommandButton ButtonCalc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate"
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   5160
      Width           =   5175
   End
   Begin VB.TextBox H1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox SideB 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox SideA 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Line Line11 
      X1              =   4680
      X2              =   4440
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line10 
      X1              =   4560
      X2              =   4560
      Y1              =   4440
      Y2              =   3120
   End
   Begin VB.Line Line8 
      X1              =   4440
      X2              =   4680
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line9 
      X1              =   3600
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label ResultArea 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Area :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Line Line7 
      Index           =   1
      X1              =   4320
      X2              =   4320
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line Line6 
      Index           =   1
      X1              =   720
      X2              =   720
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   720
      X2              =   4320
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   3840
      X2              =   3840
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line6 
      Index           =   0
      X1              =   1320
      X2              =   1320
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   1320
      X2              =   3840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label SideLabel2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Side B"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Area 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9360
      TabIndex        =   6
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label HeightLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Perpendicular Height"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label SideLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Side A"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   720
      X2              =   1320
      Y1              =   4440
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   3840
      X2              =   4320
      Y1              =   3120
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   1320
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   4320
      X2              =   720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Area of a Trapezoid"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu Resetter 
      Caption         =   "Reset"
   End
   Begin VB.Menu Tester 
      Caption         =   "Test"
   End
   Begin VB.Menu End 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "AreaCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonCalc_Click()
a = Val(SideA.Text)
b = Val(SideB.Text)
c = Val(H1.Text)
r = ((a + b) / 2) * c
Area.Caption = r
End Sub
Private Sub End_Click()
End
End Sub

Private Sub Resetter_Click()
SideA.Text = "0"
SideB.Text = "0"
H1.Text = "0"
Area.Caption = "0"
End Sub

Private Sub Tester_Click()
SideA.Text = "140"
SideB.Text = "160"
H1.Text = "10"
End Sub
