VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Calculator (Created By David Charkey) Copyright ©  - All Rights Reserved"
   ClientHeight    =   10650
   ClientLeft      =   -75
   ClientTop       =   645
   ClientWidth     =   20250
   FillColor       =   &H00404040&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   Begin VB.TextBox D2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   16800
      TabIndex        =   27
      Text            =   "0"
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox D1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   26
      Text            =   "0"
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox M2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16920
      TabIndex        =   25
      Text            =   "0"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox M1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   24
      Text            =   "0"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton CalculateD 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   10920
      TabIndex        =   23
      Top             =   7440
      Width           =   8655
   End
   Begin VB.CommandButton ResetD 
      BackColor       =   &H0000FFFF&
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
      Height          =   1095
      Index           =   2
      Left            =   17160
      TabIndex        =   22
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton CalculateM 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   10920
      TabIndex        =   21
      Top             =   2400
      Width           =   8535
   End
   Begin VB.CommandButton ResetM 
      BackColor       =   &H0000FFFF&
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
      Height          =   1095
      Index           =   1
      Left            =   17040
      TabIndex        =   20
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox S2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   13
      Text            =   "0"
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox S1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2160
      TabIndex        =   12
      Text            =   "0"
      Top             =   6015
      Width           =   2055
   End
   Begin VB.TextBox A2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   11
      Text            =   "0"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox A1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton CalculateS 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   7440
      Width           =   9135
   End
   Begin VB.CommandButton ResetS 
      BackColor       =   &H0000FFFF&
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
      Height          =   1095
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton ResetA 
      BackColor       =   &H0000FFFF&
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
      Height          =   1095
      Index           =   0
      Left            =   6960
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton CalculateA 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
   End
   Begin VB.Label ResultM 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13560
      TabIndex        =   31
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label ResultD 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13560
      TabIndex        =   30
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Label ResultS 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3480
      TabIndex        =   29
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label ResultA 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3000
      TabIndex        =   28
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   10920
      TabIndex        =   19
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Multiply 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   15000
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   11040
      TabIndex        =   17
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label divide 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   15240
      TabIndex        =   16
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Calc2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Subtraction"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   10935
   End
   Begin VB.Label Calc 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Addition"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Calc 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Multiplication"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   10080
      TabIndex        =   15
      Top             =   0
      Width           =   10695
   End
   Begin VB.Label Division 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   10680
      TabIndex        =   14
      Top             =   4800
      Width           =   10695
   End
   Begin VB.Label Subtract 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   4680
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label Add 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Menu CalcAll 
      Caption         =   "Calculate All"
   End
   Begin VB.Menu Resetter 
      Caption         =   "Reset All"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu test 
      Caption         =   "Test"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu End 
      Caption         =   "Quit "
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CalcAll_Click()
a = Val(A1.Text)
b = Val(A2.Text)
r = a + b
ResultA.Caption = r
a = Val(D1.Text)
b = Val(D2.Text)
r = a / b
ResultD.Caption = r
a = Val(M1.Text)
b = Val(M2.Text)
r = a * b
ResultM.Caption = r
a = Val(S1.Text)
b = Val(S2.Text)
r = a - b
ResultS.Caption = r
End Sub

Private Sub CalculateA_Click(Index As Integer)
a = Val(A1.Text)
b = Val(A2.Text)
r = a + b
ResultA.Caption = r
End Sub

Private Sub CalculateD_Click(Index As Integer)
a = Val(D1.Text)
b = Val(D2.Text)
r = a / b
ResultD.Caption = r
End Sub

Private Sub CalculateM_Click(Index As Integer)
a = Val(M1.Text)
b = Val(M2.Text)
r = a * b
ResultM.Caption = r
End Sub

Private Sub CalculateS_Click(Index As Integer)
a = Val(S1.Text)
b = Val(S2.Text)
r = a - b
ResultS.Caption = r
End Sub

Private Sub Haxorer_Click()
A1.Text = "69"
A2.Text = "69"
ResultA.Caption = "D.C"
M1.Text = "69"
M2.Text = "69"
ResultM.Caption = "D.C"
D1.Text = "69"
D2.Text = "69"
ResultD.Caption = "D.C"
S1.Text = "69"
S2.Text = "69"
ResultS.Caption = "D.C"
End Sub

Private Sub End_Click()
End
End Sub

Private Sub ResetA_Click(Index As Integer)
A1.Text = "0"
A2.Text = "0"
ResultA.Caption = "0"
End Sub

Private Sub ResetD_Click(Index As Integer)
D1.Text = "0"
D2.Text = "0"
ResultD.Caption = "0"
End Sub

Private Sub ResetM_Click(Index As Integer)
M1.Text = "0"
M2.Text = "0"
ResultM.Caption = "0"
End Sub

Private Sub ResetS_Click(Index As Integer)
S1.Text = "0"
S2.Text = "0"
ResultS.Caption = "0"
End Sub

Private Sub Resetter_Click()
A1.Text = "0"
A2.Text = "0"
ResultA.Caption = "0"
M1.Text = "0"
M2.Text = "0"
ResultM.Caption = "0"
D1.Text = "0"
D2.Text = "0"
ResultD.Caption = "0"
S1.Text = "0"
S2.Text = "0"
ResultS.Caption = "0"
End Sub

Private Sub test_Click()
A1.Text = "100"
A2.Text = "55"
M1.Text = "10"
M2.Text = "5"
D1.Text = "100"
D2.Text = "20"
S1.Text = "500"
S2.Text = "100"
End Sub
