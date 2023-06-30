VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator - Created By David Charkey - Copyright© - All Rights Reserved"
   ClientHeight    =   7425
   ClientLeft      =   4230
   ClientTop       =   1470
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12000
   Begin VB.CommandButton Cube 
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   14
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Square 
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Operation 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      MaskColor       =   &H00808080&
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Ans2 
      Caption         =   "Input Answer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   2640
      Width           =   3615
   End
   Begin VB.CommandButton Ans 
      Caption         =   "Input Answer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   2640
      Width           =   3615
   End
   Begin VB.CommandButton Hax 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Multiply 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Divide 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Subtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Add 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox V2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   6960
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox V1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Result 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
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
      Left            =   4440
      TabIndex        =   2
      Top             =   6120
      Width           =   6735
   End
   Begin VB.Label Answer 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
   End
   Begin VB.Menu Reset 
      Caption         =   "Reset"
   End
   Begin VB.Menu End 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Add_Click()
Operation.Caption = "+"
a = Val(V1.Text)
b = Val(V2.Text)
r = a + b
Result.Caption = r
End Sub

Private Sub Ans_Click()
V1.Text = Result.Caption
End Sub

Private Sub Cube_Click()
a = Val(V1.Text)
b = Val(V1.Text)
c = Val(V1.Text)
r = a * b * c
Result.Caption = r
V2.Text = "0"
End Sub

Private Sub Divide_Click()
Operation.Caption = "÷"
a = Val(V1.Text)
b = Val(V2.Text)
r = a / b
Result.Caption = r
End Sub

Private Sub End_Click()
End
End Sub

Private Sub Hax_Click()
Operation.Caption = "^"
a = Val(V1.Text)
b = Val(V2.Text)
r = a ^ b
Result.Caption = r
End Sub

Private Sub Multiply_Click()
Operation.Caption = "x"
a = Val(V1.Text)
b = Val(V2.Text)
r = a * b
Result.Caption = r
End Sub

Private Sub Operation_Click()
Operation.FontBold = True
End Sub

Private Sub Reset_Click()
V1.Text = "0"
V2.Text = "0"
Result.Caption = "0"
End Sub

Private Sub Square_Click()
Operation.Caption = "x"
a = Val(V1.Text)
b = Val(V1.Text)
r = a * b
Result.Caption = r
V2.Text = Val(V1.Text)
End Sub

Private Sub Subtract_Click()
Operation.Caption = "-"
a = Val(V1.Text)
b = Val(V2.Text)
r = a - b
Result.Caption = r
End Sub

Private Sub Test_Click()
V1.Text = "20"
V2.Text = "10"
Result.Caption = "0"
End Sub

Private Sub Ans2_Click()
V2.Text = Result.Caption
End Sub


Private Sub Trigger_Click()
Test.Caption = "Test"
End Sub


