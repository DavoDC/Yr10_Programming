VERSION 5.00
Begin VB.Form VolumeCalc 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volume Calculator - By David Charkey"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Radius 
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   7
      Text            =   "0"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox H 
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   6
      Text            =   "0"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   1080
      Picture         =   "VolumeCalc.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Result 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label LabelV 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Volume :"
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
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label LabelR 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Radius :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label LabelH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Height :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Volume of a Cone"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
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
      Width           =   12735
   End
End
Attribute VB_Name = "VolumeCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VolCalc()
Const Pi = 3.14
Result.Caption = Format(1 / 3 * Pi * Val(Radius.Text) ^ 2 * Val(H.Text), "#,##0.0")
End Sub

Private Sub H_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
VolCalc
End If
End Sub

Private Sub Radius_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
VolCalc
End If
End Sub

