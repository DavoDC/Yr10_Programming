VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextBox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   6375
   End
   Begin VB.CheckBox Italic 
      BackColor       =   &H00808080&
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CheckBox Bold 
      BackColor       =   &H00404040&
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Instructions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Enter your text below"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Font Styles"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Menu End 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bold_Click()
If Bold.Value = 1 Then
TextBox.FontBold = True
Else
TextBox.FontBold = False
End If
End Sub

Private Sub End_Click()
End
End Sub

Private Sub Italic_Click()
If Italic.Value = 1 Then
TextBox.FontItalic = True
Else
TextBox.FontItalic = False
End If
End Sub
