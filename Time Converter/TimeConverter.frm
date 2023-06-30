VERSION 5.00
Begin VB.Form TimeConverter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Converter - By David Charkey"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7455
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar TS 
      Height          =   5295
      Left            =   600
      Max             =   5000
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label HVal 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label LabelH 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label M2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   4080
      Width           =   2265
   End
   Begin VB.Label M1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   3345
   End
   Begin VB.Label LabelM2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   2385
   End
   Begin VB.Label LabeLM 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   3345
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Time Converter"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Menu end 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "TimeConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub end_Click()
End
End Sub

Private Sub TS_Change()
M1.Caption = TS.Value
HVal.Caption = Format(TS.Value / 60, "0")
M2.Caption = Format(TS.Value Mod 60, "0")
End Sub
