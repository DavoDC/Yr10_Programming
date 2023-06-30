VERSION 5.00
Begin VB.Form BB 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bouncing Ball by David Charkey"
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DropH 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Label Metres2 
      Alignment       =   2  'Center
      Caption         =   "Metres"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label TBounces 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
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
      Left            =   960
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label TDist 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
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
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label LabelTBounces 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total Number of Bounces"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Bouncing Ball"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Metres1 
      Alignment       =   2  'Center
      Caption         =   "Metres"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label LabelTotalD 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total Distance :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label LabelDropH 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Height Dropped :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Menu Instructions 
      Caption         =   "Instructions"
   End
   Begin VB.Menu Reset 
      Caption         =   "Reset"
   End
   Begin VB.Menu Quit 
      Caption         =   "Quit"
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "BB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calculate_Click()


If Val(DropH.Text) > 10 Then
 MsgBox " Invalid Value !  Please enter a valid value equal to or less than 10m "
 TDist.Caption = ""
 TBounces.Caption = ""
 DropH.Text = "0"
 TDist.Visible = False
 TBounces.Visible = False
 LabelTotalD.Visible = False
 Metres2.Visible = False
 LabelTBounces.Visible = False

End If


If (Val(DropH.Text) >= 0.01) And (Val(DropH.Text) < 10) Or (Val(DropH.Text) = 10) Then
 TDist.Caption = ""
 TBounces.Caption = ""
 TDist.Visible = True
 TBounces.Visible = True
 LabelTotalD.Visible = True
 Metres2.Visible = True
 LabelTBounces.Visible = True

Do While DropH > 0.01
 TBounces.Caption = Val(TBounces.Caption) + 1
 TDist.Caption = Format(Round(Val(TDist.Caption) + Val(DropH.Text) + (3 / 4) * Val(DropH.Text), 2), "#.##")
 DropH.Text = (3 / 4) * Val(DropH.Text)
Loop

DropH.Text = "0"
End If


End Sub

Private Sub Instructions_Click()
 MsgBox "To calculate the total distance the ball travels and total number of times the ball bounces when dropped from a specified height, press the button labelled 'Calculate'. To enter the specified height, click on the text box between the labels which have 'Height Dropped' and 'Metres' written on them."
End Sub

Private Sub Reset_Click()
 TDist.Visible = False
 TBounces.Visible = False
 LabelTotalD.Visible = False
 Metres2.Visible = False
 LabelTBounces.Visible = False
 TDist.Caption = ""
 TBounces.Caption = ""
 DropH.Text = "0"
End Sub


Private Sub About_Click()
 MsgBox "Created by David Charkey (2015)"
End Sub

Private Sub Quit_Click()
 End
End Sub

