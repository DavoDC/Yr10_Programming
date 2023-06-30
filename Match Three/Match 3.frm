VERSION 5.00
Begin VB.Form Interface 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Match 3 Game ©Copyright - All Rights Reserved"
   ClientHeight    =   8280
   ClientLeft      =   5025
   ClientTop       =   1065
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10905
   Begin VB.CommandButton Trigger 
      Caption         =   "Click here to have a shot at matching the images!"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   9765
   End
   Begin VB.Label PE 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label PEL 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Points Earned this Turn :"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5760
      TabIndex        =   9
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label HScore 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label HS 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "High Score :"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Turn 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Turns :"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   6000
      TabIndex        =   5
      Top             =   5160
      Width           =   15
   End
   Begin VB.Label Turns 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Result 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Image S3 
      Height          =   2535
      Left            =   7320
      Top             =   840
      Width           =   2895
   End
   Begin VB.Image S2 
      Height          =   2535
      Left            =   4080
      Top             =   840
      Width           =   2655
   End
   Begin VB.Image S1 
      Height          =   2535
      Left            =   600
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Match 3 Game"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
   End
   Begin VB.Menu Help 
      Caption         =   "Instructions"
   End
   Begin VB.Menu Hisoka 
      Caption         =   "Wild Card Explanation"
   End
   Begin VB.Menu Reset 
      Caption         =   "Restart"
   End
   Begin VB.Menu Info 
      Caption         =   "About"
   End
   Begin VB.Menu End 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Trigger_Click()

Turns.Caption = Val(Turns.Caption) - 1
PE.Caption = 0

Dim Slot1 As Integer
Dim Slot2 As Integer
Dim Slot3 As Integer
Dim TotalS As Integer

Randomize

Slot1 = Int(Rnd * 11) + 1
Select Case Slot1

Case 1
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\\Pic1.bmp")

Case 2
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic2.bmp")

Case 3
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic3.bmp")

Case 4
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic4.bmp")

Case 5
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic5.bmp")

Case 6
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic6.bmp")

Case 7
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic7.bmp")

Case 8
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic8.bmp")

Case 9
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic9.bmp")

Case 10
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic10.bmp")

Case 11
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Wild.bmp")

End Select

Slot2 = Int(Rnd * 11) + 1

Select Case Slot2

Case 1
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic1.bmp")

Case 2
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic2.bmp")

Case 3
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic3.bmp")

Case 4
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic4.bmp")

Case 5
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic5.bmp")

Case 6
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic6.bmp")

Case 7
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic7.bmp")

Case 8
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic8.bmp")

Case 9
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic9.bmp")

Case 10
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic10.bmp")

Case 11
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Wild.bmp")

End Select

Slot3 = Int(Rnd * 11) + 1

Select Case Slot3

Case 1
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic1.bmp")

Case 2
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic2.bmp")

Case 3
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic3.bmp")

Case 4
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic4.bmp")

Case 5
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic5.bmp")

Case 6
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic6.bmp")

Case 7
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic7.bmp")

Case 8
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic8.bmp")

Case 9
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic9.bmp")

Case 10
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Pic10.bmp")

Case 11
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Wild.bmp")


End Select

If (Val(Turns.Caption) = 0) Then
MsgBox "You have used up all 20 turns!.   If you have achieved a High Score for this round it will appear in the space below after closing this message"
End If

If (Slot1 = Slot2) And (Slot2 = Slot3) Then
MsgBox " Well Done! You matched three!"
Result.Caption = Val(Result.Caption) + 100
PE.Caption = Val(PE.Caption) + 100
End If


If (Slot1 = Slot2) Or (Slot2 = Slot3) Or (Slot1 = Slot3) Then
Result.Caption = Val(Result.Caption) + 20
PE.Caption = Val(PE.Caption) + 20
End If





If (Slot1 = 11) Or (Slot2 = 11) Or (Slot3 = 11) Then
Result.Caption = Val(Result.Caption) + 5
PE.Caption = Val(PE.Caption) + 5
End If


If (Slot1 = 11) And (Slot2 = 11) And (Slot3 = 11) Then
MsgBox " Well Done! You matched three!"
Result.Caption = Val(Result.Caption) + 115
PE.Caption = Val(PE.Caption) + 115

End If

If (Slot1 = 11) And (Slot2 = 11) Then
Result.Caption = Val(Result.Caption) + 10
PE.Caption = Val(PE.Caption) + 10
End If


If (Slot2 = 11) And (Slot3 = 11) Then
Result.Caption = Val(Result.Caption) + 10
PE.Caption = Val(PE.Caption) + 10
End If


If (Slot1 = 11) And (Slot3 = 11) Then
Result.Caption = Val(Result.Caption) + 10
PE.Caption = Val(PE.Caption) + 10
End If

If (Val(Turns.Caption) = 0) And (Val(HScore.Caption) < Val(Result.Caption)) Then
MsgBox "Well Done ! You beat your High Score!"
HScore.Caption = Val(Result.Caption)
Turns.Caption = Val(Turns.Caption) + 20
Result.Caption = 0
PE.Caption = 0
End If


If (Val(Turns.Caption) = 0) And (Val(HScore.Caption) > Val(Result.Caption)) Then
HScore.Caption = Val(HScore.Caption)
Turns.Caption = Val(Turns.Caption) + 20
Result.Caption = 0
PE.Caption = 0
End If


If (Val(Turns.Caption) = 0) And (Val(HScore.Caption) = Val(Result.Caption)) Then
HScore.Caption = Val(HScore.Caption)
Turns.Caption = Val(Turns.Caption) + 20
Result.Caption = 0
PE.Caption = 0
End If






End Sub

Private Sub Reset_Click()
Result.Caption = 0
HScore.Caption = 0
PE.Caption = 0
Turns.Caption = 20
S3.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Blank.bmp")
S2.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Blank.bmp")
S1.Picture = LoadPicture("C:\Users\charkeyd\Desktop\Other\School Years\Year 10\Programming\VB Programs\Match Three\Pictures\Blank.bmp")
End Sub

Private Sub End_Click()
MsgBox "Thankyou for playing Match 3"
End
End Sub


Private Sub Help_Click()
MsgBox "The objective of the game is to score as many points as possible in 20 turns.          You win 20 points when any two frames show the same graphic and 120 points when all three frames show the same graphic ( The Wild Card is an exception to these rules)."

End Sub



Private Sub Hisoka_Click()
MsgBox "You win 5 points when any frame displays the wild card graphic.   If two frames show the wild card graphic, you will score 20 points for matching two graphics and 10 points for getting 2 wild card graphics. You will also get 5 bonus points for this feat (Total = 20 + (5x2) + 5 = 35).     If all three frames show the wild card graphic, you will score 100 points for matching three graphics and 15 points for getting 3 wild card graphics . You will also gain 155 bonus points by making such a huge achievement (Total = 100 + (5x3) + 155 = 270)."

End Sub



Private Sub Info_Click()
MsgBox " Match 3 Game ©Copyright - All Rights Reserved    -   Created by David Charkey (2015) "
End Sub





