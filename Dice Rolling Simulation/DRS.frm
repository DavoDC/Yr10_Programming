VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DRS - By D.C"
   ClientHeight    =   5040
   ClientLeft      =   5610
   ClientTop       =   3405
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8550
   Begin VB.CommandButton Roller 
      Caption         =   "Click to Roll Die"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   6375
   End
   Begin VB.Image imgDie1 
      Height          =   1575
      Left            =   6000
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image imgDie2 
      Height          =   1575
      Left            =   1080
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dice Rolling Simulation"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
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
      Width           =   8640
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
Public Sub RollDice()
Dim Die1 As Integer, Die2 As Integer
Randomize
Die1 = Int(Rnd * 6) + 1
Select Case Die1
Case 1
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\one.bmp")
Case 2
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\two.bmp")
Case 3
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\three.bmp")
Case 4
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\four.bmp")
Case 5
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\five.bmp")
Case 6
imgDie1.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\six.bmp")
End Select
Die2 = Int(Rnd * 6) + 1
Select Case Die2
Case 1
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\one.bmp")
Case 2
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\two.bmp")
Case 3
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\three.bmp")
Case 4
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\four.bmp")
Case 5
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\five.bmp")
Case 6
imgDie2.Picture = LoadPicture("C:\School\Programming\VB Programs\Dice Rolling Simulation\six.bmp")
End Select
End Sub
Private Sub End_Click()
End
End Sub

Private Sub Roller_Click()
RollDice
End Sub
