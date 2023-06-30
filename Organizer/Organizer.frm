VERSION 5.00
Begin VB.Form Organizer 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Stats 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   1440
      TabIndex        =   1
      Text            =   "2910"
      Top             =   1200
      Width           =   8655
   End
   Begin VB.CommandButton Organize 
      Caption         =   "Organize"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Menu Info 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Organizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Info_Click()
MsgBox "Made by David Charkey 2015"
End Sub


Private Sub Organize_Click()

End Sub
