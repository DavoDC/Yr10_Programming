VERSION 5.00
Begin VB.Form TOS 
   BackColor       =   &H8000000E&
   Caption         =   "Table of Squares"
   ClientHeight    =   6045
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Print 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Display 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox picDisplay 
      Height          =   4815
      Left            =   480
      ScaleHeight     =   4755
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "TOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Display_Click()

Dim Num As Integer
picDisplay.Cls
picDisplay.FontUnderline = True
picDisplay.FontSize = 20
picDisplay.Print "Table of Squares"
picDisplay.FontUnderline = False
picDisplay.Print
For Num = 1 To 10
picDisplay.Print Num, Num ^ 2
Next Num

End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Print_Click()

Dim Num As Integer
MsgBox "Checkprinter is on and on-line", vbExclamation, "Table of Squares"
With Printer
.FontName = "Calibri"
.FontSize = 14
.FontUnderline = True
Printer.Print "Table of Squares"
.FontUnderline = False
End With
Printer.Print
For Num = 1 To 10
Printer.Print Num, Num ^ 2
Next Num
Printer.EndDoc

End Sub


Private Sub Clear_Click()

picDisplay.Cls


End Sub
