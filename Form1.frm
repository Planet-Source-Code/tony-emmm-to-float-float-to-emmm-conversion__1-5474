VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4020
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox picemmm 
      Height          =   4335
      Left            =   900
      ScaleHeight     =   16
      ScaleMode       =   0  'User
      ScaleWidth      =   4096
      TabIndex        =   1
      Top             =   240
      Width           =   5355
   End
   Begin VB.PictureBox picf 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   65528
      ScaleMode       =   0  'User
      ScaleWidth      =   15.432
      TabIndex        =   0
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public flgStop As Boolean

Private Sub Command1_Click()
  flgStop = False
  test Me
End Sub


Private Sub Command2_Click()
  flgStop = True

    EFTest 0
    EFTest 256
    EFTest 1024
    EFTest 4096
    EFTest 5120
    EFTest 16384
    EFTest 17408
    EFTest 61440
    EFTest 62464
    EFTest 65535

End Sub

Private Sub EFTest(nEMMM As Long)
    Dim nF As Single
    Dim nXEMMM As Long

    nF = EMMM2Float(nEMMM)
    nXEMMM = Float2EMMM(nF)

    Debug.Print nEMMM, nF, nXEMMM
End Sub

Private Sub xEFTest(nEMMM As Long)
    Dim nE As Byte, nM As Integer, nF As Single
    Dim nXE As Byte, nXM As Integer, nXEMMM As Long

    nE = (nEMMM And 61440) / 4096&
    nM = nEMMM And 4095&

    nF = FloatFromEMMM(nE, nM)

    EMMMFromFloat nF, nXE, nXM

    nXEMMM = nXM + (nXE * 4096&)

    Debug.Print nEMMM, nE, nM, nF, nXE, nXM, nXEMMM

End Sub
