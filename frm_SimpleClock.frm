VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_SimpleClock 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "simple clock with LED20 DEMO"
   ClientHeight    =   2280
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlSimpleClock 
      Left            =   5280
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "L"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1720
      FCol            =   8454016
      Val             =   ""
      OnWid           =   5
      Mod             =   1
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   5280
      Top             =   720
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1720
      FCol            =   8454016
      Val             =   ""
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   2
      Left            =   1620
      TabIndex        =   2
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1720
      FCol            =   8454016
      Val             =   ""
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   975
      Index           =   3
      Left            =   2370
      TabIndex        =   3
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1720
      FCol            =   8454016
      Val             =   ""
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   735
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1296
      FCol            =   8454016
      Val             =   ""
      OnWid           =   4
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   735
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1296
      FCol            =   8454016
      Val             =   ""
      OnWid           =   4
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   390
      Index           =   6
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
      FCol            =   8454016
      Val             =   "A"
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDClock 
      Height          =   390
      Index           =   7
      Left            =   5205
      TabIndex        =   7
      Top             =   120
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
      FCol            =   8454016
      Val             =   "M"
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   1
      Left            =   615
      TabIndex        =   9
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "E"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   2
      Left            =   990
      TabIndex        =   10
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "D"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   3
      Left            =   1365
      TabIndex        =   11
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "2"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   4
      Left            =   1740
      TabIndex        =   12
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "0"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "D"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   6
      Left            =   2655
      TabIndex        =   14
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "E"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   7
      Left            =   3030
      TabIndex        =   15
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "M"
      OnWid           =   5
      Mod             =   1
   End
   Begin prjSimpleClock.LED20 LEDLogo 
      Height          =   375
      Index           =   8
      Left            =   3405
      TabIndex        =   16
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      FCol            =   65535
      WCol            =   49344
      BCol            =   8421504
      Val             =   "O"
      OnWid           =   5
      Mod             =   1
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "Clock Settings"
      Begin VB.Menu mnuStyle 
         Caption         =   "Simple"
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Color"
         Begin VB.Menu mnuColOpt 
            Caption         =   "Fore"
            Index           =   0
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Back"
            Index           =   1
         End
         Begin VB.Menu mnuColOpt 
            Caption         =   "Wire"
            Index           =   2
         End
      End
      Begin VB.Menu mnuShadow 
         Caption         =   "Shadow ON"
      End
      Begin VB.Menu mnuWid 
         Caption         =   "OnWidth"
         Begin VB.Menu mnuWidOpt 
            Caption         =   "2"
            Index           =   0
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "3"
            Index           =   1
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "4"
            Index           =   2
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "5"
            Index           =   3
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "6"
            Index           =   4
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "7"
            Index           =   5
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "8"
            Index           =   6
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "9"
            Index           =   7
         End
         Begin VB.Menu mnuWidOpt 
            Caption         =   "10"
            Index           =   8
         End
      End
   End
End
Attribute VB_Name = "frm_SimpleClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  MenuChecking mnuWidOpt, LEDClock(0).ONWidth - 2

End Sub

Private Sub mnuColOpt_Click(Index As Integer)

  Dim I       As Long
  Dim TestCol As Long

  Select Case Index
   Case 0
    TestCol = LEDClock(0).ForeColor
   Case 1
    TestCol = LEDClock(0).BackColor
   Case 2
    TestCol = LEDClock(0).WireColor
  End Select
  With cdlSimpleClock
    .Flags = cdlCCRGBInit Or cdlCCFullOpen
    .Color = TestCol
    .ShowColor
    TestCol = .Color
  End With
  Select Case Index
   Case 0
    For I = 0 To 7
      LEDClock(I).ForeColor = TestCol
    Next I
   Case 1
    For I = 0 To 7
      LEDClock(I).BackColor = TestCol
    Next I
   Case 2
    For I = 0 To 7
      LEDClock(I).WireColor = TestCol
    Next I
  End Select

End Sub

Private Sub mnuExit_Click()

  tmrUpdate.Enabled = False
  Unload Me

End Sub

Private Sub mnuShadow_Click()

  Dim I As Long

  For I = 0 To 7
    LEDClock(I).OFFShadow = Not LEDClock(I).OFFShadow
  Next I
  mnuShadow.Caption = "Shadow " & IIf(LEDClock(0).OFFShadow, "OFF", "ON")

End Sub

Private Sub mnuStyle_Click()

  Dim I As Long

  For I = 0 To 5
    LEDClock(I).Value = " "
    If LEDClock(I).Mode = eFancy Then
      LEDClock(I).Mode = eSimple
     Else
      LEDClock(I).Mode = eFancy
    End If
  Next I
  mnuStyle.Caption = IIf(LEDClock(0).Mode = eFancy, "Simple", "Fancy")

End Sub

Private Sub mnuWidOpt_Click(Index As Integer)

  Dim I As Long

  MenuChecking mnuWidOpt, Index
  For I = 0 To 5
    LEDClock(I).ONWidth = Index + 2
  Next I

End Sub

Private Sub tmrUpdate_Timer()

  Dim strTime2 As String

  strTime2 = Format$(Time, "hh:mm:ss AM/PM")
  LEDClock(0).Value = Left$(strTime2, 1)
  LEDClock(1).Value = Mid$(strTime2, 2, 1)
  LEDClock(2).Value = Mid$(strTime2, 4, 1)
  LEDClock(3).Value = Mid$(strTime2, 5, 1)
  LEDClock(4).Value = Mid$(strTime2, 7, 1)
  LEDClock(5).Value = Mid$(strTime2, 8, 1)
  LEDClock(6).Value = Left$(Right$(strTime2, 2), 1)
  With LEDLogo 'Only affects the count property of the control array
    LEDLogo(Int(Rnd * .Count)).WireColor = Rnd * vbWhite
    LEDLogo(Int(Rnd * .Count)).ForeColor = Rnd * vbWhite
    LEDLogo(Int(Rnd * .Count)).BackColor = vbWhite - LEDLogo(Int(Rnd * .Count)).ForeColor
    LEDLogo(Int(Rnd * .Count)).OFFShadow = Rnd > 0.5
  End With 'LEDLogo

End Sub

':)Code Fixer V2.8.9 (19/01/2005 3:48:13 PM) 1 + 111 = 112 Lines Thanks Ulli for inspiration and lots of code.
