VERSION 5.00
Begin VB.UserControl LED20 
   BackColor       =   &H80000008&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LockControls    =   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   2985
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   19
      X1              =   2805
      X2              =   2805
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   18
      X1              =   2670
      X2              =   2670
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   17
      X1              =   2535
      X2              =   2535
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   16
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   15
      X1              =   2265
      X2              =   2265
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   14
      X1              =   2130
      X2              =   2130
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   13
      X1              =   1995
      X2              =   1995
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   12
      X1              =   1860
      X2              =   1860
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   11
      X1              =   1725
      X2              =   1725
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   10
      X1              =   1590
      X2              =   1590
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   9
      X1              =   1455
      X2              =   1455
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   8
      X1              =   1320
      X2              =   1320
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   7
      X1              =   1185
      X2              =   1185
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   6
      X1              =   1050
      X2              =   1050
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      X1              =   915
      X2              =   915
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      X1              =   780
      X2              =   780
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      X1              =   645
      X2              =   645
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      X1              =   510
      X2              =   510
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   375
      X2              =   375
      Y1              =   240
      Y2              =   400
   End
   Begin VB.Line Wire 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   400
   End
End
Attribute VB_Name = "LED20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Just a quickie LED UserControl
'LED20 because it uses 20 'Wires' to display characters
'This is a spin off from the LED clock challange but definately not a contender (way too many lines)
'If you look at the control in IDE you'll see that there has been no attempt to position the wires ahead of time
'In case you do want to edit it please note that the form has been locked to stop unnecessary fiddling about.
'
'PROPERTIES
'Value      = Letter(Ucase only)/number to display set to " " for all Wires off (see Mode for limitations)
'Mode       = eSimple = Basic LED numbers only; using only the vertical and horizontal lines
'           = eFancy  = Letters and numbers
'BackColour = backfield of control
'WireColor  = colour to use for Off Wires. Usually a duller version of the ForColor
'             Set to match BackColor and OffShadow = False for invisible
'ForeColor  = colour to display Value/On Wires
'ONWidth    = how thick the OnWires wires are (2 to ???( depends on size of control))
'OFFShadow  = True  = Off Wires overlay On Wires (more machine look)
'           = False = On Wires overlay Off Wires (smoother characters)
'
'see UserControl_Initialize for default values
'
'PROCEDURES
'SetLine = makes positioning the lines easer by making code more readable(called by 'UserControl_Resize')
'InArray = simple test that a value is in an array (Designed to use paramArrays could be changed to handle ordinary pre-defined arrays)
'LiteUp  =switchboard decides which Mode to use
'DoOnColour = Change Colour of On Wires to ForeColour
'Blocky     = Rules for basic 7 line numerals used by Mode = eSimple
'FancyAlpha  = Rules for 20 line chracters used by Mode = eFancy
'UserControl_Initialize = Set intial appearance when you create a new control
'UserControl_Resize = layout and size lines.
'
'Suggestions/improvements welcome
'
'ADDING/EDITING CHARACTER
'see 'FancyAlpha' for how to create your own characters
'1. Create a copy of 'FancyAlpha' changing the name to suit
'2. Add an Enum member to Enum blkMode
'3. and a 'Case yourEnumName' to 'LiteUp'
'
'If you craft your own set of characters
'and send them to me I'll add them to future versions
'
Private m_value         As String
Public Enum blkMode
  eSimple
  eFancy
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eSimple, eFancy
#End If
Private m_FColour       As OLE_COLOR
Private m_WColour       As OLE_COLOR
Private m_OnWidth       As Long
Private m_OFFShadow     As Boolean
Private m_Mode          As blkMode

Public Property Get BackColor() As OLE_COLOR

  BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewCol As OLE_COLOR)

  UserControl.BackColor = NewCol
  PropertyChanged "BCol"

End Property

Private Sub Blocky()

  'numbers only the classic LED numerals
  ' use the headers to work out what turns on
  'TopLine =0,1

  Wire(0).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 7, 8, 9), ONWidth, 1)
  Wire(1).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 7, 8, 9), ONWidth, 1)
  'Right =2,3
  Wire(2).BorderWidth = IIf(InArray(m_value, 0, 1, 2, 3, 4, 7, 8, 9), ONWidth, 1)
  Wire(3).BorderWidth = IIf(InArray(m_value, 0, 1, 3, 4, 5, 6, 7, 8, 9), ONWidth, 1)
  'Bottom =4,5
  Wire(4).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 8, 9), ONWidth, 1)
  Wire(5).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 8, 9), ONWidth, 1)
  'left = 6,7
  Wire(6).BorderWidth = IIf(InArray(m_value, 0, 2, 6, 8), ONWidth, 1)
  Wire(7).BorderWidth = IIf(InArray(m_value, 0, 4, 5, 6, 8, 9), ONWidth, 1)
  '' if you want a centred '1' then uncomment these and
  '' delete the 1's in wire(2) and Wire(3) arrays
  '    Line1(8).BorderWidth = IIf(InArray(m_value, 1), ONWidth, 1)
  '    Line1(9).BorderWidth = IIf(InArray(m_value, 1), ONWidth, 1)
  '9-3 =10,11
  Wire(10).BorderWidth = IIf(InArray(m_value, 2, 3, 4, 5, 6, 8, 9), ONWidth, 1)
  Wire(11).BorderWidth = IIf(InArray(m_value, 2, 3, 4, 5, 6, 8, 9), ONWidth, 1)

End Sub

Private Sub DoOnColour()

  Dim I As Long

  For I = 0 To Wire.Count - 1
    Wire(I).BorderColor = WireColor
    If m_OFFShadow Then
      Wire(I).ZOrder 0
    End If
  Next I
  For I = 0 To Wire.Count - 1
    If Wire(I).BorderWidth > 1 Then
      Wire(I).BorderColor = ForeColor
      If Not m_OFFShadow Then
        Wire(I).ZOrder 0
      End If
    End If
  Next I

End Sub

Private Sub FancyAlpha()

  'For each character use the chart to work out which lines to turn on
  'and add the character to that line's activating array
  ' ----0---' '---1---'
  '| \    /  | \    / |
  '|  \  /   |  \  /  |
  '7   \/    8   \/   2
  '|   /\    |   /\   |
  '| 16  12  | 14  17 |
  '| /    \  |/      \|
  '----10---' '---11--'
  '| \    /  | \    / |
  '|  \  /   |  \  /  |
  '6   \/    9   \/   3
  '|   /\    |   /\   |
  '| 15  19  | 18  13 |
  '| /    \  |/      \|
  ' ----5----''--4----'
  'TopLeft

  Wire(0).BorderWidth = IIf(InArray(m_value, 2, 3, 5, 7, "B", "D", "P", "R", "T", "Z"), ONWidth, 1)
  'TopRight
  Wire(1).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 7, 9, "B", "C", "E", "F", "G", "R", "S", "T", "Z"), ONWidth, 1)
  'RightTop
  Wire(2).BorderWidth = IIf(InArray(m_value, 0, 2, 9, "H", "M", "N", "U", "V", "W"), ONWidth, 1)
  'RightBottom
  Wire(3).BorderWidth = IIf(InArray(m_value, "A", "G", "H", "M", "N", "W"), ONWidth, 1)
  'BottomRight
  Wire(4).BorderWidth = IIf(InArray(m_value, 2, 3, 5, 6, 8, "B", "C", "E", "L", "Z"), ONWidth, 1)
  'BottomLEft
  Wire(5).BorderWidth = IIf(InArray(m_value, 0, 2, 3, 5, 6, 8, 9, "B", "D", "L", "U", "S", "Z"), ONWidth, 1)
  'LeftBottom
  Wire(6).BorderWidth = IIf(InArray(m_value, 0, 6, "A", "B", "D", "F", "H", "K", "L", "M", "N", "P", "U", "R", "W"), ONWidth, 1)
  'LeftTop
  Wire(7).BorderWidth = IIf(InArray(m_value, 5, "B", "D", "H", "K", "L", "M", "N", "P", "R", "U", "V", "W"), ONWidth, 1)
  'MidVertTop
  Wire(8).BorderWidth = IIf(InArray(m_value, 1, 4, "I", "J", "T"), ONWidth, 1)
  'MidVertBottom
  Wire(9).BorderWidth = IIf(InArray(m_value, 1, 4, "I", "J", "T", "Y"), ONWidth, 1)
  '9-3 =10,11
  'MidHorzLeft
  Wire(10).BorderWidth = IIf(InArray(m_value, 4, 5, 6, 8, 9, "A", "B", "E", "F", "H", "K", "P", "R", "S"), ONWidth, 1)
  'MidHorzRight
  Wire(11).BorderWidth = IIf(InArray(m_value, 2, 4, 8, 9, "A", "E", "F", "H", "P", "S"), ONWidth, 1)
  'centre-11
  Wire(12).BorderWidth = IIf(InArray(m_value, "M", "N", "X", "Y"), ONWidth, 1)
  'centre-4
  Wire(13).BorderWidth = IIf(InArray(m_value, 5, 3, 6, 8, "B", "K", "N", "Q", "R", "W", "X"), ONWidth, 1)
  'centre-2
  Wire(14).BorderWidth = IIf(InArray(m_value, 3, 7, "B", "K", "M", "R", "X", "Y", "Z"), ONWidth, 1)
  'centre-7
  Wire(15).BorderWidth = IIf(InArray(m_value, 2, 7, 8, "W", "X", "Z"), ONWidth, 1)
  'clock face positioning
  '9-12
  Wire(16).BorderWidth = IIf(InArray(m_value, 0, 4, 6, 8, 9, "A", "C", "E", "F", "G", "O", "Q", "S"), ONWidth, 1)
  '12-3
  Wire(17).BorderWidth = IIf(InArray(m_value, 8, "A", "D", "O", "P", "Q"), ONWidth, 1)
  '3-6
  Wire(18).BorderWidth = IIf(InArray(m_value, 0, 9, "D", "G", "O", "Q", "S", "U", "V"), ONWidth, 1)
  '6-9
  Wire(19).BorderWidth = IIf(InArray(m_value, "C", "E", "G", "J", "O", "Q", "V"), ONWidth, 1)

End Sub

Public Property Get ForeColor() As OLE_COLOR

  ForeColor = m_FColour

End Property

Public Property Let ForeColor(ByVal NewCol As OLE_COLOR)

  m_FColour = NewCol
  PropertyChanged "FCol"
  DoOnColour

End Property

Private Function InArray(ByVal Test As String, _
                         ParamArray arr() As Variant) As Boolean

  Dim P As Variant

  For Each P In arr
    If Test = P Then
      InArray = True
    End If
  Next P

End Function

Private Sub LiteUp()

  If ONWidth Then ' safety to prevent trying to draw if ONWidth has not been set
    Select Case Mode
     Case eSimple
      Blocky
     Case eFancy
      FancyAlpha
    End Select
    DoOnColour
  End If

End Sub

Public Property Get Mode() As blkMode

  Mode = m_Mode

End Property

Public Property Let Mode(NewMode As blkMode)

  m_Mode = NewMode
  PropertyChanged "Mod"

End Property

Public Property Get OFFShadow() As Boolean

  OFFShadow = m_OFFShadow

End Property

Public Property Let OFFShadow(ByVal BShow As Boolean)

  m_OFFShadow = BShow

End Property

Public Property Get ONWidth() As Long

  ONWidth = m_OnWidth

End Property

Public Property Let ONWidth(ByVal NewWid As Long)

  If NewWid > 0 Then
    m_OnWidth = NewWid
    PropertyChanged "ONWid"
    LiteUp
  End If

End Property

Private Sub SetLine(ByVal lno As Long, _
                    ByVal w1 As Long, _
                    ByVal h1 As Long, _
                    ByVal w2 As Long, _
                    ByVal h2 As Long, _
                    Optional ByVal Col As Long = vbWhite)

  'this just makes it easier to see what is going on in UserControl_Resize

  With Wire(lno)
    .X1 = w1
    .Y1 = h1
    .X2 = w2
    .Y2 = h2
    .BorderColor = Col
  End With

End Sub

Private Sub UserControl_Initialize()

  ForeColor = vbGreen
  WireColor = &HC000&
  BackColor = vbBlack
  Mode = eFancy
  ONWidth = 3
  OFFShadow = True

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  With PropBag
    ForeColor = .ReadProperty("FCol", vbGreen)
    WireColor = .ReadProperty("WCol", &HC000&)
    BackColor = .ReadProperty("BCol", vbBlack)
    OFFShadow = .ReadProperty("Shad", True)
    ONWidth = .ReadProperty("OnWid", 3)
    Mode = .ReadProperty("Mod", eSimple)
    Value = .ReadProperty("Val", " ")
  End With 'PropBag

End Sub

Private Sub UserControl_Resize()

  'position and size the lines
  
  Dim Hi   As Long
  Dim Wid  As Long
  Dim HMin As Long
  Dim WMin As Long
  Dim HMax As Long
  Dim WMax As Long
  Dim WMid As Long
  Dim HMid As Long

  Hi = UserControl.Height
  Wid = UserControl.Width
  HMin = Hi * 0.05
  WMin = Wid * 0.05
  HMax = Hi - HMin
  WMax = Wid - WMin
  WMid = Wid / 2
  HMid = Hi / 2
  'TopLine
  SetLine 0, WMin, HMin, WMid, HMin
  SetLine 1, WMid, HMin, WMax, HMin
  'Right
  SetLine 2, WMax, HMin, WMax, HMid
  SetLine 3, WMax, HMid, WMax, HMax
  'Bottom
  SetLine 4, WMax, HMax, WMid, HMax
  SetLine 5, WMid, HMax, WMin, HMax
  'left
  SetLine 6, WMin, HMax, WMin, HMid
  SetLine 7, WMin, HMid, WMin, HMin
  '12-6
  SetLine 8, WMid, HMin, WMid, HMid
  SetLine 9, WMid, HMid, WMid, HMax
  '9-3
  SetLine 10, WMin, HMid, WMid, HMid
  SetLine 11, WMid, HMid, WMax, HMid
  '11-4
  SetLine 12, WMin, HMin, WMid, HMid
  SetLine 13, WMid, HMid, WMax, HMax
  '2-7
  SetLine 14, WMax, HMin, WMid, HMid
  SetLine 15, WMid, HMid, WMin, HMax
  '9-12
  SetLine 16, WMin, HMid, WMid, HMin
  '12-3
  SetLine 17, WMid, HMin, WMax, HMid
  '3-6
  SetLine 18, WMax, HMid, WMid, HMax
  '6-9
  SetLine 19, WMid, HMax, WMin, HMid

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  With PropBag
    .WriteProperty "FCol", ForeColor, vbGreen
    .WriteProperty "WCol", WireColor, &HC000&
    .WriteProperty "BCol", BackColor, vbBlack
    .WriteProperty "Val", Value, " "
    .WriteProperty "OnWid", ONWidth, 3
    .WriteProperty "Shad", OFFShadow, True
    .WriteProperty "Mod", Mode, eSimple
  End With 'PropBag

End Sub

Public Property Get Value() As String

  Value = m_value

End Property

Public Property Let Value(ByVal NewValue As String)

  m_value = UCase$(NewValue)
  LiteUp
  PropertyChanged "Val"

End Property

Public Property Get WireColor() As OLE_COLOR

  WireColor = m_WColour

End Property

Public Property Let WireColor(ByVal NewCol As OLE_COLOR)

  m_WColour = NewCol
  PropertyChanged "WCol"
  DoOnColour

End Property

':)Code Fixer V2.8.9 (19/01/2005 3:48:18 PM) 55 + 347 = 402 Lines Thanks Ulli for inspiration and lots of code.

