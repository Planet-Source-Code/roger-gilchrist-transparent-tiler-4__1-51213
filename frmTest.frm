VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmTest 
   BackColor       =   &H80000007&
   Caption         =   "Transparant Tiler (Thanks for inspiration Jim, Carle PV & Robert Rayment for API help  & redbird77 for Ambiant Light)"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleMode       =   0  'User
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsclblInterpolation 
      Height          =   255
      Left            =   4440
      Max             =   100
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fraDrawWith 
      BackColor       =   &H80000007&
      Caption         =   "Draw With"
      ForeColor       =   &H80000004&
      Height          =   1095
      Left            =   5880
      TabIndex        =   31
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox picCFXPBugFixfrmTest 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   845
         Left            =   100
         ScaleHeight     =   840
         ScaleWidth      =   1260
         TabIndex        =   32
         Top             =   175
         Width           =   1255
         Begin VB.OptionButton optDrawWith 
            BackColor       =   &H80000007&
            Caption         =   "API"
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   1
            Left            =   20
            TabIndex        =   34
            Top             =   280
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDrawWith 
            BackColor       =   &H80000007&
            Caption         =   "VB"
            ForeColor       =   &H80000004&
            Height          =   195
            Index           =   0
            Left            =   20
            TabIndex        =   33
            Top             =   40
            Width           =   975
         End
         Begin VB.Label lblTime 
            BackColor       =   &H80000007&
            Caption         =   "Time :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   20
            TabIndex        =   35
            Top             =   520
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picAmbiantColour 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "Hide Comment"
      Height          =   255
      Left            =   10680
      TabIndex        =   28
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random Tile"
      Height          =   495
      Left            =   9600
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox chkLightAngle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Light Angle >"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2340
      TabIndex        =   22
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   19
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkLightAngle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdBevel 
      Caption         =   "Do Tiles (Refresh)"
      Height          =   435
      Index           =   2
      Left            =   7680
      TabIndex        =   14
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox chkSoftEdges 
      BackColor       =   &H00000000&
      Caption         =   "Hard Edges "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      ToolTipText     =   "Use  high Bevel Levels with Soft && Low with Hard Edges"
      Top             =   765
      Width           =   1215
   End
   Begin VB.CheckBox chkSquares 
      BackColor       =   &H00000000&
      Caption         =   "Squares "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1935
      TabIndex        =   12
      ToolTipText     =   "Ignores High value. Bottom row may be incomplete"
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox picDefault 
      Height          =   1455
      Left            =   6840
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   4995
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   255
      Index           =   3
      Left            =   0
      Max             =   200
      Min             =   1
      TabIndex        =   9
      Top             =   720
      Value           =   3
      Width           =   495
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   255
      Index           =   2
      Left            =   0
      Max             =   255
      Min             =   1
      TabIndex        =   7
      Top             =   510
      Value           =   64
      Width           =   495
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   255
      Index           =   1
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   6
      Top             =   255
      Value           =   1
      Width           =   495
   End
   Begin VB.HScrollBar hscSettings 
      Height          =   255
      Index           =   0
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   0
      Value           =   1
      Width           =   495
   End
   Begin VB.CommandButton cmdBevel 
      Caption         =   "Do Tiles (Cumulative)"
      Height          =   495
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      ToolTipText     =   "Apply setting again"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "Make transparent beveled frame"
      Height          =   495
      Left            =   10680
      TabIndex        =   1
      ToolTipText     =   "Jim K's idea"
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   0
      Top             =   1680
      Width           =   5880
   End
   Begin MSComDlg.CommonDialog cdlPic 
      Left            =   5880
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblInterpolation 
      BackColor       =   &H00404040&
      Caption         =   "Intensity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   37
      ToolTipText     =   "Values below .5 generally look better"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00000000&
      Caption         =   " Ambiant Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   30
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblhidden 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTest.frx":FDC5
      Height          =   1095
      Left            =   6840
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblIllusion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   5655
      Left            =   6360
      TabIndex        =   27
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label lblCurCode 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   1200
      Width           =   12735
   End
   Begin VB.Label lblLightValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   360
      Width           =   135
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 Bevel Width"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "64 Transparency"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 High"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   255
      Width           =   1455
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 Wide"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Save"
         Index           =   1
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "Ori&ginal"
         Index           =   3
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Blank picutre"
         Index           =   4
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This demo takes advantage of a weakness in VB;
'it doesn't notice if you have Private routines with the same name in different modules,
'so both modules contain the functions LightEffect and SafeRGBValue.
'
'This is a potential source of bugs in code, if the 2 routines are not exact copies of each other.
'
'In this demo they are exactly the same,
'I could also have renamed them by adding 'VB' or 'API' to end of names
'as I did with other routines (which are not identical but optimised (I think) for VB or API).
'
'Alternatively a single copy of the routines declared Public
'(in a 3rd module or either of the exiting modules) would work but would make the modules less portable.
'
'This Demo form is fairly crude but allows you to test the modules
'
Public Enum DMode
  VBPure
  APIPure
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private VBPure, APIPure
#End If
Private BevelDrawMode       As DMode
Private TileHigh            As Long
Private TileWidth           As Long
Private lngTrans            As Long
Private lngBwidth           As Long
Private bSquare             As Boolean
Private bHard               As Boolean
Private lngLightAng         As Long
Private sngIFactor As Single
Private Const strParam      As String = "BevelTiles PicB, lngTilesWide , [lngTilesHigh], [LngTransparency], [lngBevelWidth], [lngLightAngle], [bSquares], [bHardEdge], [lngAmbiantLight],[sngIntensityFactor]"

Private Sub chkLightAngle_Click(Index As Integer)

  chkLightAngle(lngLightAng).Value = vbUnchecked
  lngLightAng = Index
  lblLightValue.Caption = lngLightAng

End Sub

Private Sub chkSoftEdges_Click()

  bHard = chkSoftEdges.Value = 1

End Sub

Private Sub chkSquares_Click()

  bSquare = chkSquares.Value = 1

End Sub

Private Sub cmdBevel_Click(Index As Integer)

  DoBevel Index

End Sub

Private Sub cmdComment_Click()

  Select Case cmdComment.Caption
   Case "Hide Comment"
    cmdComment.Caption = "Show Comment"
    lblIllusion.Visible = False
   Case "Show Comment"
    cmdComment.Caption = "Hide Comment"
    lblIllusion.Visible = True
  End Select

End Sub

Private Sub cmdFrame_Click()

  Reset
  If BevelDrawMode = VBPure Then
    BevelFrameVB picDisplay, cdlPic.FileName, picAmbiantColour.BackColor
   Else
    BevelFrameAPI picDisplay, cdlPic.FileName, picAmbiantColour.BackColor
  End If

End Sub

Private Sub cmdRandom_Click()

  hscSettings(0).Value = Int(Rnd * 30) + 1
  hscSettings(1).Value = Int(Rnd * 30) + 1
  hscSettings(2).Value = Int(Rnd * 125) + 1
  hscSettings(3).Value = Int(Rnd * 95) + 3
  chkSquares.Value = IIf(Rnd > 0.5, 1, 0)
  chkSoftEdges.Value = IIf(Rnd > 0.5, 1, 0)
  picAmbiantColour.BackColor = CLng(Rnd * vbWhite)
  hsclblInterpolation.Value = Rnd(100)
  chkLightAngle.Item(Int(Rnd * 8)).Value = vbChecked
  DoEvents
  DoBevel 2

End Sub

Private Sub DoBevel(ByVal intIndex As Integer)


  Dim sngStartTime As Single

  lblCurCode.Caption = strParam & vbNewLine & "BevelTiles pic1 , " & TileWidth & " , " & TileHigh & " , " & lngTrans & " , " & lngBwidth & " , " & lngLightAng & " , " & bSquare & " , " & bHard & " ," & picAmbiantColour.BackColor & ", " & sngIFactor
  If intIndex = 2 Then
    Reset
  End If
  sngStartTime = Timer
  If BevelDrawMode = VBPure Then
    BevelTilesVB picDisplay, TileWidth, TileHigh, lngTrans, lngBwidth, lngLightAng, bSquare, bHard, picAmbiantColour.BackColor, sngIFactor
   Else
    BevelTilesAPI picDisplay, TileWidth, TileHigh, lngTrans, lngBwidth, lngLightAng, bSquare, bHard, picAmbiantColour.BackColor, sngIFactor
  End If
  lblTime.Caption = "Time:" & Timer - sngStartTime

End Sub

Private Sub DoClear()

  picDisplay.Picture = picDefault.Picture
  cdlPic.FileName = ""

End Sub

Private Sub DoScroll(ByVal intIndex As Integer)


  Select Case intIndex
   Case 0
    TileWidth = hscSettings(0).Value
    lblSettings(0).Caption = TileWidth & " Wide"
   Case 1
    TileHigh = hscSettings(1).Value
    lblSettings(1).Caption = TileHigh & " High"
   Case 2
    lngTrans = hscSettings(2).Value
    lblSettings(2).Caption = lngTrans & " Transparency"
   Case 3
    lngBwidth = hscSettings(3).Value
    lblSettings(3).Caption = lngBwidth & " Bevel Width"
  End Select

End Sub

Private Sub Form_Load()

  Dim I As Long

  lblIllusion.Caption = "The Tiling effect is an optical illusion based on your brain's natural bias/expectation about light." & vbNewLine & _
   "Your brain is hard - wired for light from above , so any shape ( the angled joints also contribute to the illusion)" & "which has light and dark edges is seen as raised if the upper edge is lighter and indented if it is darker." & vbNewLine & _
   "In this program light angle 1 almost always appears raised and 5 indented." & vbNewLine & "The other settings are more interesting. Your brain has no hard wiring about left/right because they change as you move about, so it uses various cues in the environment to decide what it is seeing." & "The most important cue is ambient light levels. The illusion depends on your real world environment; if it is " & "brighter to the left of your screen then 7 is raised and 3 indented but if it is darker " & "then  7 is indented and 3 raised. The corners (0, 2, 4 && 6) are more complex as they combine the up/down bias with envornmental levels, the greater the difference the stronger the illusion. " & "Indeed the presence of this light coloured comment to the right of the picture may influence your perception," & "Click 'Hide Comment' to see if this is happening on your screen." & vbNewLine & _
   "You can also use your desklamp to demostrate this by moving it around the screen." & vbNewLine & "This also means that the image you are tiling may interfer with the illusion; if it contains lighting cues (shadows, relativly brighter areas) which contradict the effect of the lighting angle you select." & vbNewLine & _
   "In fact you can override the up/down effect by putting the lamp at the bottom of the screen but it is not quite as convincing, because of the strong up/down bias." & vbNewLine & _
   "Note 1 To produce better effects for the vertical and horizontal angles the code shifts them slightly clock-wise of true; so 3 has a slight down bias and 7 an up bias." & vbNewLine & _
   "Note 2 if you load your own picture this comment will be hidden auomatically." & vbNewLine & vbNewLine & "Thanks to Carle P.V. and and Robert Rayment who pointed out the dumb mistake that stopped API working in earlier versions. API is 40-60% faster, VB method is only for Demo purposes. Thanks redbird77 for solving the Ambiant light problem"

  Randomize Timer
  lblCurCode.Caption = strParam
  lngLightAng = 1
  For I = 0 To 3
    DoScroll I
  Next I
  Show
  DoClear
  DoEvents
  sngIFactor = 0.1
  hsclblInterpolation.Value = sngIFactor * 100
  RandomLite

End Sub

Private Sub Form_Resize()

  lblCurCode.Left = Me.ScaleLeft
  lblCurCode.Width = Me.ScaleWidth
  picDisplay.Top = lblCurCode.Top + lblCurCode.Height

End Sub

Private Sub hsclblInterpolation_Change()
sngIFactor = hsclblInterpolation.Value / 100
lblInterpolation = "Intensity: " & sngIFactor
End Sub

Private Sub hscSettings_Change(Index As Integer)

  DoScroll Index

End Sub

Private Sub mnufileOpt_Click(Index As Integer)

  Select Case Index
   Case 0 'open
    If cmdComment.Caption = "Hide Comment" Then
      cmdComment_Click
    End If
    With cdlPic
      .Filter = "Picture files (*.gif*.bmp*.jpg)|*.bmp;*.gif;*.jpg;*.jpeg"
      .FilterIndex = 1
      .ShowOpen
      If Len(.FileName) Then
        picDisplay.Picture = LoadPicture(.FileName)
      End If
    End With 'cdlPic
   Case 1 'save
    With cdlPic
      .Filter = "Picture files (*.bmp)|*.bmp"
      .FilterIndex = 1
' you may want to replace but this warns you
      .Flags = cdlOFNOverwritePrompt
      .ShowSave
      If Len(.FileName) Then
'NOTE The code modifies the Image not the Picture
'so without next line you just get a copy of the original
        picDisplay.Picture = picDisplay.Image
        SavePicture picDisplay.Picture, .FileName
      End If
    End With 'cdlPic
   Case 3 'original
    Reset
   Case 4 'blank
    Set picDisplay = Nothing
   Case 6
    Unload Me
  End Select

End Sub

Private Sub optDrawWith_Click(Index As Integer)

  BevelDrawMode = Index

End Sub

Private Sub picAmbiantColour_Click()

'This is here to help anyone who can solve the colourising problem
'As is' the code only allows white light
' MsgBox "Not Yet Available" & vbNewLine & "If you know how to do this contact me.", vbInformation, "Change lighting colour"
'Exit Sub

  With cdlPic
    .Color = picAmbiantColour.BackColor
    .ShowColor
    picAmbiantColour.BackColor = .Color
  End With

End Sub

Private Sub RandomLite()

'this is used in the start up so that whatever settings it picks it doesn't take too long
''
'See cmdRandom_Click for the heavy version

  hscSettings(0).Value = Int(Rnd * 4) + 3
  hscSettings(1).Value = Int(Rnd * 4) + 2
  chkSquares.Value = IIf(Rnd > 0.5, 1, 0)
  chkSoftEdges.Value = IIf(Rnd > 0.5, 1, 0)
  If chkSoftEdges.Value Then ' this is just to ensure it is good looking
    hscSettings(2).Value = Int(Rnd * 25) + 40
    hscSettings(3).Value = Int(Rnd * 5) + 3
   Else
    hscSettings(2).Value = Int(Rnd * 25) + 80
    hscSettings(3).Value = Int(Rnd * 5) + 3
  End If
  picAmbiantColour.BackColor = CLng(Rnd * vbWhite)
  hsclblInterpolation.Value = Rnd(90) + 10
  chkLightAngle.Item(Int(Rnd * 8)).Value = vbChecked
  DoEvents
  DoBevel 2

End Sub

Private Sub Reset()

  If Len(cdlPic.FileName) Then
    picDisplay.Picture = LoadPicture(cdlPic.FileName)
   Else
    picDisplay.Picture = picDefault.Picture
  End If
  DoEvents

End Sub


':)Roja's VB Code Fixer V1.1.91 (24/01/2004 4:33:07 PM) 32 + 248 = 280 Lines Thanks Ulli for inspiration and lots of code.

