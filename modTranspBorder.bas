Attribute VB_Name = "modTranspBorder"
'Based on code By Jim K  vb6@c2i.net 'Transparent bevels'
'Posted at PSC http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=50750&lngWId=1
'------------------------------------------------------
' Terms     : This code can be used for free
' of use    : If you use it, you're encouraged to give
'           : credits to the author of the code.
'------------------------------------------------------
'here's the credit Jim!
'Modifications by roger gilchrist rojagilkrist@hotmail.com
'Same terms of use as Jim K
'have converted it to a seperate module
'BevelTiles: added many options to allow a large range of possible tiles
'BevelFrame: simple conversion with no further additions other than that it works out the picture size for itself
'
'NB I use Long rather than Integer even when I know the value (ie RGB colour values)
'will be well within the Integer or even Byte range. Despite it's own documentation
'VB on 32-bit systems handles Longs slightly faster (very small but real amount) than Integers
'
'REQUEST
'using PSet to do this is veeeeeeeeeerrrry slow. Anyone have the code to do it with API?
'For Some reason GetPixel and SetPixel didn't work properly on my system
'(Almost certainly cause I've messed up some setting of the picturebox)
'only change is to use ScaleWidth/Height to make his code
'control independant and to shorten the caption to filename only
Option Explicit

Public Sub BevelFrame(Ctrl As PictureBox, _
                      ByVal strCaption As String)

'Based on code By Jim K  vb6@c2i.net 'Transparent bevels'
'Frame
'3 pixels thick

  Dim I                As Long
  Dim lngPrevScaleMode As Long

  With Ctrl
    .Parent.MousePointer = vbHourglass
    lngPrevScaleMode = .ScaleMode
    .ScaleMode = 3
'Outer bevel
  End With 'Ctrl
  With Ctrl
    For I = 0 To 2
      DrawBevel Ctrl, I, I, .ScaleWidth - I, .ScaleHeight - I, 64, 1
    Next I
'Inner bevel
    For I = 4 To 6
      DrawBevel Ctrl, I, 20 + I, .ScaleWidth - I, .ScaleHeight - I, -64, 1
    Next I
  End With
  With Ctrl
    .CurrentX = 6
    .CurrentY = 3
    .FontSize = 10
    .ForeColor = vbWhite
  End With 'ctrl
  If LenB(strCaption) Then
    Ctrl.Print "[" & Mid$(strCaption, InStrRev(strCaption, "\") + 1) & "]"
   Else
    Ctrl.Print "[Default Picture]"
  End If
  Ctrl.ScaleMode = lngPrevScaleMode
  Ctrl.Parent.MousePointer = vbDefault

End Sub

Public Sub BevelTiles(picB As PictureBox, _
                      ByVal lngTilesWide As Long, _
                      Optional ByVal lngTilesHigh As Long = -1, _
                      Optional ByVal LngTransparency As Long = 64, _
                      Optional ByVal lngBevelWidth As Long = 3, _
                      Optional ByVal lngLightAngle As Long = 0, _
                      Optional ByVal bSquares As Boolean = False, _
                      Optional ByVal bHardEdge As Boolean = True)


  Dim LT               As Long
  Dim WH               As Long
  Dim lngPrevScaleMode As Long
  Dim I                As Long
  Dim J                As Long
  Dim K                As Long
  Dim TileW            As Single
  Dim TileH            As Single

  With picB
    .Parent.MousePointer = vbHourglass 'let user know something is happening
    lngPrevScaleMode = .ScaleMode      'preserve value (other bits of a program might assume/need something else
    .ScaleMode = 3                     'settings this code needs
  End With 'picB
  If lngTilesHigh = -1 Then                           ' if only required settings used the
    lngTilesHigh = lngTilesWide                       ' set up for square tiling
    bSquares = True
  End If
  If bSquares Then                                    'if square tiling
    TileW = picB.ScaleWidth / lngTilesWide + 1        'adding 1 extra column to  cope with unsquarable area
    TileH = TileW                                     'match height to width
    lngTilesHigh = 0
    Do                                                ' calculate the number of tiles high to use
      lngTilesHigh = lngTilesHigh + 1
    Loop Until lngTilesHigh * TileH > picB.ScaleHeight
   Else                                               '-----------------------------------
    TileW = picB.ScaleWidth / lngTilesWide            'just set the height and width based on simple math
    TileH = picB.ScaleHeight / lngTilesHigh
  End If
  
  For I = 0 To lngTilesWide - 1
    For J = 0 To lngTilesHigh - 1
      LT = TileW * I
      WH = TileH * J
      For K = 0 To lngBevelWidth - 1
        DrawBevel picB, LT + K, WH + K, LT + TileW - K - 1, WH + TileH - K - 1, IIf(bHardEdge, LngTransparency, LngTransparency / (K + 1)), lngLightAngle
      Next K
      picB.Refresh                   ' remove for higher speed or move above a 'Next' for more or less feedback
    Next J
  Next I
  With picB                           'restore control settings
  .ScaleMode = lngPrevScaleMode
  .Parent.MousePointer = vbDefault
  End With

End Sub

Private Sub DrawBevel(picB As PictureBox, _
                      ByVal X1 As Long, _
                      ByVal Y1 As Long, _
                      ByVal X2 As Long, _
                      ByVal Y2 As Long, _
                      HSVal As Long, _
                      lngLightAngle As Long)

'major rewrite of jim's code to allow light angles
' and to test and set values outside the loops where possible

  Dim I    As Long
  Dim Tval As Long
  Dim Bval As Long
  Dim Lval As Long
  Dim Rval As Long

  Tval = LightEffect(lngLightAngle, 1, HSVal)
  Rval = LightEffect(lngLightAngle, 2, HSVal)
  Bval = LightEffect(lngLightAngle, 3, HSVal)
  Lval = LightEffect(lngLightAngle, 4, HSVal)
  With picB
    For I = X1 To X2
      picB.PSet (I, Y1), FindColor(.Point(I, Y1), Tval)
      picB.PSet (I, Y2), FindColor(.Point(I, Y2), Bval)
    Next I
    For I = Y1 + 1 To Y2 - 1
      picB.PSet (X1, I), FindColor(.Point(X1, I), Lval)
      picB.PSet (X2, I), FindColor(.Point(X2, I), Rval)
    Next I
  End With

End Sub

Private Function FindColor(ByVal Pixel As Long, _
                           ByVal HSVal As Long) As Long

'Finds the current pixel's RGB Value and adds
'the HSVal that is set to either make a
'brighter or darker version of the color
'-----------------------------------------
'simplified Jim's code by using the vbColour variables
'which make the code less mysterious
'  R = (Pixel And vbRed) + HSVal 'Red
'  G = (Pixel And vbGreen) \ 256 + HSVal 'Green
'  B = (Pixel And vbBlue) \ 65536 + HSVal 'Blue
' and moved it in-line to cut out the extra variables
  
  FindColor = RGB(SafeRGBValue((Pixel And vbRed) + HSVal), SafeRGBValue((Pixel And vbGreen) \ 256 + HSVal), SafeRGBValue((Pixel And vbBlue) \ 65536 + HSVal))

End Function

Private Function LightEffect(ByVal lngLightAngle As Long, _
                             ByVal lngElement As Long, _
                             lngBasicLight As Long) As Long

  Dim Side1 As Long
  Dim Side2 As Long

'All new
'set the +ve/-ve and (sometimes) streahgth of the lighting value
'Depends on which side of frame is drawing
'   1
' 4   2  lngElement
'   3
'and 'direction' light is coming from
' 0 1 2
' 7   3 lngLightAngle
' 6 5 4
'
'the Side1 & Side2 are arbitary values used with lightangles 1,3,5,7 which
'allow a visible difference between facing bevels at sides of tiles
'in 'reality' they should be the same, thes number have the effect of making angles 1,3,5,7
'slightly clockwise (if Side1<Side2) or anti-clockwise(Side1>side2) of the 90 degrees position.
'you might change them but don't make them identical (the resulting frame is pretty ugly.)
'
'NB this all assumes you are viewing the tiles as raised. For most people there is an optical
'illusion that the tiles are sunken if light is from below (4,5,6) and raised if it is from above(0,1,2).
'because Side1 & Side2 variables move the angle slightly clockwise angle 3 is below and 7 is above.
'This is just your brain filling in details (in real life light normal comes from above,
'so light on bottom of an object and dark above usually means indented)
'If you see the tiles as raised at one angle then the clock mirror setting will usually seem sunken.
  Side1 = 2
  Side2 = 3
  Select Case lngElement
   Case 1, 3 'Top' bottom
    Select Case lngLightAngle
     Case 0, 1, 2
      LightEffect = IIf(lngElement = 1, lngBasicLight, -lngBasicLight)
     Case 4, 5, 6
      LightEffect = IIf(lngElement = 1, -lngBasicLight, lngBasicLight)
     Case 3
      LightEffect = -lngBasicLight / IIf(lngElement = 1, Side1, Side2)
     Case 7
      LightEffect = -lngBasicLight / IIf(lngElement = 1, Side2, Side1)
    End Select
   Case 2, 4 'right ' left
    Select Case lngLightAngle
     Case 0, 6, 7
      LightEffect = IIf(lngElement = 2, -lngBasicLight, lngBasicLight)
     Case 2, 3, 4
      LightEffect = IIf(lngElement = 2, lngBasicLight, -lngBasicLight)
     Case 1
      LightEffect = -lngBasicLight / IIf(lngElement = 2, Side2, Side1)
     Case 5
      LightEffect = -lngBasicLight / IIf(lngElement = 2, Side1, Side2)
    End Select
  End Select

End Function

Private Function SafeRGBValue(lngVal As Long) As Long

'modification of a more general KeepInBounds (Min,Val,Max) routine
'as the min and max are always the same for this prog

  SafeRGBValue = lngVal
  If SafeRGBValue < 0 Then
    SafeRGBValue = 0
   ElseIf SafeRGBValue > 255 Then
    SafeRGBValue = 255
  End If

End Function

':)Roja's VB Code Fixer V1.1.82 (5/01/2004 9:44:52 AM) 1 + 237 = 238 Lines Thanks Ulli for inspiration and lots of code.

