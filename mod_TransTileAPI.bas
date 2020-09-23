Attribute VB_Name = "mod_TransTileAPI"
Option Explicit
'USE THIS MODULE IN YOUR CODE
'THE VB VERSION IS FOR DEMO PURPOSES ONLY
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
' Enormous heaps of Thanks to Carles P.V. and Robert Rayment
'the API problem in earlier versions was a classic dumb cut'n'paste error
'I wasn't sending the right variables into the colour setting parameter
'everything got the top colour modifier.
'Heaps more thanks to redbird77 for the ambiant light solution
'IF YOU ARE LOOKING AT THE 2 VERSIONS YOU WILL NOTICE THAT SOME OF THE PROCEDURES ARE DUPLICATED
'THIS IS DELIBERATE YOU ONLY NEED TO TAKE THE MODULE YOU WANT.
'VB IS NOT UPSET BY PROCEDURES WITH THE SAME NAME AS LONG AS THEY ARE PRIVATE AND IN DIFFERENT MODULES
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long

Public Sub BevelFrameAPI(Ctrl As PictureBox, _
                         ByVal strCaption As String, _
                         Optional lngAmbiantLight As Long = vbWhite, _
                         Optional sngIntensityFactor As Single = 0.1)

'only change is to use ScaleWidth/Height to make his code
''Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'control independant and to shorten the caption to filename only
'Based on code By Jim K  vb6@c2i.net 'Transparent bevels'

  Dim I                As Long
  Dim lngPrevScaleMode As Long

  With Ctrl
    .Parent.MousePointer = vbHourglass
    lngPrevScaleMode = .ScaleMode
    .ScaleMode = vbPixels
  End With 'Ctrl
  With Ctrl
'Outer bevel
    For I = 0 To 2
      DrawBevelAPI Ctrl, I, I, .ScaleWidth - I, .ScaleHeight - I, 64, 1, lngAmbiantLight, sngIntensityFactor
    Next I
'Inner bevel
    For I = 4 To 6
      DrawBevelAPI Ctrl, I, 20 + I, .ScaleWidth - I, .ScaleHeight - I, -64, 1, lngAmbiantLight, sngIntensityFactor
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

Public Sub BevelTilesAPI(picB As PictureBox, _
                         ByVal lngTilesWide As Long, _
                         Optional ByVal lngTilesHigh As Long = -1, _
                         Optional ByVal LngTransparency As Long = 64, _
                         Optional ByVal lngBevelWidth As Long = 3, _
                         Optional ByVal lngLightAngle As Long = 0, _
                         Optional ByVal bSquares As Boolean = False, _
                         Optional ByVal bHardEdge As Boolean = True, _
                         Optional lngAmbiantLight As Long = vbWhite, _
                        Optional sngIntensityFactor As Single = 0.1)

'PARAMETERS
'picB           -PictureBox you wish to tile
'lngTilesWide   -no of tiles across
'lngTilesHigh   -no of tile down; if set to -1then bSquares is set to True auomatically
'LngTransparency- degree of transparance 0=clear (pretty useless) to 255 solid (White at present)
'lngBevelWidth  -no of lines in frame
'lngLightAngle  -lighting angle
'bSquares       -if True then lngTilesHigh is ignored what ever value is set
'bHardEdge      -True each line in frame edge is same colour. False and the LngTransparency value decrease with each level of lngBevelWidth
'lngAmbiantLight - allows you to specify what colour to draw tiles in Thanks redbird77

  Dim LT               As Long
  Dim WH               As Long
  Dim lngPrevScaleMode As Long
  Dim I                As Long
  Dim J                As Long
  Dim K                As Single
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
      For K = 0 To lngBevelWidth - 1 Step 0.5
        DrawBevelAPI picB, LT + K, WH + K, LT + TileW - K - 1, WH + TileH - K - 1, IIf(bHardEdge, LngTransparency, LngTransparency / (K + 1)), lngLightAngle, lngAmbiantLight, sngIntensityFactor
      Next K
      picB.Refresh                   ' remove for higher speed or move above a 'Next' for more or less feedback
    Next J
  Next I
  picB.Picture = picB.Image
  With picB                           'restore control settings
    .ScaleMode = lngPrevScaleMode
    .Parent.MousePointer = vbDefault
  End With

End Sub

Private Sub DrawBevelAPI(picB As PictureBox, _
                         ByVal X1 As Long, _
                         ByVal Y1 As Long, _
                         ByVal X2 As Long, _
                         ByVal Y2 As Long, _
                         ByVal HSVal As Long, _
                         ByVal lngLightAngle As Long, _
                         ByVal lngAmbiantLight As Long, _
                         ByVal sngIntensityFactor As Single)

'major rewrite of jim's code to allow light angles
' and to test and set values outside the loops where possible

  Dim I    As Long

  With picB
    For I = X1 To X2
      SetPixel .hdc, I, Y1, FindColorAPI(GetPixel(.hdc, I, Y1), LightEffect(lngLightAngle, 1, HSVal), lngAmbiantLight, sngIntensityFactor)
      SetPixel .hdc, I, Y2, FindColorAPI(GetPixel(.hdc, I, Y2), LightEffect(lngLightAngle, 3, HSVal), lngAmbiantLight, sngIntensityFactor)
    Next I
    For I = Y1 To Y2
      SetPixel .hdc, X1, I, FindColorAPI(GetPixel(.hdc, X1, I), LightEffect(lngLightAngle, 4, HSVal), lngAmbiantLight, sngIntensityFactor)
      SetPixel .hdc, X2, I, FindColorAPI(GetPixel(.hdc, X2, I), LightEffect(lngLightAngle, 2, HSVal), lngAmbiantLight, sngIntensityFactor)
    Next I
  End With

End Sub

Private Function FindColorAPI(ByVal lColor As Long, _
                              ByVal HSVal As Long, _
                              ByVal lLightColour As Long, _
                              ByVal sngIntensityFactor As Single) As Long

'Thanks to redbird77 for this code AND even commented it!!
'modified so that test for valid values  r, g & b (0-255)
'occurs in redbird77's Interpolate function
'added the sngIntensityFactor parameter to vary strength of Interpolate effect
'From redbird77's notes
' 0.5 will give you the average of the Ambiant light and original pixel
' colors, so you'll probably get best results with a position between 0 and 0.5
'Demo lets you go all the way to 1
'NOTE 1 lr, lg & lb don't need to be tested as they are safe as set from demo
'if you are providing them from code you may have to test them
'NOTE 2 for more speed you could dump the Dims and plug the equation straight into the RGB( ) code
'but that is very hard to read

  Dim r  As Long
  Dim g  As Long
  Dim b  As Long
  Dim lr As Long
  Dim lg As Long
  Dim lb As Long

' get parts of ambient
  lr = (lLightColour And vbRed)
  lg = (lLightColour And vbGreen) \ 256
  lb = (lLightColour And vbBlue) \ 65536
' get new color and adjust lightness
  r = (lColor And vbRed) + HSVal
  g = (lColor And vbGreen) \ 256 + HSVal
  b = (lColor And vbBlue) \ 65536 + HSVal
  FindColorAPI = RGB(Interpolate(r, lr, sngIntensityFactor), Interpolate(g, lg, sngIntensityFactor), Interpolate(b, lb, sngIntensityFactor))

End Function

Private Function Interpolate(ByVal OrigColour As Long, _
                             ByVal AmbiantColour As Long, _
                             ByVal sngIntensityFactor As Single) As Byte

'Thanks to redbird77 for this code
'changed names of parameters to help me work it out
'changed OrigColour & AmbiantColour to Long as the HSVal might make OrigColour exceed limits
'and AmbiantColour could be set from code and suspect
'SafeRGBValue takes care of that
'Finds a value somewhere between OrigColour and AmbiantColour
'where 0 is OrigColour and 1 is AmbiantColour
  Interpolate = SafeRGBValue(OrigColour) * (1 - sngIntensityFactor) + SafeRGBValue(AmbiantColour) * sngIntensityFactor

End Function

Private Function LightEffect(ByVal lngLightAngle As Long, _
                             ByVal lngElement As Long, _
                             ByVal lngBasicLight As Long) As Long

  Dim Side1 As Single
  Dim Side2 As Single

'lngBasicLight = Abs(lngBasicLight)
'All new
'set the +ve/-ve and (sometimes) strength of the lighting value
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
'
'for greater speed hard code the values into the code rather than assigning them to variables
  Side1 = 2
  Side2 = 3
''
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
  LightEffect = LightEffect

End Function

Private Function SafeRGBValue(ByVal lngVal As Long) As Long

'modification of a more general KeepInBounds (Min,Val,Max) routine
'as the min and max are always the same for this prog

  SafeRGBValue = lngVal
  If SafeRGBValue < 0 Then
    SafeRGBValue = 0
   ElseIf SafeRGBValue > 255 Then
    SafeRGBValue = 255
  End If

End Function

':)Roja's VB Code Fixer V1.1.91 (24/01/2004 4:33:09 PM) 30 + 236 = 266 Lines Thanks Ulli for inspiration and lots of code.

