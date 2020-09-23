VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gdipPathWarper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GdipReversePath Lib "gdiplus" (ByVal path As Long) As Long

' sources for warping and matrix math
' http://www.codeguru.com/cpp/g-m/gdi/gdi/article.php/c3657/
' http://www.java2s.com/Tutorial/VB/0040__Data-Type/Solveequationsusingmatrices.htm

' Note: The class requires two calls (any order) before anything can be rendered to a DC
' 1) Call to SetPathString to set the string to be displayed
' 2) Call to SetPathDest_Points or SetPathDest_Rect to set destination drawing area

' Last but not least. I included options to add rectangles & ellipses. Feel free
' to add code for Lines, Pies, Arcs, Curves, etc, etc.  The SetPathShape_Ellipse and
' SetPathShape_Rectangle functions can be used as templates.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' GDI+ functions
' ---====| GDI+ Rendering quality |====---
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMode As SmoothingMode) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolation As InterpolationMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As Long

' ---====| GDI+ Pens & Brushes |====---
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, Pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (ByRef pRect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal wrapMode As Long, ByRef lineGradient As Long) As Long

' ---====| GDI+ graphics object |====---
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long


' ---====| GDI+ Path functions |====---
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal Pen As Long, ByVal path As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal path As Long) As Long
Private Declare Function GdipClonePath Lib "gdiplus.dll" (ByVal path As Long, ByRef clonePath As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus.dll" (ByVal brushMode As gdiFillMode, ByRef path As Long) As Long
Private Declare Function GdipCreatePath2 Lib "gdiplus.dll" (ByVal pPointsPtr As Long, ByVal typesPtr As Long, ByVal Count As Long, ByVal fillMode As Long, ByRef path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus.dll" (ByVal path As Long) As Long
Private Declare Function GdipGetPathData Lib "gdiplus.dll" (ByVal path As Long, ByVal pDataPtr As Long) As Long
Private Declare Function GdipGetPathFillMode Lib "gdiplus.dll" (ByVal path As Long, ByRef fillMode As Long) As Long
Private Declare Function GdipGetPathPoints Lib "gdiplus.dll" (ByVal path As Long, ByRef pPointsPtr As Long, ByVal Count As Long) As Long
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus.dll" (ByVal path As Long, ByRef Bounds As RECTL, ByVal matrix As Long, ByVal Pen As Long) As Long
Private Declare Function GdipGetPointCount Lib "gdiplus.dll" (ByVal path As Long, ByRef Count As Long) As Long
Private Declare Function GdipResetPath Lib "gdiplus.dll" (ByVal path As Long) As Long
Private Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal path As Long, ByVal str As Long, ByVal Length As Long, ByVal family As Long, ByVal Style As Long, ByVal emSize As Single, layoutRect As RECTL, ByVal StringFormat As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal path As Long, ByVal addingPath As Long, ByVal pConnect As Long) As Boolean

' ---====| GDI+ String functions |====---
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long

Private Enum gdiFillMode    ' used for CreatePath
    FillModeAlternate = 0
    FillModeWinding = 1
End Enum
Public Enum gdipWarpModes   ' path warping options
    warpPerspective = 0
    warpBilinear = 1        ' default
    warpSkew = 2
End Enum
Private Enum gdipFontStyles ' used for AddPathString
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
Public Enum LinearGradientMode      ' gradient brush options
    LinearGradientModeHorizontal = 0
    LinearGradientModeVertical = 1
    LinearGradientModeForwardDiagonal = 2
    LinearGradientModeBackwardDiagonal = 3
End Enum

' ---====| graphics quality constants & enumerations |====---
Private Const QualityModeInvalid As Long = -1&
Private Const QualityModeDefault As Long = 0&
Private Const QualityModeLow As Long = 1&
Private Const QualityModeHigh As Long = 2&
Private Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone = QualityModeHigh + 1
    SmoothingModeAntiAlias8x4 = QualityModeHigh + 2
    SmoothingModeAntiAlias = SmoothingModeAntiAlias8x4
    SmoothingModeAntiAlias8x8 = QualityModeHigh + 3
End Enum
Private Enum InterpolationMode
   InterpolationModeInvalid = QualityModeInvalid
   InterpolationModeDefault = QualityModeDefault
   InterpolationModeLowQuality = QualityModeLow
   InterpolationModeHighQuality = QualityModeHigh
   InterpolationModeBilinear
   InterpolationModeBicubic
   InterpolationModeNearestNeighbor
   InterpolationModeHighQualityBilinear
   InterpolationModeHighQualityBicubic
End Enum
Private Enum PixelOffsetMode
   PixelOffsetModeInvalid = QualityModeInvalid
   PixelOffsetModeDefault = QualityModeDefault
   PixelOffsetModeHighSpeed = QualityModeLow
   PixelOffsetModeHighQuality = QualityModeHigh
   PixelOffsetModeNone    ' No pixel offset
   PixelOffsetModeHalf     ' Offset by -0.5 -0.5 for fast anti-alias perf
End Enum

' ---====| Enumerations |====---
Public Enum ePathOrder
    TopLeft_Clockwise = 0 ' left-top, right-top, right-bottom, left-bottom
    TopLeft_zPattern = 1  ' left-top, right-top, left-bottom, right-bottom
End Enum
Public Enum ePathPoint
    TopLeft = 0
    TopRight = 1
    BottomLeft = 2
    BottomRight = 3
End Enum
Private Enum StatusCodes
    statTextApplied = 1
    statBoundsApplied = 2
    statCanDraw = 3
    statRecalcWarp = 4
    statInvalid = 8
End Enum

' ---====| User-Defined Types |====---
Private Type PointF ' used by many GDI+ functions
   X As Single
   Y As Single
End Type
Private Type PathData   ' used to retrieve path points & types
   Count As Long
   Points As Long ' pointer to array of PointF
   Types As Long ' pointer to array of gdiPathTypes
End Type
Private Type RECTL  ' used by many GDI+ functions
   X As Long
   Y As Long
   mWidth As Long
   mHeight As Long
End Type
Private Type PenBrushData   ' cached path information
    Pen As Long         ' pen handle -- must be destroyed
    Brush As Long       ' brush handle -- must be destroyed
    PenColor As Long    ' pen/outline color
    color1 As Long      ' solid brush color or 1st gradient color
    color2 As Long      ' second gradient color
    GradDirection As LinearGradientMode
    GradBounds As RECTL ' latest warp bounds
End Type

' ---====| Class variables |====---
Private m_WarpFactors() As Single   ' calc factors for manual warping
Private m_Style As gdipWarpModes    ' current warp option
Private m_Path As Long              ' handle to a Path -- must be destroyed
Private m_SrcPts() As PointF        ' rect coordinates of non-warped path
Private m_DestPts() As PointF       ' coordinates of warped path
Private m_Status As StatusCodes     ' class-only status flags
Private m_FillObject As PenBrushData    ' GDI+ objects, pen/brush colors, etc

Public Function SetPathString(Text As String, Font As StdFont, _
                            Optional X As Long, Optional Y As Long, _
                            Optional hDC As Long, Optional ResetPath As Boolean = True) As Boolean

    ' if function returns zero; failed to create path from passed Font -- true type fonts only!
    
    ' Text is the formatted text, including vbCrLf characters if appropriate
    ' Font is the font to be used with Text. Use TrueType fonts
    ' X & Y are where in the path the string should start
    ' hDC if provided will draw the updated path to the passed DC
    
    Dim fRect As RECTL, fontFam As Long, fontStyle As Long
    Dim tPath As Long
    
    ' get font style from passed font object
    With Font
        fontStyle = (Abs(.Italic) * FontStyleItalic) Or (Abs(.Strikethrough) * FontStyleStrikeout) _
        Or (Abs(.Underline) * FontStyleUnderline) Or (Abs(.Bold) * FontStyleBold)
    End With

    ' create new path and add text as the path
    Call GdipCreateFontFamilyFromName(StrPtr(Font.name), 0, fontFam)
    If fontFam Then
        If GdipCreatePath(FillModeAlternate, tPath) = 0& Then
            fRect.X = X: fRect.Y = Y
            Call GdipAddPathStringI(tPath, StrPtr(Text), -1, fontFam, fontStyle, Font.Size, fRect, 0&)
            GdipDeleteFontFamily fontFam
            SetPathString = AppendPath(tPath, ResetPath, hDC)
        Else
            GdipDeleteFontFamily fontFam
        End If
    End If
End Function

Public Function SetPathShape_Rectangle(Left As Long, Top As Long, _
                            Right As Long, Bottom As Long, _
                            Optional hDC As Long, Optional ResetPath As Boolean = True) As Boolean

    ' if function returns false; failed to create path
    
    ' Left & Top are where the rectangle starts
    ' Right & Bottom are where it ends
    ' hDC if provided will draw the updated path to the passed DC
    
    Dim tPath As Long
        
    If GdipCreatePath(FillModeAlternate, tPath) = 0& Then
        GdipAddPathRectangleI tPath, Left, Top, Right - Left + 1, Bottom - Top + 1
        If AppendPath(tPath, ResetPath, 0&) Then
            SetPathDest_Rect Left, Top, Right, Bottom, hDC
            SetPathShape_Rectangle = True
        End If
    End If

End Function

Public Function SetPathShape_Ellipse(Left As Long, Top As Long, _
                            Right As Long, Bottom As Long, _
                            Optional hDC As Long, Optional ResetPath As Boolean = True) As Boolean

    ' if function returns false; failed to create path
    
    ' Left & Top are where the ellipsis starts
    ' Right & Bottom are where it ends
    ' hDC if provided will draw the updated path to the passed DC
    
    Dim tPath As Long
        
    If GdipCreatePath(FillModeAlternate, tPath) = 0& Then
        GdipAddPathEllipseI tPath, Left, Top, Right - Left + 1, Bottom - Top + 1
        If AppendPath(tPath, ResetPath, 0&) Then
            SetPathDest_Rect Left, Top, Right, Bottom, hDC
            SetPathShape_Ellipse = True
        End If
    End If

End Function

Public Sub SetPathDest_Points(Points() As Single, ByVal Order As ePathOrder, Optional ByVal hDC As Long)
    
'    // warp points are arranged in Z-configuration internally
'    //
'    //      0 ------- 1
'    //      |         |
'    //      |         |
'    //      2 ------- 3
'    //
    
    ' Points must be a 2D array: (0 to 1, 0 to 3)
    ' If Order = TopLeft_Clockwise then point order is
    '   (0,0)=TopLeft-X     (1,0)=TopLeft-Y
    '   (0,1)=TopRight-X    (1,1)=TopRight-Y
    '   (0,2)=BotRight-X    (1,2)=BotRight-Y
    '   (0,3)=BotLeft-X     (1,3)=BotLeft-Y
    ' if Order = TopLeft_zPattern then
    '   (0,0)=TopLeft-X     (1,0)=TopLeft-Y
    '   (0,1)=TopRight-X    (1,1)=TopRight-Y
    '   (0,2)=BotLeft-X     (1,2)=BotLeft-Y
    '   (0,3)=BotRight-X    (1,3)=BotRight-Y
    ' hDC if provided will draw the updated path to the passed DC
    On Error GoTo EH
    
    If UBound(Points, 2) = 3 And UBound(Points, 1) = 1 Then
        If LBound(Points, 2) = 0 And LBound(Points, 1) = 0 Then
            If Order = TopLeft_zPattern Then
                CopyMemory m_DestPts(0), Points(0, 0), 32&
            ElseIf Order = TopLeft_Clockwise Then
                CopyMemory m_DestPts(0), Points(0, 0), 16&
                CopyMemory m_DestPts(3), Points(2, 0), 8&
                CopyMemory m_DestPts(2), Points(3, 0), 8&
            Else
                Exit Sub
            End If
            If m_Style = warpSkew Then
                m_DestPts(3).X = m_DestPts(1).X + (m_DestPts(2).X - m_DestPts(0).X)
                m_DestPts(3).Y = m_DestPts(1).Y + (m_DestPts(2).Y - m_DestPts(0).Y)
            End If
            BuildWarpMatrix hDC
        End If
    End If
EH:
    If Err Then Err.Clear ' passed wrong dimension array
End Sub

Public Sub SetPathDest_Rect(rectLeft As Long, rectTop As Long, _
                        rectRight As Long, rectBottom As Long, Optional hDC As Long)

    ' // warp points are arranged in Z-configuration -- See SetPathDest_Points routine
    ' hDC if provided will draw the updated path to the passed DC
    
    m_DestPts(0).X = rectLeft: m_DestPts(0).Y = rectTop
    m_DestPts(1).X = rectRight: m_DestPts(1).Y = rectTop
    m_DestPts(2).X = rectLeft: m_DestPts(2).Y = rectBottom
    m_DestPts(3).X = rectRight: m_DestPts(3).Y = rectBottom
    BuildWarpMatrix hDC

End Sub

Public Function UpdateDestPoint(ByVal X As Single, ByVal Y As Single, ByVal whichPoint As ePathPoint, Optional hDC As Long) As Boolean
    ' update a single point in the warp bounds
    ' hDC if provided will draw the updated path to the passed DC
    
    ' when skewing, the warp shape must be a parallelogram. This requires calculating
    ' and also modifying the opposite corner point from the one being updated
    If (m_Status And statBoundsApplied) = statBoundsApplied Then
        Select Case whichPoint
        Case TopLeft, BottomRight ' pts 0,3 modifies 3,0
            m_DestPts(whichPoint).X = X: m_DestPts(whichPoint).Y = Y
            If m_Style = warpSkew Then
                m_DestPts(3 - whichPoint).X = m_DestPts(Abs(whichPoint - 1)).X + m_DestPts(Abs(whichPoint - 2)).X - m_DestPts(whichPoint).X
                m_DestPts(3 - whichPoint).Y = m_DestPts(Abs(whichPoint - 1)).Y + m_DestPts(Abs(whichPoint - 2)).Y - m_DestPts(whichPoint).Y
            End If
            BuildWarpMatrix hDC
            UpdateDestPoint = True
        
        Case TopRight, BottomLeft ' pts 1,2 modifies 2,1
            m_DestPts(whichPoint).X = X: m_DestPts(whichPoint).Y = Y
            If m_Style = warpSkew Then
                m_DestPts(3 - whichPoint).X = m_DestPts((whichPoint - 1) * 3).X + m_DestPts((2 - whichPoint) * 3).X - m_DestPts(whichPoint).X
                m_DestPts(3 - whichPoint).Y = m_DestPts((whichPoint - 1) * 3).Y + m_DestPts((2 - whichPoint) * 3).Y - m_DestPts(whichPoint).Y
            End If
            BuildWarpMatrix hDC
            UpdateDestPoint = True
        End Select
    End If
End Function

Public Function OffsetDestination(ByVal Xoffset As Single, ByVal Yoffset As Single, Optional hDC As Long) As Boolean

    ' function shifts the path by X,Y offsets
    ' hDC if provided will draw the updated path to the passed DC
    
    If (m_Status And statBoundsApplied) = statBoundsApplied Then
        On Error Resume Next ' errors? Possibly if user sets huge offset & overflow occurs
        Dim I As Long, tPts() As PointF
        tPts() = m_DestPts()
        For I = 0 To 3
            m_DestPts(I).X = m_DestPts(I).X + Xoffset
            m_DestPts(I).Y = m_DestPts(I).Y + Yoffset
        Next
        If Err Then
            Err.Clear
            m_DestPts = tPts
        Else
            BuildWarpMatrix hDC
            OffsetDestination = True
        End If
    End If
End Function

Public Property Let WarpStyle(Style As gdipWarpModes)
    ' option to set the warp mode/style
    If Style >= warpPerspective And Style <= warpSkew Then
        If Not Style = m_Style Then
            m_Style = Style
            If m_Style = warpSkew Then
                ' UpdateDestPoint will recalc the bottom right corner as needed
                UpdateDestPoint m_DestPts(0).X, m_DestPts(0).Y, TopLeft
            Else
                BuildWarpMatrix 0&
            End If
        End If
    End If
End Property
Public Property Get WarpStyle() As gdipWarpModes
    WarpStyle = m_Style
End Property

Public Sub SetOutLine(ByVal PenThickness As Single, ByVal PenColor As Long, Optional ByVal Opacity As Long = 100)
    ' Outline pen color or null color
    ' Pass -1 for PenColor to assign a null pen
    If m_FillObject.Pen Then GdipDeletePen m_FillObject.Pen
    m_FillObject.Pen = 0&
    m_FillObject.PenColor = PenColor
    If PenColor > -1 Then
        If CreateBGRobject(m_FillObject.PenColor, Opacity) Then
            GdipCreatePen1 m_FillObject.PenColor, PenThickness, 2, m_FillObject.Pen
        End If
    End If
End Sub

Public Sub SetBrush(ByVal Color As Long, Optional ByVal gradientColor As Long = -1, Optional gradientDirection As LinearGradientMode, Optional ByVal Opacity As Long = 100)
    ' Fill brush, either solid, gradient or null
    ' Pass -1 for Color to assign a null brush
    With m_FillObject
        If .Brush Then GdipDeleteBrush .Brush
        .Brush = 0&
        .color1 = Color
        .GradDirection = -1 ' non-gradient brush, default
        If .color1 > -1 Then
            If CreateBGRobject(.color1, Opacity) Then
                .color2 = gradientColor
                If .color2 > -1 Then
                    If CreateBGRobject(.color2, Opacity) Then
                        .GradDirection = gradientDirection
                        If (m_Status And statRecalcWarp) = 0& Then
                            GdipCreateLineBrushFromRectI .GradBounds, .color1, .color2, .GradDirection, 0&, .Brush
                        End If
                    End If
                End If
                If .GradDirection = -1 Then ' use solid brush
                    GdipCreateSolidFill .color1, m_FillObject.Brush
                End If
            End If
        End If
    End With
End Sub

Public Sub GetBoundingRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
    ' Retrieves the bounding rectangle of the current warped path
    If m_Path = 0& Then
        Left = 0&: Top = 0&
        Right = 0&: Bottom = 0&
    Else
        Dim tPath As Long, srcPts() As Single, ptTypes() As Byte
        If (m_Status And statRecalcWarp) = statRecalcWarp Then
            tPath = WarpPoints(srcPts, ptTypes)
            If tPath Then GdipDeletePath tPath
        End If
        With m_FillObject
            Left = .GradBounds.X: Top = .GradBounds.Y
            Right = .GradBounds.mWidth + Left
            Bottom = .GradBounds.mHeight + Top
        End With
    End If

End Sub

Public Function GetBoundingPoints(Points() As Single) As Long

'   If function retuns zero, the points array should be considered null
'       Array is returned as 2D :: (0 to 1, 0 to nrPoints-1)
'       The 1st element of 1st dim are X values, the 2nd are Y values
'       There are always four points in the array

'    // bouding points are arranged in Z-configuration
'    //
'    //      0 ------- 1
'    //      |         |
'    //      |         |
'    //      2 ------- 3
'    //
    If (m_Status And statBoundsApplied) = statBoundsApplied Then
        ReDim Points(0 To 1, 0 To 3)
        CopyMemory Points(0, 0), m_DestPts(0), 32&
        GetBoundingPoints = 4
    End If
End Function

Public Function GetPathPoints(Points() As Single, PointType() As Byte) As Long
        ' returns the path points & path types
        ' Points when returned is 2D array: (0 to 1, 0 to nrPts-1)
        '       the X coords are 1D 0 element, Y is 1 element
        ' PointType are members of the GDI+ documented PathPointType enumeration
        ' If function returns 0 consider arrays null
        If m_Path Then
            Dim tPath As Long, Count As Long
            tPath = WarpPoints(Points, PointType, Count)
            If tPath Then GdipDeletePath tPath
            GetPathPoints = Count
        End If

End Function


Private Function AppendPath(fromPath As Long, bReset As Boolean, hDC As Long) As Boolean

    Dim fRect As RECTL
    
    m_Status = m_Status Or statRecalcWarp Or statTextApplied
    
    If m_Path = 0& Then
        m_Path = fromPath
    Else
        If bReset Then GdipResetPath m_Path
        GdipAddPathPath m_Path, fromPath, True
        GdipDeletePath fromPath
    End If
    ' get the actual bounds of the path before any warping is applied
    GdipGetPathWorldBoundsI m_Path, fRect, 0&, 0&
    m_SrcPts(0).X = fRect.X: m_SrcPts(0).Y = fRect.Y
    m_SrcPts(1).X = fRect.mWidth + fRect.X: m_SrcPts(1).Y = m_SrcPts(0).Y
    m_SrcPts(2).X = m_SrcPts(0).X: m_SrcPts(2).Y = fRect.mHeight + m_SrcPts(0).Y
    m_SrcPts(3).X = m_SrcPts(1).X: m_SrcPts(3).Y = m_SrcPts(2).Y
        
    ' recalculate the warp matrix only if destination points have been set
    If (m_Status And statBoundsApplied) = statBoundsApplied Then BuildWarpMatrix hDC
    AppendPath = True
    
End Function
Private Function CreateBGRobject(Color As Long, ByVal Opacity As Long) As Boolean
    ' Local use to create RGB color to BGRA
    If Opacity < 1 Then
        Exit Function
    ElseIf Opacity > 100 Then
        Opacity = 100
    End If
    Opacity = (255 * Opacity) \ 100
    Color = ((Color And &HFF) * &H10000) Or _
        (((Color \ &H100) And &HFF) * &H100) Or ((Color \ &H10000) And &HFF)
    If Opacity > 127 Then
        Color = Color Or (&H80000000 Or ((Opacity - 128) * &H1000000))
    Else
        Color = (Opacity * &H1000000) Or Color
    End If
    CreateBGRobject = True
    
End Function

Private Sub BuildWarpMatrix(hDC As Long)

    ' Visit the following link to better understand. This matrix and the associated
    ' SolveEquation routines are required to manually warp points vs using the
    ' flawed GDI+ function.  Again, look at the link
    ' http://www.codeguru.com/cpp/g-m/gdi/gdi/article.php/c3657/
    
    Dim aMatrix() As Single
    Dim Row As Long, Index As Long
    
    m_Status = m_Status Or statRecalcWarp Or statBoundsApplied
    
    If (m_Status And statTextApplied) = statTextApplied Then ' else no source points
        
        If m_Style = warpPerspective Then
       
    '        // In perspective mode, the warping functions are:
    '        //  x' = (a0 + a1 x + a2 y) / (c0 x + c1 y + 1)
    '        //  y' = (b0 + b1 x + b2 y) / (c0 x + c1 y + 1)
    '        //
    '        // The following calculates the factors a#, b# and c#.
    '        // We do this by creating a set of eight equations with a#, b# and c# as unknowns.
    '        // The equations are derived by:
    '        // 1. substituting the srcPoints for (x, y);
    '        // 2. substituting the corresponding destPoints for (x', y');
    '        // 3. solving the resulting set of equations, with the factors as unknowns.
    '        //
    '        // The equations are like these:
    '        //  a0  x a1    y a2    0       0       0       -xx'c0  -yx'c1  = x'
    '        //  0   0       0       b0      x b1    y b2    -xy'c0  -yy'c1  = y'
    '        // The known factors of left hand side ar put in the 8x8 matrix mxLeft for
    '        // all four point pairs, and the right hand side in the one column matrix mxRight.
    '        // After solving, m_mxWarpFactors contains a0, a1, a2, b0, b1, b2, c0, c1.
    
            ReDim aMatrix(0 To 7, 0 To 7)
            ReDim m_WarpFactors(0 To 7, 0 To 0)
            
            For Index = 0 To 3
                aMatrix(Row, 0) = 1!
                aMatrix(Row, 1) = m_SrcPts(Index).X
                aMatrix(Row, 2) = m_SrcPts(Index).Y
                aMatrix(Row, 6) = -m_SrcPts(Index).X * m_DestPts(Index).X
                aMatrix(Row, 7) = -m_SrcPts(Index).Y * m_DestPts(Index).X
                m_WarpFactors(Row, 0) = m_DestPts(Index).X
                Row = Row + 1
                aMatrix(Row, 3) = 1!
                aMatrix(Row, 4) = m_SrcPts(Index).X
                aMatrix(Row, 5) = m_SrcPts(Index).Y
                aMatrix(Row, 6) = -m_SrcPts(Index).X * m_DestPts(Index).Y
                aMatrix(Row, 7) = -m_SrcPts(Index).Y * m_DestPts(Index).Y
                m_WarpFactors(Row, 0) = m_DestPts(Index).Y
                Row = Row + 1
            Next
        
        Else ' BiLinear & Simple Skew
        
    '        // In bilinear mode, the warping functions are:
    '        //  x' = a0 + a1 x y + a2 x + a3 y
    '        //  y' = b0 + b1 x y + b2 x + b3 y
    '        //
    '        // Here, we have two sets of four equations. In the first set, the a# factors
    '        // are the unknowns, in the second set the b# factors.
    '        // The equations are of the form:
    '        //  a0      + xy a1     + x a2      + y a3  = x'
    '        // The left hand side is identical for both sets. The right hand side differs.
    '        // Therefore, we can solve them in one operation.
    '        // The left hand side factors are put in the 4x4 matrix mxLeft, the right side
    '        // factors are put in the 4x2 matrix mxRight.
    '        // After solving, the first column of m_mxWarpFactors contains a0, a1, a2, a3; the
    '        // second columne contains b0, b1, b2, b3.
    
            ReDim aMatrix(0 To 3, 0 To 3)
            ReDim m_WarpFactors(0 To 3, 0 To 1)
    
            For Index = 0 To 3
                aMatrix(Index, 0) = 1
                aMatrix(Index, 1) = m_SrcPts(Index).X * m_SrcPts(Index).Y
                aMatrix(Index, 2) = m_SrcPts(Index).X
                aMatrix(Index, 3) = m_SrcPts(Index).Y
                
                m_WarpFactors(Index, 0) = m_DestPts(Index).X
                m_WarpFactors(Index, 1) = m_DestPts(Index).Y
            Next
            
        End If
        
        m_WarpFactors = SolveEquation(aMatrix)
        
    End If
    
    If Not hDC = 0 Then DrawWarpPath hDC

End Sub

Private Function SolveEquation(sourceArray() As Single) As Single()

    ' Using a matrix to solve simultaneous equations
    ' http://www.java2s.com/Tutorial/VB/0040__Data-Type/Solveequationsusingmatrices.htm
    ' Not commented: see above link. I'd need a couple of pages to try to explain it
    
    Dim I As Long, J As Long, K As Long
    Dim Rank As Long, RankUnk As Long
    Dim a1 As Single
    Dim temp() As Single, Solution() As Single
    
    Rank = UBound(sourceArray, 2)
    
On Error GoTo EH
    
    RankUnk = UBound(m_WarpFactors, 2) + 1
    ReDim temp(0 To Rank, 0 To Rank + RankUnk)
    
    For I = 0 To Rank
        For J = 0 To Rank
            temp(I, J) = sourceArray(I, J)
        Next
        For K = 0 To RankUnk - 1
            temp(I, J + K) = m_WarpFactors(I, K)
        Next
    Next
    
    For K = 0 To Rank
        If Pivot(temp, K, Rank) = False Then
            ' singular matrix, no solution
            m_Status = m_Status Or statInvalid
            Exit Function
        End If
        a1 = temp(K, K)
        If a1 = 0 Then Stop ' debugging
        For J = K To Rank + RankUnk
            temp(K, J) = temp(K, J) / a1
        Next
        For I = K + 1 To Rank
            a1 = temp(I, K)
            For J = K To Rank + RankUnk
                temp(I, J) = temp(I, J) - a1 * temp(K, J)
            Next
        Next
    Next
    
    ReDim Solution(0 To Rank, 0 To RankUnk - 1)
    For K = 0 To RankUnk - 1
        For I = Rank To 0 Step -1
            Solution(I, K) = temp(I, Rank + 1 + K)
            For J = I + 1 To Rank
                Solution(I, K) = Solution(I, K) - temp(I, J) * Solution(J, K)
            Next
        Next
    Next
    m_Status = (m_Status And Not statInvalid)
    SolveEquation = Solution
EH:
If Err Then
    Stop                ' debugging purposes only
    Err.Clear
    Resume
End If
End Function

Private Function Pivot(srcMatrix() As Single, Row As Long, Rank As Long) As Boolean

    ' Helper function of the SolveEquation function
    Dim K As Long, I As Long
    Dim aMax As Single, tmpVal As Single
    
    K = Row
    For I = Row To Rank ' row
        tmpVal = Abs(srcMatrix(I, Row))
        If tmpVal > aMax Then
            If tmpVal <> 0! Then
                aMax = tmpVal
                K = I
            End If
         End If
    Next
    If Not srcMatrix(K, Row) = 0! Then
        If K <> Row Then
            For I = 0 To UBound(srcMatrix, 2)
                tmpVal = srcMatrix(K, I)
                srcMatrix(K, I) = srcMatrix(Row, I)
                srcMatrix(Row, I) = tmpVal
            Next
        End If
        Pivot = True
    End If
End Function

Private Function WarpPoints(srcPts() As Single, ptTypes() As Byte, Optional Count As Long) As Long
    
    ' This routine manually warps points
    
    Dim Index As Long, tPath As Long
    Dim X As Single, Y As Single, tVal As Single
    Dim pd As PathData
    
    GdipGetPointCount m_Path, pd.Count
    
    If pd.Count Then
        ' get the points from the non-warped path
        ReDim srcPts(0 To 1, 0 To pd.Count - 1)
        ReDim ptTypes(0 To pd.Count - 1)
        pd.Points = VarPtr(srcPts(0, 0))
        pd.Types = VarPtr(ptTypes(0))
        GdipGetPathData m_Path, VarPtr(pd)
        
        If (m_Status And statInvalid) = 0& Then ' see SolveEquation
            
            If m_Style = warpPerspective Then ' formula for perspective warping
                For Index = 0 To pd.Count - 1
                    X = srcPts(0, Index)
                    Y = srcPts(1, Index)
                    
                    tVal = m_WarpFactors(6, 0) * X + m_WarpFactors(7, 0) * Y + 1!
                    srcPts(0, Index) = ((m_WarpFactors(0, 0) + m_WarpFactors(1, 0) * X + m_WarpFactors(2, 0) * Y) / tVal)
                    srcPts(1, Index) = ((m_WarpFactors(3, 0) + m_WarpFactors(4, 0) * X + m_WarpFactors(5, 0) * Y) / tVal)
                Next
            Else
                For Index = 0 To pd.Count - 1   ' formula for bilinear & simple skew warping
                    X = srcPts(0, Index)
                    Y = srcPts(1, Index)
                    tVal = X * Y
                    srcPts(0, Index) = m_WarpFactors(0, 0) + m_WarpFactors(1, 0) * tVal + m_WarpFactors(2, 0) * X + m_WarpFactors(3, 0) * Y
                    srcPts(1, Index) = m_WarpFactors(0, 1) + m_WarpFactors(1, 1) * tVal + m_WarpFactors(2, 1) * X + m_WarpFactors(3, 1) * Y
                Next
        
            End If
        End If
        ' GDI+ has no update path points function; therefore, we will create a new path
        ' using these points, then that path will be added to clone to create the new warp
        GdipGetPathFillMode m_Path, Index
        GdipCreatePath2 VarPtr(srcPts(0, 0)), VarPtr(ptTypes(0)), pd.Count, Index, tPath
        
        ' determine if new gradient brush needs to be created
        If (m_Status And statRecalcWarp) = statRecalcWarp Then
            m_Status = (m_Status And Not statRecalcWarp)
            GdipGetPathWorldBoundsI tPath, m_FillObject.GradBounds, 0&, 0&
            If m_FillObject.GradDirection > -1 Then
                If m_FillObject.Brush Then GdipDeleteBrush m_FillObject.Brush
                m_FillObject.Brush = 0&
                With m_FillObject
                    GdipCreateLineBrushFromRectI .GradBounds, .color1, .color2, .GradDirection, 0&, .Brush
                End With
            End If
        End If
    
    End If
    Count = pd.Count        ' return point count
    WarpPoints = tPath      ' return GDI+ path -- must be destroyed

End Function

Public Sub DrawWarpPath(hDC As Long)

    ' function draws a warped path.
    ' Note: If a path was rendered over & over again, it would be a better idea to
    ' cache the current warp, vs re-warping every re-drawing.  Your choice
    
    If (m_Status And statCanDraw) = 0& Then Exit Sub
    If hDC = 0& Then Exit Sub

    Dim lGraphics As Long
    Dim rPath As Long, tPath As Long
    Dim srcPts() As Single, ptTypes() As Byte
    
    rPath = WarpPoints(srcPts, ptTypes) ' create path with warped points
    If rPath = 0& Then rPath = m_Path
    
    ' now draw the path
    Call GdipCreateFromHDC(hDC, lGraphics)
    ' set output quality
    GdipSetSmoothingMode lGraphics, SmoothingModeHighQuality
    GdipSetInterpolationMode lGraphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode lGraphics, PixelOffsetModeHighQuality
    
    ' fill the path
    If m_FillObject.Brush Then GdipFillPath lGraphics, m_FillObject.Brush, rPath
    ' outline the path
    If m_FillObject.Pen Then GdipDrawPath lGraphics, m_FillObject.Pen, rPath
    
    ' clean up
    GdipDeleteGraphics lGraphics
    If Not rPath = m_Path Then GdipDeletePath rPath
    
End Sub

Private Sub Class_Initialize()
    ReDim m_SrcPts(0 To 3)
    ReDim m_DestPts(0 To 3)
    m_Style = warpBilinear
End Sub

Private Sub Class_Terminate()
    ' clean up
    If m_Path Then GdipDeletePath m_Path
    If m_FillObject.Pen Then GdipDeletePen m_FillObject.Pen
    If m_FillObject.Brush Then GdipDeleteBrush m_FillObject.Brush
End Sub
