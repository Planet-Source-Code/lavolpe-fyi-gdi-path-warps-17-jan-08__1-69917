VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GDI+ Path Warping"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picShapes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   60
      ScaleHeight     =   1545
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   60
      Width           =   1875
      Begin VB.OptionButton optShape 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play with Ellipse"
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   25
         Top             =   1260
         Width           =   1710
      End
      Begin VB.OptionButton optShape 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play with Rectangle"
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   24
         Top             =   990
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   300
         Width           =   1740
      End
      Begin VB.OptionButton optShape 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play with Text"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   60
         Width           =   1710
      End
   End
   Begin VB.OptionButton optWarp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simple Skew"
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   17
      Top             =   2265
      Width           =   1770
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manually Warp Text - Can change Warp Type"
      Height          =   270
      Left            =   2595
      TabIndex        =   16
      Top             =   3975
      Width           =   3750
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   975
      TabIndex        =   15
      Top             =   4260
      Width           =   900
   End
   Begin VB.CheckBox chkPen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Outline Color"
      Height          =   270
      Left            =   480
      TabIndex        =   9
      Top             =   2595
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3195
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboGradient 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   90
      List            =   "Form1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3600
      Width           =   1770
   End
   Begin VB.CheckBox chkFillType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gradient Fill"
      Height          =   255
      Index           =   1
      Left            =   495
      TabIndex        =   5
      Top             =   3315
      Width           =   1350
   End
   Begin VB.CheckBox chkFillType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solid Fill"
      Height          =   255
      Index           =   0
      Left            =   495
      TabIndex        =   4
      Top             =   3030
      Width           =   1245
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   4260
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   1950
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   3
      Top             =   45
      Width           =   4890
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H0043E9D8&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   5
         Left            =   2280
         MousePointer    =   7  'Size N S
         TabIndex        =   20
         ToolTipText     =   "Vertical Sizing Only"
         Top             =   765
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   6
         Left            =   2685
         MousePointer    =   5  'Size
         TabIndex        =   19
         ToolTipText     =   "Move Object"
         Top             =   420
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H0043E9D8&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   4
         Left            =   1875
         MousePointer    =   9  'Size W E
         TabIndex        =   18
         ToolTipText     =   "Horizontal Sizing Only"
         Top             =   1155
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4B1AC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   660
         MousePointer    =   2  'Cross
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4B1AC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   1365
         MousePointer    =   2  'Cross
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4B1AC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   330
         MousePointer    =   2  'Cross
         TabIndex        =   12
         Top             =   465
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4B1AC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   1005
         MousePointer    =   2  'Cross
         TabIndex        =   11
         Top             =   1170
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin VB.OptionButton optWarp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BiLinear Warp"
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   1980
      Width           =   1905
   End
   Begin VB.OptionButton optWarp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Perspective Warp"
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   1695
      Width           =   1905
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   3315
      Width           =   360
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   3030
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' An example of warping paths using GDI+

' The class has acknowledgements and is well documented.
' The class can be modified and expanded to include other path options too, like
' adding shapes, images, etc. When I have time, I may update the class for that purpose.

' This project is just a sample to get your creative juices flowing. Other things that
' can be added for example, radial gradients and paths that follow other paths (curves).


Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    ByRef lpOutput As GdiplusStartupOutput) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

' used for manually drawing an outline of a GDI+ path
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hDC As Long, ByRef lppt As POINTAPI, ByVal cCount As Long) As Long
Private Enum gdiPathPointType
    PathPointTypeStart = 0
    PathPointTypeLine = 1
    PathPointTypeBezier = 3
    PathPointTypePathTypeMask = &H7
    PathPointTypeDashMode = &H10
    PathPointTypePathMarker = &H20
    PathPointTypeCloseSubpath = &H80
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

' used when manually warping path
Private curX As Single, curY As Single  ' used to drag labels
Private pathPoints() As Single          ' GDI+ path point X,Y coords
Private pathType() As Byte              ' GDI+ path types: see gdiPathPointType above
Private pathPtCount As Long             ' number of points in the path

Private GdipToken As Long
Private cGDIwarper As gdipPathWarper

Private Sub Form_Load()
    StartUpGDIPlus 1        ' start up gdi
    If GdipToken = 0& Then
        MsgBox "Failed to start GDI+, closing application", vbExclamation + vbOKOnly
        Unload Me
        Exit Sub
    End If
    
    Set cGDIwarper = New gdipPathWarper
    Picture1.ScaleMode = vbPixels   ' want pixel scalemode when doing graphics
    Picture1.AutoRedraw = True
    cboGradient.ListIndex = 1       ' set initial combobox item
    
    ' set up initial pens/brushes and stuff
    chkPen.Value = 1            ' use outline pen
    chkFillType(1) = 1          ' use solid vs gradient brush
    optWarp(0) = True           ' set initial warp mode
    
    Text1.Text = "LaVolpe"
    optShape(0) = True

    ' modify the path coordinates to squeeze vertically the right edge of the path
    cGDIwarper.UpdateDestPoint 247, 110, TopRight
    cGDIwarper.UpdateDestPoint 247, 150, BottomRight
    Call cmdRefresh_Click
    Show                        ' show our form

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Tip: When using GDI+, always unload any classes/objects that can be using
    ' GDI+ before you actually shutdown GDI+.
    ' Best to put GDI+ shutdown in terminate event
    Set cGDIwarper = Nothing
    ShutdownGDIPlus

End Sub

Private Sub cboGradient_Click()
    ' option to change gradient -- uses GDI+ settings
    If chkFillType(1).Value Then
        cGDIwarper.SetBrush lblColor(0).BackColor, lblColor(1).BackColor, cboGradient.ListIndex
        Call cmdRefresh_Click
    End If

End Sub

Private Sub Check1_Click()
    ' option to manually warp path
    
    Dim I As Integer, halfCx As Long, handlePts() As Single
    Dim bEnabled As Boolean
    
    If Check1.Value Then ' manually warping path
    
        ' get its point X,Y coords and the type of each point
        pathPtCount = cGDIwarper.GetPathPoints(pathPoints(), pathType())
        
        ' position handles & make them visible
        PositionHandles True, True
        For I = 0 To lblHandle.UBound
            lblHandle(I).Visible = True
        Next
        bEnabled = False    ' disable other controls except the warp option buttons
        
        ' show message first time only
        If Check1.Tag = vbNullString Then
            Check1.Tag = "NoMsgBox"
            MsgBox "When warping and points cross over opposite bounds, mirroring of the path occurs." & vbNewLine & _
                "This mirroring is agreeable only in BiLinear mode. Perspective mirroring produces poor results." & vbNewLine & vbNewLine & _
                "To see the diffeernt results, drag the top left handle to the bottom/center of the picturebox. " & vbNewLine & _
                "Then toggle the BiLinear & Perspective warp options", vbInformation + vbOKOnly, "Path Mirroring"
        End If
    
        DrawSelectionBox -1 ' show the bounding rectangle & path
        
    Else
        Erase pathPoints
        Erase pathType
        For I = 0 To lblHandle.UBound   ' hide the handles
            lblHandle(I).Visible = False
        Next
        Call cmdRefresh_Click       ' refresh
        bEnabled = True             ' enable other controls
    End If
    
    ' enable/disable controls
    cboGradient.Enabled = bEnabled
    For I = 0 To lblColor.UBound
        lblColor(I).Enabled = bEnabled
    Next
    For I = 0 To chkFillType.UBound
        chkFillType(I).Enabled = bEnabled
    Next
    chkPen.Enabled = bEnabled
    Text1.Enabled = bEnabled
    cmdRefresh.Enabled = bEnabled
    cmdReset.Enabled = bEnabled
    picShapes.Enabled = bEnabled

End Sub

Private Sub chkFillType_Click(Index As Integer)
    ' option to use solid or gradient brush
    ' Note: the class SetBrush function also has an Opacity setting where brushes can be semitransparent
    If chkFillType(0).Tag = vbNullString Then
        If chkFillType(Abs(Index - 1)).Value Then   ' only one check box, uncheck the other
            chkFillType(0).Tag = "No Update"
            chkFillType(Abs(Index - 1)).Value = 0
            chkFillType(0).Tag = vbNullString
        End If
        
        If chkFillType(0) = 1 Then ' solid brush
            cGDIwarper.SetBrush lblColor(Index).BackColor
        ElseIf chkFillType(1) = 1 Then ' gradient brush
            cGDIwarper.SetBrush lblColor(0).BackColor, lblColor(1).BackColor, cboGradient.ListIndex
        Else ' nothing is selected
            cGDIwarper.SetBrush -1 ' transparent fill, no brush
        End If
        Call cmdRefresh_Click       ' refresh
    End If
End Sub

Private Sub chkPen_Click()
    ' option to use pen
    ' Note: the class SetOutline function also has an Opacity setting where pens can be semitransparent
    If chkPen.Value Then
        cGDIwarper.SetOutLine 2, lblColor(2).BackColor
    Else
        cGDIwarper.SetOutLine 2, -1 ' no outline color
    End If
    Call cmdRefresh_Click   ' refresh
End Sub

Private Sub cmdRefresh_Click()
    If Check1.Value Then                ' manual warping now
        DrawSelectionBox 5
        PositionHandles False, True
    Else                                ' refresh
        Picture1.Cls
        cGDIwarper.DrawWarpPath Picture1.hDC
        Picture1.Refresh
    End If
End Sub

Private Sub cmdReset_Click()
    ' reset path to a non-warped state
    ' Warped is simply a non-rectangular path
    cGDIwarper.SetPathDest_Rect 40, 53, 247, 188
    Call cmdRefresh_Click
End Sub

Private Sub lblColor_Click(Index As Integer)
    ' allow brush & color options
    With CommonDialog1
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .CancelError = True
        .Color = lblColor(Index).BackColor
    End With
    On Error GoTo EH
    CommonDialog1.ShowColor
    lblColor(Index).BackColor = CommonDialog1.Color
    If Index = 2 Then
        If chkPen.Value Then Call chkPen_Click
    ElseIf chkFillType(Index).Value And Index = 0 Then
        Call chkFillType_Click(Index)
    ElseIf chkFillType(1).Value Then
        Call chkFillType_Click(1)
    End If
EH:
If Err Then Err.Clear ' user pressed cancel
End Sub

Private Sub lblHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then curX = X: curY = Y   ' cache handle's current x,y
End Sub

Private Sub lblHandle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then PositionHandles False, True
End Sub

Private Sub lblHandle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ' update warp outline
        If Index = 4 Then ' resizing horizontally only
            X = ((X - curX) \ Screen.TwipsPerPixelX)
            lblHandle(Index).Left = lblHandle(Index).Left + X
            cGDIwarper.UpdateDestPoint lblHandle(0).Left + lblHandle(0).Width \ 2 + X, lblHandle(0).Top + lblHandle(0).Height \ 2, TopLeft
            cGDIwarper.UpdateDestPoint lblHandle(2).Left + lblHandle(2).Width \ 2 + X, lblHandle(2).Top + lblHandle(0).Height \ 2, BottomLeft
            DrawSelectionBox Index
        ElseIf Index = 5 Then 'resizing vertically only
            Y = ((Y - curX) \ Screen.TwipsPerPixelX)
            lblHandle(Index).Top = lblHandle(Index).Top + Y
            cGDIwarper.UpdateDestPoint lblHandle(0).Left + lblHandle(0).Width \ 2, lblHandle(0).Top + lblHandle(0).Height \ 2 + Y, TopLeft
            cGDIwarper.UpdateDestPoint lblHandle(1).Left + lblHandle(0).Width \ 2, lblHandle(1).Top + lblHandle(0).Height \ 2 + Y, TopRight
            DrawSelectionBox Index
        ElseIf Index = 6 Then ' moving vs resizing
            X = ((X - curX) \ Screen.TwipsPerPixelX)
            Y = ((Y - curX) \ Screen.TwipsPerPixelX)
            lblHandle(Index).Move lblHandle(Index).Left + X, lblHandle(Index).Top + Y
            cGDIwarper.OffsetDestination X, Y
            DrawSelectionBox 6
        Else
            cGDIwarper.UpdateDestPoint lblHandle(Index).Left + (X - curX) \ Screen.TwipsPerPixelX, lblHandle(Index).Top + (Y - curY) \ Screen.TwipsPerPixelX, Index
            DrawSelectionBox Index
        End If
    End If
End Sub

Private Sub optShape_Click(Index As Integer)
    Picture1.Cls
    Select Case Index
        Case 0: Call Text1_LostFocus
        Case 1: cGDIwarper.SetPathShape_Rectangle 40, 53, 247, 188
                Call cmdRefresh_Click
        Case 2: cGDIwarper.SetPathShape_Ellipse 40, 53, 247, 188
                Call cmdRefresh_Click
    End Select
End Sub

Private Sub optWarp_Click(Index As Integer)
    ' option to use BiLinear or Perspective Warp
    If optWarp(0) Then ' perspective
        cGDIwarper.WarpStyle = warpPerspective
    ElseIf optWarp(1) Then ' bilinear
        cGDIwarper.WarpStyle = warpBilinear
    Else    ' skew - warp shape is parallelogram
        cGDIwarper.WarpStyle = warpSkew
    End If
    Call cmdRefresh_Click
    
    If optWarp(0).Tag = vbNullString Then
        If cGDIwarper.WarpStyle = warpSkew Then
            optWarp(0).Tag = "shown"
            MsgBox "FYI: When a warp path is a parallelogram/rectangle, " & vbNewLine & _
                "all warp modes produce the same results.", vbInformation + vbOKOnly
        End If
    End If
    
    ' FYI: When warp path is a perfect parallelogram/rectangle,
    ' all warp options produce the same results.
    ' Skewing is simply bilinear/perspective warping forcing use of a parallelogram shape
    
End Sub


Private Function ShutdownGDIPlus() As Long
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)    ' shut down GDI+
End Function

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Long
    ' Initialisieren der GDI+ Instanz
    Dim GdipStartupInput As GDIPlusStartupInput
    Dim GdipStartupOutput As GdiplusStartupOutput
    
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, GdipStartupOutput)
End Function

Private Sub PositionHandles(Optional blueHandles As Boolean, Optional goldHandles As Boolean)
    ' simply positions handle labels
    Dim I As Integer, halfCx As Long, handlePts() As Single
    
    halfCx = lblHandle(0).Width \ 2
    cGDIwarper.GetBoundingPoints handlePts()
    
    ' Note. GDI path points are Z-clockwise order:
    '   -- Top Left, Top Right, Bottom Left, Bottom Right
    If blueHandles Then ' positiong blue corner labels
        For I = 0 To 3
            lblHandle(I).Move handlePts(0, I) - halfCx, handlePts(1, I) - halfCx
        Next
    End If
    If goldHandles Then ' position additional handles
        lblHandle(4).Move handlePts(0, 0) + (handlePts(0, 2) - handlePts(0, 0)) \ 2 - halfCx, handlePts(1, 0) + (handlePts(1, 2) - handlePts(1, 0)) \ 2
        lblHandle(5).Move handlePts(0, 2) + (handlePts(0, 3) - handlePts(0, 2)) \ 2, handlePts(1, 2) + (handlePts(1, 3) - handlePts(1, 2)) \ 2
        lblHandle(6).Move lblHandle(5).Left, lblHandle(4).Top
    End If
    
End Sub

Private Sub DrawSelectionBox(ByVal UpdatePos As Long)
    ' draw the path's bounding rectangle
    
    Dim I As Integer, halfCx As Long, handlePts() As Single
    halfCx = lblHandle(0).Width \ 2
    
    With Picture1
        .DrawMode = 10 ' notxorpen
        .DrawStyle = 2 ' dotted pen
        
        ' draw box. If already drawn, this will erase it
        .ForeColor = RGB(192, 192, 192) ' light gray
        Picture1.Line (lblHandle(0).Left + halfCx, lblHandle(0).Top + halfCx)-(lblHandle(1).Left + halfCx, lblHandle(1).Top + halfCx)
        Picture1.Line (lblHandle(1).Left + halfCx, lblHandle(1).Top + halfCx)-(lblHandle(3).Left + halfCx, lblHandle(3).Top + halfCx)
        Picture1.Line (lblHandle(3).Left + halfCx, lblHandle(3).Top + halfCx)-(lblHandle(2).Left + halfCx, lblHandle(2).Top + halfCx)
        Picture1.Line (lblHandle(2).Left + halfCx, lblHandle(2).Top + halfCx)-(lblHandle(0).Left + halfCx, lblHandle(0).Top + halfCx)
        
        .ForeColor = vbRed
        OutlinePath ' draw path & if already drawn, this will erase it
        
        If UpdatePos > -1 Then  ' else first time thru & we won't need to do this
        
            ' draw new bounding rectangle
            .ForeColor = RGB(192, 192, 192)
            PositionHandles True
            Picture1.Line (lblHandle(0).Left + halfCx, lblHandle(0).Top + halfCx)-(lblHandle(1).Left + halfCx, lblHandle(1).Top + halfCx)
            Picture1.Line (lblHandle(1).Left + halfCx, lblHandle(1).Top + halfCx)-(lblHandle(3).Left + halfCx, lblHandle(3).Top + halfCx)
            Picture1.Line (lblHandle(3).Left + halfCx, lblHandle(3).Top + halfCx)-(lblHandle(2).Left + halfCx, lblHandle(2).Top + halfCx)
            Picture1.Line (lblHandle(2).Left + halfCx, lblHandle(2).Top + halfCx)-(lblHandle(0).Left + halfCx, lblHandle(0).Top + halfCx)
            
            ' get the points for the new warped path & draw the outline
            pathPtCount = cGDIwarper.GetPathPoints(pathPoints(), pathType())
            .ForeColor = vbRed
            OutlinePath
            
        End If
        .ForeColor = vbBlack
        .DrawMode = 13  ' reset back to normal
        .DrawStyle = 0
        Picture1.Refresh
    End With
End Sub

Private Sub OutlinePath()
    
    ' this is an example how to manually render the path
    ' It isn't too complicated & I thought I'd add it for fun
    
    Dim I As Long, lastPt As Long, bzPt(0 To 2) As POINTAPI
    If pathPtCount Then
        
        For I = 0 To pathPtCount - 1
            Select Case ((pathType(I) And Not PathPointTypeDashMode) And Not PathPointTypePathMarker)
            Case PathPointTypeStart
                If lastPt Then
                    LineTo Picture1.hDC, pathPoints(0, lastPt), pathPoints(1, lastPt)
                End If
                MoveToEx Picture1.hDC, pathPoints(0, I), pathPoints(1, I), ByVal 0&
                lastPt = I
            
            Case PathPointTypeLine Or PathPointTypeCloseSubpath
                LineTo Picture1.hDC, pathPoints(0, I), pathPoints(1, I)
                If lastPt Then
                    LineTo Picture1.hDC, pathPoints(0, lastPt), pathPoints(1, lastPt)
                    lastPt = 0&
                End If
                
            Case PathPointTypeLine
                LineTo Picture1.hDC, pathPoints(0, I), pathPoints(1, I)
                
            Case PathPointTypeBezier Or PathPointTypeCloseSubpath
                ' convert single to long for the API
                bzPt(0).X = pathPoints(0, I): bzPt(0).Y = pathPoints(1, I)
                bzPt(1).X = pathPoints(0, I + 1): bzPt(1).Y = pathPoints(1, I + 1)
                bzPt(2).X = pathPoints(0, I + 2): bzPt(2).Y = pathPoints(1, I + 2)
                PolyBezierTo Picture1.hDC, bzPt(0), 3
                I = I + 2
                If lastPt Then
                    LineTo Picture1.hDC, pathPoints(0, lastPt), pathPoints(1, lastPt)
                    lastPt = 0&
                End If
            
            Case PathPointTypeBezier
                ' convert single to long for the API
                bzPt(0).X = pathPoints(0, I): bzPt(0).Y = pathPoints(1, I)
                bzPt(1).X = pathPoints(0, I + 1): bzPt(1).Y = pathPoints(1, I + 1)
                bzPt(2).X = pathPoints(0, I + 2): bzPt(2).Y = pathPoints(1, I + 2)
                PolyBezierTo Picture1.hDC, bzPt(0), 3
                I = I + 2
            Case Else
                Stop

            End Select
        Next
        If lastPt Then
            LineTo Picture1.hDC, pathPoints(0, lastPt), pathPoints(1, lastPt)
        End If
        
    End If
End Sub

Private Sub Text1_LostFocus()
    ' option to change path text. Only triggers on lost focus
    If Not Text1.Text = vbNullString Then
        Dim tFont As StdFont
        Set tFont = New StdFont
        tFont.name = "Georgia" ' Use true-type fonts only
        tFont.Size = 24 ' arbitrary, can be any size
        If cGDIwarper.SetPathString(Text1.Text, tFont) = False Then
            MsgBox "Failed to create string using the font: " & tFont.name, vbExclamation + vbOKCancel
            Picture1.Cls
        Else
            Call cmdReset_Click ' refresh
        End If
    End If
End Sub
