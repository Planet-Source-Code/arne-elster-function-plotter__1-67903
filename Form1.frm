VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Plotter"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8400
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  '2D
      BackColor       =   &H00A29F7D&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9000
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   21
      Top             =   600
      Width           =   315
   End
   Begin VB.CheckBox chkAntialias 
      Caption         =   "Antialias"
      Height          =   240
      Left            =   6225
      TabIndex        =   14
      Top             =   3600
      Value           =   1  'Aktiviert
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "window"
      Height          =   2115
      Left            =   6150
      TabIndex        =   4
      Top             =   1200
      Width           =   3240
      Begin VB.PictureBox picWnd 
         BorderStyle     =   0  'Kein
         Height          =   1815
         Left            =   75
         ScaleHeight     =   1815
         ScaleWidth      =   2940
         TabIndex        =   5
         Top             =   225
         Width           =   2940
         Begin VB.TextBox txtXSteps 
            Height          =   285
            Left            =   2325
            TabIndex        =   18
            Text            =   "1"
            Top             =   150
            Width           =   540
         End
         Begin VB.TextBox txtYSteps 
            Height          =   285
            Left            =   2325
            TabIndex        =   17
            Text            =   "1"
            Top             =   975
            Width           =   540
         End
         Begin VB.TextBox txtMaxY 
            Height          =   285
            Left            =   675
            TabIndex        =   13
            Text            =   "10"
            Top             =   1350
            Width           =   915
         End
         Begin VB.TextBox txtMinY 
            Height          =   285
            Left            =   675
            TabIndex        =   11
            Text            =   "-10"
            Top             =   975
            Width           =   915
         End
         Begin VB.TextBox txtMaxX 
            Height          =   285
            Left            =   675
            TabIndex        =   9
            Text            =   "10"
            Top             =   525
            Width           =   915
         End
         Begin VB.TextBox txtMinX 
            Height          =   285
            Left            =   675
            TabIndex        =   7
            Text            =   "-10"
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "X Steps:"
            Height          =   195
            Left            =   1650
            TabIndex        =   20
            Top             =   150
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Y Steps:"
            Height          =   195
            Left            =   1650
            TabIndex        =   19
            Top             =   975
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Y Max:"
            Height          =   195
            Left            =   75
            TabIndex        =   12
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Y Min:"
            Height          =   195
            Left            =   75
            TabIndex        =   10
            Top             =   975
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "X Max:"
            Height          =   195
            Left            =   75
            TabIndex        =   8
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "X Min:"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   150
            Width           =   450
         End
      End
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "draw"
      Default         =   -1  'True
      Height          =   390
      Left            =   6450
      TabIndex        =   3
      Top             =   600
      Width           =   1365
   End
   Begin VB.TextBox txtFnc 
      Height          =   305
      Left            =   6450
      TabIndex        =   2
      Text            =   "x^2"
      Top             =   195
      Width           =   2940
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5640
      Left            =   225
      ScaleHeight     =   372
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   0
      Top             =   150
      Width           =   5640
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   195
      Left            =   6225
      TabIndex        =   16
      Top             =   4275
      Width           =   150
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   195
      Left            =   6225
      TabIndex        =   15
      Top             =   4050
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "f(x)="
      Height          =   195
      Left            =   6075
      TabIndex        =   1
      Top             =   225
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" _
Alias "DrawTextA" ( _
    ByVal hdc As Long, _
    ByVal lpStr As String, ByVal nCount As Long, _
    lpRect As RECT, ByVal wFormat As Long _
) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long _
) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Private Const DT_BOTTOM         As Long = &H8
Private Const DT_CENTER         As Long = &H1
Private Const DT_LEFT           As Long = &H0
Private Const DT_RIGHT          As Long = &H2
Private Const DT_TOP            As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_WORDBREAK      As Long = &H10

Private Const DT_CALCRECT       As Long = &H400
Private Const DT_EDITCONTROL    As Long = &H2000
Private Const DT_NOCLIP         As Long = &H100

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type FNCPOINT
    X                           As Double
    y                           As Double
    PxX                         As Double
    PxY                         As Double
    m                           As Double
    c                           As Double
    IsValid                     As Boolean
End Type

Private m_dblMaxX               As Double
Private m_dblMaxY               As Double
Private m_dblMinX               As Double
Private m_dblMinY               As Double

Private m_hGDIToken             As Long
Private m_hGDIGraphics          As Long
Private m_hGDIPen               As Long

Private m_dblStepsX             As Double
Private m_dblStepsY             As Double

Private m_hDCBack               As Long
Private m_lngDCBackBmp          As Long
Private m_lngDCBackOldBmp       As Long

Private Const POINTS_PER_PIXEL  As Long = 10

Private Const LE_LINE           As Long = 7   ' pixel

Private Function LEtoPixel_X(ByVal drw_width As Long, ByVal X As Double) As Double
    LEtoPixel_X = (X - m_dblMinX) / ((m_dblMaxX - m_dblMinX) / drw_width)
End Function

Private Function LEtoPixel_Y(ByVal drw_height As Long, ByVal y As Double) As Double
    LEtoPixel_Y = (y - m_dblMinY) / ((m_dblMaxY - m_dblMinY) / drw_height)
End Function

Private Function PixelToLE_X(ByVal drw_width As Long, ByVal X As Double) As Double
    PixelToLE_X = X * ((m_dblMaxX - m_dblMinX) / drw_width) + m_dblMinX
End Function

Private Function PixelToLE_Y(ByVal drw_height As Long, ByVal y As Double) As Double
    PixelToLE_Y = y * ((m_dblMaxY - m_dblMinY) / drw_height) + m_dblMinY
End Function

Private Sub DrawNet()
    Dim i       As Double
    Dim dyBox   As Double
    Dim dxBox   As Double
    Dim dy      As Long
    Dim dx      As Long
    Dim strText As String
    Dim rcText  As RECT
    
    With picGraph
        dy = .ScaleHeight
        dx = .ScaleWidth
    End With
    
    dyBox = LEtoPixel_Y(dy, m_dblStepsY) - LEtoPixel_Y(dy, 0)
    dxBox = LEtoPixel_X(dx, m_dblStepsX) - LEtoPixel_X(dx, 0)
    
    ' Y lines
    picGraph.DrawStyle = vbDot
    picGraph.ForeColor = &HC0C0C0
    
    i = dy - LEtoPixel_Y(dy, 0)
    Do
        picGraph.Line (0, i)-(dx, i)
        i = i - dyBox
        
        strText = CStr(Round(PixelToLE_Y(dy, dy - i), 1))
        
        With rcText
            .Top = i - 2
            .Left = LEtoPixel_X(dx, 0) + 2
            .Right = .Left + picGraph.TextWidth(strText)
            .Bottom = .Top + picGraph.TextHeight(strText)
        End With
        
        DrawText picGraph.hdc, strText, Len(strText), rcText, 0
    Loop While i > 0
    
    i = dy - LEtoPixel_Y(dy, 0) + dyBox
    Do
        picGraph.Line (0, i)-(dx, i)
        
        strText = CStr(Round(PixelToLE_Y(dy, dy - i), 1))
        
        With rcText
            .Top = i
            .Left = LEtoPixel_X(dx, 0) + 2
            .Right = .Left + picGraph.TextWidth(strText)
            .Bottom = .Top + picGraph.TextHeight(strText)
        End With
        
        DrawText picGraph.hdc, strText, Len(strText), rcText, 0
        i = i + dyBox
    Loop While i <= dy
    
    ' X lines
    i = LEtoPixel_X(dx, 0)
    Do
        picGraph.Line (i, 0)-(i, dy)
        i = i - dxBox
        
        strText = CStr(Round(PixelToLE_X(dx, i), 1))
        
        With rcText
            .Top = dy - LEtoPixel_Y(dy, 0) + 2
            .Left = i
            .Right = .Left + picGraph.TextWidth(strText)
            .Bottom = .Top + picGraph.TextHeight(strText)
        End With
        
        DrawText picGraph.hdc, strText, Len(strText), rcText, 0
    Loop While i > 0
    
    i = LEtoPixel_X(dx, 0) + dxBox
    Do
        picGraph.Line (i, 0)-(i, dy)
        
        strText = CStr(Round(PixelToLE_X(dx, i), 1))
        
        With rcText
            .Top = dy - LEtoPixel_Y(dy, 0) + 2
            .Left = i - picGraph.TextHeight(strText) / 2
            .Right = .Left + picGraph.TextWidth(strText)
            .Bottom = .Top + picGraph.TextHeight(strText)
        End With
        
        DrawText picGraph.hdc, strText, Len(strText), rcText, 0
        
        i = i + dxBox
    Loop While i <= dx
    
    ' axes
    picGraph.DrawStyle = vbSolid
    picGraph.ForeColor = &H606060
    
    picGraph.Line (LEtoPixel_X(dx, 0), 0)-(LEtoPixel_X(dx, 0), dy)
    picGraph.Line (0, dy - LEtoPixel_Y(dy, 0))-(dx, dy - LEtoPixel_Y(dy, 0))
    
    BitBlt m_hDCBack, 0, 0, dx, dy, picGraph.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub DrawGraph()
    Dim udtDrawPoints() As POINTF
    Dim udtPoints()     As FNCPOINT
    Dim dx              As Long
    Dim dy              As Long
    Dim xidx            As Long
    Dim lngLastUsed     As Long
    Dim X               As Double
    Dim k               As Double

    dx = picGraph.ScaleWidth
    dy = picGraph.ScaleHeight
      
    xidx = modFormula.GetVariableIndex("x")
    modFormula.ExpressionString = txtFnc.Text
    If Not modFormula.ValidateExpression() Then
        MsgBox "syntax error in expression: " & modFormula.ErrorMessage
        Exit Sub
    End If

    ReDim udtDrawPoints(dx * POINTS_PER_PIXEL - 1) As POINTF
    ReDim udtPoints(dx * POINTS_PER_PIXEL - 1) As FNCPOINT
    
    For X = 0 To dx Step (1 / POINTS_PER_PIXEL)
        modFormula.VariableValue(xidx) = PixelToLE_X(dx, X)
        
        With udtPoints(k)
            .X = PixelToLE_X(dx, X)
            .y = modFormula.Evaluate()
            .IsValid = Not modFormula.CalculationError
            
            If .IsValid Then
                .PxX = X
                .PxY = dy - LEtoPixel_Y(dy, .y)
                
                ' slope in the point
                If X > 0 Then
                    If udtPoints(k - 1).IsValid Then
                        .m = (.y - udtPoints(k - 1).y) / (.X - udtPoints(k - 1).X)
                        .c = .y - .m * .X
                    Else
                        .m = .y
                        .c = .y
                    End If
                Else
                    .m = .y
                    .c = .y
                End If
                
                If .PxY < 0 Then
                    .PxY = -1
                ElseIf .PxY > dy Then
                    .PxY = dy + 1
                End If
            End If
        End With
        
        k = k + 1
    Next
    
    ' draw the points with GDI+
    '
    ' if there are differences between points which can cause drawing errors,
    ' draw till the first point and skip the next one
    For k = 0 To dx * POINTS_PER_PIXEL - 1
        With udtPoints(k)
            udtDrawPoints(k).X = .PxX
            udtDrawPoints(k).y = .PxY
            
            If Not .IsValid Then
                GdipDrawCurve m_hGDIGraphics, m_hGDIPen, udtDrawPoints(lngLastUsed), k - lngLastUsed - 1
                lngLastUsed = k + 1
            ElseIf k > 0 Then
                If udtPoints(k - 1).IsValid Then
                    If udtPoints(k - 1).y <= m_dblMinY And udtPoints(k - 1).m < 0 Then
                        ' change from -infinity to +infinity
                        If .y >= m_dblMaxY Then
                            GdipDrawCurve m_hGDIGraphics, m_hGDIPen, udtDrawPoints(lngLastUsed), k - lngLastUsed - 1
                            lngLastUsed = k + 1
                        End If
                    ElseIf udtPoints(k - 1).y >= m_dblMaxY And udtPoints(k - 1).m > 0 Then
                        ' change from +infinity to -infinity
                        If .y <= m_dblMinY Then
                            GdipDrawCurve m_hGDIGraphics, m_hGDIPen, udtDrawPoints(lngLastUsed), k - lngLastUsed - 1
                            lngLastUsed = k + 1
                        End If
                    ElseIf Abs(.m - udtPoints(k - 1).m) >= 200 Then
                        ' extreme slope in a short very distance
                        GdipDrawCurve m_hGDIGraphics, m_hGDIPen, udtDrawPoints(lngLastUsed), k - lngLastUsed
                        If .PxY < udtPoints(k - 1).PxY Then
                            GdipDrawLine m_hGDIGraphics, m_hGDIPen, .PxX, udtPoints(k - 1).PxY - 1, .PxX, .PxY
                        Else
                            GdipDrawLine m_hGDIGraphics, m_hGDIPen, .PxX, udtPoints(k - 1).PxY, .PxX, .PxY
                        End If
                        lngLastUsed = k
                    End If
                End If
            End If
        End With
    Next
    
    If lngLastUsed < k Then
        GdipDrawCurve m_hGDIGraphics, m_hGDIPen, udtDrawPoints(lngLastUsed), k - lngLastUsed - 1
    End If
End Sub

Private Sub chkAntialias_Click()
    If chkAntialias.value Then
        GdipSetSmoothingMode m_hGDIGraphics, SmoothingModeAntiAlias
    Else
        GdipSetSmoothingMode m_hGDIGraphics, SmoothingModeHighSpeed
    End If
End Sub

Private Sub cmdPlot_Click()
    Dim btR As Byte, btG As Byte, btB As Byte

    m_dblMaxX = Val(txtMaxX.Text)
    m_dblMaxY = Val(txtMaxY.Text)
    m_dblMinX = Val(txtMinX.Text)
    m_dblMinY = Val(txtMinY.Text)
    m_dblStepsX = Val(txtXSteps.Text)
    m_dblStepsY = Val(txtYSteps.Text)
    
    If txtFnc.Text = "" Then
        MsgBox "function missing!", vbExclamation
    Else
        btR = picColor.BackColor And &HFF
        btG = picColor.BackColor \ &H100 And &HFF
        btB = picColor.BackColor \ &H10000 And &HFF
        
        GdipSetPenColor m_hGDIPen, &HFF000000 Or RGB(btB, btG, btR)
        
        With picGraph
            .Cls
            DrawNet
            DrawGraph
            BitBlt picGraph.hdc, 0, 0, picGraph.ScaleWidth, picGraph.ScaleHeight, m_hDCBack, 0, 0, vbSrcCopy
        End With
    End If
End Sub

Private Sub Form_Load()
    modFormula.Formula_Init
    modFormula.AddVariable 0, "x"
    
    With picGraph
        m_hDCBack = CreateCompatibleDC(picGraph.hdc)
        m_lngDCBackBmp = CreateCompatibleBitmap(picGraph.hdc, .ScaleWidth, .ScaleHeight)
        m_lngDCBackOldBmp = SelectObject(m_hDCBack, m_lngDCBackBmp)
    End With
    
    If Not InitGdip() Then
        MsgBox "Could not init GDI+!"
    Else
        GdipCreateFromHDC m_hDCBack, m_hGDIGraphics
        GdipCreatePen1 Linen, 1, UnitPixel, m_hGDIPen
        GdipSetPenColor m_hGDIPen, &HFF9050FF
        picColor.BackColor = &HFF5090
    
        GdipSetCompositingQuality m_hGDIGraphics, CompositingQualityHighSpeed
        
        If chkAntialias.value Then
            GdipSetSmoothingMode m_hGDIGraphics, SmoothingModeAntiAlias
        Else
            GdipSetSmoothingMode m_hGDIGraphics, SmoothingModeHighSpeed
        End If
    End If
    
    m_dblStepsX = 1
    m_dblStepsY = 1
    m_dblMaxX = 10
    m_dblMaxY = 10
    m_dblMinX = -10
    m_dblMinY = -10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modFormula.Formula_Terminate
    
    GdipDeletePen m_hGDIPen
    GdipDeleteGraphics m_hGDIGraphics
    
    ShutdownGdip
    
    SelectObject m_hDCBack, m_lngDCBackOldBmp
    DeleteDC m_hDCBack
End Sub

Private Sub picColor_Click()
    dlg.ShowColor
    picColor.BackColor = dlg.Color
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblX.Caption = "X: " & Round(PixelToLE_X(picGraph.ScaleWidth - 1, X), 2)
    modFormula.VariableValue(modFormula.GetVariableIndex("x")) = PixelToLE_X(picGraph.ScaleWidth - 1, X)
    lblY.Caption = "Y: " & Round(modFormula.Evaluate(), 2)
End Sub

Private Sub txtMaxX_Validate(Cancel As Boolean)
    m_dblMaxX = Val(Replace(txtMaxX.Text, ",", "."))
End Sub

Private Sub txtMaxY_Validate(Cancel As Boolean)
    m_dblMaxY = Val(Replace(txtMaxY.Text, ",", "."))
End Sub

Private Sub txtMinX_Validate(Cancel As Boolean)
    m_dblMinX = Val(Replace(txtMinX.Text, ",", "."))
End Sub

Private Sub txtMinY_Validate(Cancel As Boolean)
    m_dblMinY = Val(Replace(txtMinY.Text, ",", "."))
End Sub

Private Sub txtXSteps_Validate(Cancel As Boolean)
    m_dblStepsX = Val(Replace(txtXSteps.Text, ",", "."))
End Sub

Private Sub txtYSteps_Validate(Cancel As Boolean)
    m_dblStepsY = Val(Replace(txtYSteps.Text, ",", "."))
End Sub

Private Function InitGdip() As Boolean
    Dim uInput  As GdiplusStartupInput

    If m_hGDIToken = 0 Then
        uInput.GdiplusVersion = 1
        InitGdip = GdiplusStartup(m_hGDIToken, uInput) = Ok
    End If
End Function

Private Function ShutdownGdip() As Boolean
    If m_hGDIToken <> 0 Then
        GdiplusShutdown m_hGDIToken
        m_hGDIToken = 0
        ShutdownGdip = True
    End If
End Function
