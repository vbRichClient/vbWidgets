VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click()
Event DblClick()
Event MouseEnter()
Event MouseLeave()
Event MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Event MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Event MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)

'--- Variables used for the QB-DrawString-Implementation
Private sDraw$, Pos&, Cmd$, Args(1) As String, x#, y#, Scl#, Angles(), Agl#, AglX#, AglY#
Private IsAbs As Boolean, NoCoordUpdates As Boolean, StrokeColor&, FillColor&, RevFillColor&
 
Private mCanvasSrf As cCairoSurface, mCanvas As cCairoContext, SVG As cSVG, SVGFileName As String
Private mLockedBytes As Boolean, BGRA() As Byte, PxlL() As Long

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Angles = Array(0, 90, 180, 270) 'used for the QB-DrawString-Implementation
  Set SVG = New_c.SVG
  Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
  W.CanGetFocus = False
End Sub

Public Property Get Widget() As cWidgetBase
  Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = W.Widgets
End Property
 
Public Sub SetCanvasSize(ByVal Width As Long, ByVal Height As Long)
  If Not mCanvasSrf Is Nothing Then UnLockCanvas
  Set mCanvasSrf = Cairo.CreateWin32Surface(Width, Height)
End Sub

Public Property Get hDC() As Long
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  hDC = mCanvasSrf.GetDC
End Property

Public Property Get CanvasWidth() As Long
  If Not mCanvasSrf Is Nothing Then CanvasWidth = mCanvasSrf.Width
End Property
Public Property Get CanvasHeight() As Long
  If Not mCanvasSrf Is Nothing Then CanvasHeight = mCanvasSrf.Height
End Property

Public Property Get Canvas() As cCairoContext
Attribute Canvas.VB_UserMemId = 0
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If mCanvas Is Nothing Then Set mCanvas = mCanvasSrf.CreateContext
  Set Canvas = mCanvas
End Property

Public Sub Cls(Optional ByVal BackColor As Long = -1)
  If mCanvasSrf Is Nothing Then Exit Sub
  If BackColor = -1 Then Canvas.Operator = CAIRO_OPERATOR_CLEAR Else Canvas.SetSourceColor BackColor
  Canvas.Paint
  Canvas.Operator = CAIRO_OPERATOR_OVER
End Sub
Public Sub Refresh()
  W.Refresh
End Sub

Public Sub LockCanvas(Optional ByVal DeMultiplyAlpha As Boolean = True)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If DeMultiplyAlpha Then Cairo.DeMultiplyAlpha mCanvasSrf.DataPtr, mCanvasSrf.Stride * mCanvasSrf.Height
  mCanvasSrf.BindToArray BGRA
  mCanvasSrf.BindToArrayLong PxlL
  mLockedBytes = True
End Sub
Public Sub UnLockCanvas(Optional ByVal PreMultiplyAlpha As Boolean = True)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If Not mLockedBytes Then Exit Sub
  mCanvasSrf.ReleaseArray BGRA
  mCanvasSrf.ReleaseArrayLong PxlL
  If PreMultiplyAlpha Then Cairo.PreMultiplyAlpha mCanvasSrf.DataPtr, mCanvasSrf.Stride * mCanvasSrf.Height
  mCanvas.Save
  mCanvas.DrawLine 0, 0, 1, 1, , 0.000001, 0, 0.02 'force a Dummy-Drawing-Op
  mCanvas.Restore
  mLockedBytes = False
End Sub

Public Sub SetPixelRGBA(ByVal x As Long, ByVal y As Long, r, G, B, A)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If Not mLockedBytes Then Err.Raise vbObjectError, , "the Canvas needs to be locked, to use this method"
  x = x * 4
  BGRA(x, y) = B: BGRA(x + 1, y) = G: BGRA(x + 2, y) = r: BGRA(x + 3, y) = A
End Sub
Public Sub GetPixelRGBA(ByVal x As Long, ByVal y As Long, r, G, B, A)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If Not mLockedBytes Then Err.Raise vbObjectError, , "the Canvas needs to be locked, to use this method"
  x = x * 4
  B = BGRA(x, y): G = BGRA(x + 1, y): r = BGRA(x + 2, y): A = BGRA(x + 3, y)
End Sub

Public Sub SetPixel(ByVal x As Long, ByVal y As Long, ByVal BGR24Bit As Long, Optional ByVal AlphaByte As Byte = 255)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If Not mLockedBytes Then Err.Raise vbObjectError, , "the Canvas needs to be locked, to use this method"
  PxlL(x, y) = BGR24Bit Or (&H1000000 * AlphaByte)
End Sub
Public Function GetPixel(ByVal x As Long, ByVal y As Long, Optional ByVal IncludeAlphaByte As Boolean) As Long
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  If Not mLockedBytes Then Err.Raise vbObjectError, , "the Canvas needs to be locked, to use this method"
  If IncludeAlphaByte Then GetPixel = PxlL(x, y) Else GetPixel = PxlL(x, y) And Not &HFF000000
End Function
Public Sub FloodFill(ByVal x As Long, ByVal y As Long, ByVal NewColor As Long)
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  Dim Pxl() As Long
  If Not mCanvasSrf.BindToArrayLong(Pxl) Then Exit Sub
  RecFloodFill Pxl, x, y, NewColor
  mCanvasSrf.ReleaseArrayLong Pxl
End Sub

Private Sub RecFloodFill(Pxl() As Long, ByVal x&, ByVal y&, ByVal NewColor&)
Dim xL&, xR&, OldColor&, UBx&, UBy&
  UBx = UBound(Pxl, 1): UBy = UBound(Pxl, 2)
  If x < 0 Or x > UBx Or y < 0 Or y > UBy Then Exit Sub
  If Pxl(x, y) = NewColor Then Exit Sub Else OldColor = Pxl(x, y)

  For xL = x To 1 Step -1 'find most left pixel
    If Pxl(xL - 1, y) <> OldColor Then Exit For
  Next xL
  For xR = x To UBx - 1 'find most right pixel
    If Pxl(xR + 1, y) <> OldColor Then Exit For
  Next xR
  For x = xL To xR: Pxl(x, y) = NewColor: Next 'fill this part

  'Recursion only in y-direction
  If y > 0 Then 'check scanline above
    For x = xL To xR
      If Pxl(x, y - 1) = OldColor Then RecFloodFill Pxl, x, y - 1, NewColor
    Next
  End If
  If y < UBy Then 'check scanline below
    For x = xL To xR
      If Pxl(x, y + 1) = OldColor Then RecFloodFill Pxl, x, y + 1, NewColor
    Next
  End If
End Sub

Public Sub SetImage(ImageSrfOrFileName As Variant, Optional ByVal BufferSVGs As Boolean, Optional ByVal SVGWidth As Long, Optional ByVal SVGHeight As Long)
    If Not mCanvasSrf Is Nothing Then UnLockCanvas
    SVGFileName = "": Set mCanvasSrf = Nothing: Set mCanvas = Nothing
    If VarType(ImageSrfOrFileName) = vbString Then
        If InStr(ImageSrfOrFileName, ".svg") Then
           If SVGFileName <> ImageSrfOrFileName Then 'this ensures, that the Parsing-Part of reading an SVG is done only on FileName-Changes
              SVGFileName = ImageSrfOrFileName
              SVG.ParseContent SVGFileName
              If BufferSVGs Then
                If SVGWidth > 0 And SVGHeight > 0 Then
                  Set mCanvasSrf = Cairo.CreateSurface(SVGWidth + 2, SVGHeight + 2)
                Else
                  Set mCanvasSrf = Cairo.CreateSurface(W.Width + 2, W.Height + 2)
                End If
                SVG.RenderFromDOM mCanvasSrf.CreateContext, 0, 0, mCanvasSrf.Width, mCanvasSrf.Height
              End If
           End If
        ElseIf SVGWidth > 0 And SVGHeight > 0 Then
          Set mCanvasSrf = Cairo.ImageList.AddImage("", ImageSrfOrFileName, SVGWidth, SVGHeight)
        Else
          Set mCanvasSrf = Cairo.CreateSurface(0, 0, ImageSurface, ImageSrfOrFileName)
        End If
    ElseIf VarType(ImageSrfOrFileName) = vbByte Or vbArray Then
      If SVGWidth > 0 And SVGHeight > 0 Then
        Set mCanvasSrf = Cairo.ImageList.AddImage("", ImageSrfOrFileName, SVGWidth, SVGHeight)
      Else
        Set mCanvasSrf = Cairo.CreateSurface(0, 0, ImageSurface, ImageSrfOrFileName)
      End If
    ElseIf IsObject(ImageSrfOrFileName) Then
      Set mCanvasSrf = ImageSrfOrFileName
    End If
End Sub

Private Sub W_Click()
  RaiseEvent Click
End Sub
Private Sub W_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
  RaiseEvent MouseEnter
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
  RaiseEvent MouseLeave
End Sub
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not mCanvasSrf Is Nothing Then x = x * mCanvasSrf.Width / W.Width: y = y * mCanvasSrf.Height / W.Height
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not mCanvasSrf Is Nothing Then x = x * mCanvasSrf.Width / W.Width: y = y * mCanvasSrf.Height / W.Height
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not mCanvasSrf Is Nothing Then x = x * mCanvasSrf.Width / W.Width: y = y * mCanvasSrf.Height / W.Height
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
 
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  If Len(W.ImageKey) Then CC.RenderSurfaceContent W.ImageKey, 0, 0, dx_Aligned, dy_Aligned

  If Len(SVGFileName) Then
    If mCanvasSrf Is Nothing Then 'unbuffered SVG-mode (slower, but accurate in any zoom-level and size)
      SVG.RenderFromDOM CC, 0, 0, dx_Aligned, dy_Aligned
    Else 'SVG-buffered mode, with one pixel overscan on all edges
      CC.RenderSurfaceContent mCanvasSrf, -1, -1, dx_Aligned + 2, dy_Aligned + 2
    End If
  ElseIf Not mCanvasSrf Is Nothing Then
    CC.RenderSurfaceContent mCanvasSrf, 0, 0, dx_Aligned, dy_Aligned
  End If
End Sub
 
Private Sub Class_Terminate()
  If Not mCanvasSrf Is Nothing Then UnLockCanvas
  Set mCanvas = Nothing
  Set mCanvasSrf = Nothing
End Sub


'*** QB-DrawString-Support, using the publically exposed routine below (and a few Private Helper-Functions further down)
Public Sub Draw(ByVal sDrawString, Optional ByVal UseFloodFill As Boolean, Optional ByVal UseQBColors As Boolean)
Dim Pxl() As Long, sx#, sy#, xx!, yy!
  If mCanvasSrf Is Nothing Then Err.Raise vbObjectError, , "no Canvas exists currently, to create one - use SetCanvasSize(Width, Height)"
  
  If Len(sDrawString) = 0 Then x = 0: y = 0: Agl = 0: AglX = 0: AglY = 0: Scl = 1: SetScaleAndAngle: Exit Sub
  If Scl = 0 Then
    Scl = 1: SetScaleAndAngle
    Canvas.SetLineWidth 1
    Canvas.SetSourceColor vbBlack
    Canvas.FillRule = CAIRO_FILL_RULE_EVEN_ODD
  End If
  sDraw = Replace$(Replace$(Replace$(Trim$(UCase$(sDrawString)), " ", ""), "+-", "-"), "-+", "-")
  Pos = 0
  Canvas.AntiAlias = IIf(UseFloodFill, CAIRO_ANTIALIAS_NONE, CAIRO_ANTIALIAS_DEFAULT)
  Do
    NoCoordUpdates = False
    ParseCommand
    If Len(Cmd) Then ParseArgs UseQBColors Else Exit Do
    Select Case Cmd
      Case "C"
        mCanvas.Stroke
        mCanvas.MoveTo x, y
        mCanvas.SetSourceColor StrokeColor
      Case "BW"
        mCanvas.SetLineWidth Args(0)
      Case "BM"
        mCanvas.MoveTo Args(0), Args(1)
        If Not NoCoordUpdates Then x = Args(0): y = Args(1)
        If IsAbs Then AglX = x: AglY = y: SetScaleAndAngle
      Case "M"
        mCanvas.LineTo Args(0), Args(1)
        If NoCoordUpdates Then mCanvas.MoveTo x, y Else x = Args(0): y = Args(1)
      Case "P"
        If UseFloodFill Then
          mCanvas.Save
            mCanvas.Matrix.CalculateScale sx, sy: If mCanvas.GetLineWidth * sx < 1 Then mCanvas.SetLineWidth 1 / sx
            mCanvas.SetSourceColor StrokeColor: mCanvas.Stroke
          mCanvas.Restore
          xx = x: yy = y: mCanvas.Matrix.CalculatePointSingle xx, yy
          FloodFill xx, yy, RevFillColor
        Else
          mCanvas.SetSourceColor FillColor:   mCanvas.Fill True
          mCanvas.SetSourceColor StrokeColor: mCanvas.Stroke
        End If
      Case "S", "TA"
        If Cmd = "S" Then Scl = Args(0) / 4 Else Agl = Args(0): AglX = x: AglY = y
        SetScaleAndAngle
    End Select
  Loop
  mCanvas.SetSourceColor StrokeColor: mCanvas.Stroke
  mCanvas.AntiAlias = CAIRO_ANTIALIAS_DEFAULT
End Sub
 
Private Sub ParseCommand()
  Cmd = ""
  For Pos = Pos + 1 To Len(sDraw)
    Cmd = Mid$(sDraw, Pos, 1)
    If Cmd = "B" And Mid$(sDraw, Pos + 1, 1) = "B" Then Pos = Pos + 1 ': mCanvas.Stroke: mCanvas.ClearPath
    Select Case Cmd
      Case "A", "C", "M", "L", "R", "U", "D", "E", "F", "G", "H", "S", "P": Exit Sub
      Case "B", "T": Pos = Pos + 1: Cmd = Cmd & Mid(sDraw, Pos, 1): Exit Sub
      Case "N": NoCoordUpdates = True: ParseCommand: Exit Sub
    End Select
  Next
End Sub
 
Private Sub ParseArgs(UseQBColors As Boolean)
  Dim i As Long, A() As String, BPrefix As Boolean
  For i = Pos + 1 To Len(sDraw)
    If Mid$(sDraw, i, 1) >= "A" Then Exit For
  Next
  
  A = Split(Trim$(Mid$(sDraw, Pos + 1, i - Pos - 1)), ",")
  ReDim Preserve A(1)
  Pos = i - 1
  IsAbs = True
  For i = 0 To UBound(A)
    A(i) = Trim$(A(i))
    Args(i) = Val(IIf(Len(A(i)), A(i), 1))
    If Left(A(i), 1) = "-" Or Left(A(i), 1) = "+" Or IsAbs = False Then
      Args(i) = IIf(i, y, x) + Args(i): IsAbs = False
    End If
  Next
  
  If Left$(Cmd, 1) = "B" Then Cmd = Right$(Cmd, 1): BPrefix = True
  Select Case Cmd
    Case "L": Cmd = "M":  Args(1) = y:           Args(0) = x - Args(0)
    Case "R": Cmd = "M":  Args(1) = y:           Args(0) = x + Args(0)
    Case "U": Cmd = "M":  Args(1) = y - Args(0): Args(0) = x
    Case "D": Cmd = "M":  Args(1) = y + Args(0): Args(0) = x
    Case "E": Cmd = "M":  Args(1) = y - Args(0): Args(0) = x + Args(0)
    Case "F": Cmd = "M":  Args(1) = y + Args(0): Args(0) = x + Args(0)
    Case "G": Cmd = "M":  Args(1) = y + Args(0): Args(0) = x - Args(0)
    Case "H": Cmd = "M":  Args(1) = y - Args(0): Args(0) = x - Args(0)
    Case "A": Cmd = "TA": Args(0) = Angles(Args(0))
    Case "C": StrokeColor = GetColor(Args(0), UseQBColors): FillColor = -1
    Case "P": StrokeColor = GetColor(Args(1), UseQBColors): FillColor = GetColor(Args(0), UseQBColors, RevFillColor)
    Case "TA": If Not IsAbs Then Args(0) = Args(0) - x
  End Select
  If BPrefix Then Cmd = "B" & Cmd
End Sub

Private Function GetColor(ByVal C As Currency, UseQBColors As Boolean, Optional ColorReverse As Long)
  If C > 2147483648@ Then C = C - 2147483648@
  If UseQBColors Then GetColor = QBColor(C And &HF&): Exit Function
  GetColor = RGB((C \ 65536) And &HFF, (C \ CLng(256)) And &HFF, C And &HFF)
  ColorReverse = RGB(C And &HFF, (C \ CLng(256)) And &HFF, (C \ 65536) And &HFF)
End Function

Private Sub SetScaleAndAngle()
  mCanvas.MatrixResetToIdentity
  mCanvas.TranslateDrawings AglX + 0.5, AglY + 0.5
    mCanvas.ScaleDrawings Scl, Scl
    mCanvas.RotateDrawingsDeg -Agl
  mCanvas.TranslateDrawings -AglX, -AglY
End Sub
'********** End of QB-Drawstring-Support *************
