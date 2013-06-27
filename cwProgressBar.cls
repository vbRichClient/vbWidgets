VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eProgressStyle
    pg_Standard = 0
    pg_MarqueeLoop = 1
    pg_MarqueeBounce = 2
End Enum

Private mInternalW As Single, mInternalH As Single
Private dx As Single, dy As Single, Alpha As Single

Private mStyle As eProgressStyle
Private mMin As Single, mMax As Single, mValue As Single
Private mCaption As String
Private mBorderRadius As Double
 
Const MarqueeWidthPct As Double = 0.25
Private WithEvents MarqueeTimer As cTimer, TaskBar As cTaskBar
Attribute MarqueeTimer.VB_VarHelpID = -1

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set W = Cairo.WidgetBase
    W.CanGetFocus = False
    W.SelectionColor = vbGreen
    mBorderRadius = 2.5
    mMax = 100
End Sub
Public Property Get Widget() As cWidgetBase
    Set Widget = W
End Property
Public Property Get Widgets() As cWidgets
    Set Widgets = W.Widgets
End Property

'after the necessities above, I start usually with all the Public Props
Public Sub Clear()
    mCaption = vbNullString
    mValue = mMin
    W.Refresh
    If Not TaskBar Is Nothing And Not W.Root Is Nothing Then
        TaskBar.SetProgressState W.Root.DialogFormHwnd, TBPF_NOPROGRESS
    End If
End Sub

Public Property Get Value() As Single
  Value = mValue
End Property
Public Property Let Value(ByVal NewValue As Single)
    If NewValue > mMax Then
        mValue = mMax
    ElseIf NewValue < mMin Then
        mValue = mMin
    Else
        mValue = NewValue
    End If
    W.Refresh
    If Not TaskBar Is Nothing And Not W.Root Is Nothing Then
        TaskBar.SetProgressState W.Root.DialogFormHwnd, TBPF_NORMAL
        TaskBar.SetProgressValue W.Root.DialogFormHwnd, mValue - mMin, mMax - mMin
    End If
End Property

Public Property Get Min() As Single
  Min = mMin
End Property
Public Property Let Min(ByVal NewValue As Single)
  mMin = NewValue
  If mMin > mValue Then mValue = mMin
End Property

Public Property Get Max() As Single
  Max = mMax
End Property
Public Property Let Max(ByVal NewValue As Single)
  mMax = NewValue
  If mMax < mValue Then mValue = mMax
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Caption(NewValue As String)
    If mCaption = NewValue Then Exit Property
    mCaption = NewValue
    W.Refresh
End Property

Public Property Get BorderRadius() As Double
  BorderRadius = mBorderRadius
End Property
Public Property Let BorderRadius(ByVal NewValue As Double)
  If NewValue < 0 Then NewValue = 0
  If NewValue = mBorderRadius Then Exit Property
  mBorderRadius = NewValue
  W.Refresh
End Property

Public Property Get ReflectInTaskBar() As Boolean
  ReflectInTaskBar = Not TaskBar Is Nothing
End Property
Public Property Let ReflectInTaskBar(ByVal NewValue As Boolean)
  If ReflectInTaskBar = NewValue Then Exit Property
  If NewValue Then Set TaskBar = New_c.Cairo.TaskBar Else Set TaskBar = Nothing
End Property
Public Property Get Style() As eProgressStyle
    Style = mStyle
End Property
Public Property Let Style(ByVal NewValue As eProgressStyle)
    If mStyle = NewValue Then Exit Property
    
    mStyle = NewValue
    If mStyle = pg_Standard Then
        Set MarqueeTimer = Nothing 'let's not waste resources, if we don't need to in the standard-case
    Else
        mValue = 0
        mMin = 0
        mMax = 100
        Set MarqueeTimer = New_c.Timer(15, True, 1)
    End If
    
    W.Refresh
    If Not TaskBar Is Nothing And Not W.Root Is Nothing Then
        TaskBar.SetProgressState W.Root.DialogFormHwnd, IIf(mStyle = pg_Standard, TBPF_NOPROGRESS, TBPF_INDETERMINATE)
    End If
End Property
'after the Public Props, the internal Event-Handlers
Private Sub MarqueeTimer_Timer()
    mValue = mValue + MarqueeTimer.Tag
    
    Select Case True
        Case mStyle = pg_MarqueeLoop
            If mValue > (mMax * (1 + MarqueeWidthPct)) Then mValue = mMin
 
        Case mStyle = pg_MarqueeBounce
            If mValue > mMax Then MarqueeTimer.Tag = -1
            If mValue < (mMax * MarqueeWidthPct) Then MarqueeTimer.Tag = 1
    End Select
        
    W.Refresh
End Sub

'with the Paint-Event usually as the last one...
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    dx = dx_Aligned: dy = dy_Aligned
    Alpha = W.AlphaInherited
    
    If dx >= dy Then 'normal, horizontal Bar
        mInternalW = dx: mInternalH = dy
        Draw CC
    Else 'switch to autodetected, vertical Orientation
        mInternalW = dy: mInternalH = dx
        CC.Save
          CC.RotateDrawingsDeg 270
          CC.TranslateDrawings -dy, 0
          Draw CC
        CC.Restore
    End If
End Sub

'and the related Drawing-Cascade in the right order...
Private Sub Draw(CC As cCairoContext)
    DrawBackground CC, 0.6, 0, 1.05, 0.3, 0.99, 0.3, 0.95, 1, 0.9
    DrawBorder CC, 0.85
    DrawUsedArea CC
    DrawCaption CC
    DrawShine CC
End Sub
Private Sub DrawBackground(CC As cCairoContext, ByVal WhiteStrokeAlpha As Double, ParamArray StopsAndShade())
Dim i&
    If W.BackColor = -1 Then Exit Sub 'no BackColor, so we leave here (and the BackGround transparent)
    
    CC.SetLineWidth 2, True
    CC.RoundedRect 0, 0, mInternalW, mInternalH, mBorderRadius, True
    With Cairo.CreateLinearPattern(0, 0, 0, mInternalH)
    
        For i = 0 To UBound(StopsAndShade) Step 2 'on the even indices are the Stop-Values - and on the uneven ones the Shade-Values
            .AddColorStop StopsAndShade(i), W.BackColor, Alpha, StopsAndShade(i + 1)
        Next i

        CC.Fill True, .This
        
        CC.SetSourceColor vbWhite, WhiteStrokeAlpha * Alpha
        CC.Stroke
    End With
End Sub
Private Sub DrawBorder(CC As cCairoContext, ByVal BorderAlpha As Double)
    CC.SetLineWidth 1, True
    CC.RoundedRect 0, 0, mInternalW, mInternalH, mBorderRadius, True
    CC.SetSourceColor W.BorderColor, BorderAlpha * Alpha
    CC.Stroke
End Sub
Private Sub DrawUsedArea(CC As cCairoContext)
Dim ScaledValue As Double, StartX As Double
 
    CC.SetLineWidth 0
    ScaledValue = ((mValue - mMin) / (mMax - mMin)) * (mInternalW - 2.5)
    
    If mStyle <> pg_Standard Then
      StartX = ScaledValue - (mInternalW * MarqueeWidthPct)
    End If

    Cairo.Theme.DrawTo CC, W, thmTypeProgressFace, thmStateSolidColor, StartX, 0, ScaledValue - StartX, mInternalH, mBorderRadius
End Sub
Private Sub DrawCaption(CC As cCairoContext)
  If Len(mCaption) = 0 Then Exit Sub
  W.SelectFontSettingsInto CC
  CC.DrawText 0.05, 0.55, mInternalW, mInternalH, mCaption, True, vbCenter, , 1
End Sub
Private Sub DrawShine(CC As cCairoContext)
  Cairo.Theme.DrawTo CC, W, thmTypeShine, 0, 1, 1, mInternalW - 2, mInternalH \ 4 + 1, 1
  Cairo.Theme.DrawTo CC, W, thmTypeShine, 0, 1, 1, mInternalW - 2, mInternalH \ 4 + 1, 1
  Cairo.Theme.DrawTo CC, W, thmTypeShine, 0, 2, 1, mInternalW - 4, mInternalH \ 8 + 1, 1
End Sub
