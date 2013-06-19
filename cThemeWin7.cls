VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cThemeWin7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Implements cTheme

Private mThemeColorArray(0 To 6) As Long, mScrollerSize As Long

Private Sub Class_Initialize() 'fill up some Values here, for faster lookup later on
  mThemeColorArray(thmBackColor) = RGB(235, 235, 235) 'RGB(241, 241, 241)
  mThemeColorArray(thmForeColor) = RGB(30, 30, 30)
  mThemeColorArray(thmBorderColor) = RGB(103, 103, 98)
  mThemeColorArray(thmHoverColor) = RGB(190, 228, 255)  'RGB(190, 238, 255)
  mThemeColorArray(thmFocusColor) = RGB(0, 140, 195)
  mThemeColorArray(thmSelectionColor) = RGB(178, 214, 255)
  mThemeColorArray(thmDisabledColor) = RGB(166, 166, 166)
  
  Const SM_CXVSCROLL = 2
  mScrollerSize = GetSystemMetrics(SM_CXVSCROLL) - 2 '<- on Win7 the active ScrollerButtons are two pixels smaller
End Sub

Private Function cTheme_GetIconFontName() As String
Const SPI_GETICONTITLELOGFONT As Long = 31
Dim S As String: S = String$(30, 0)
  SystemParametersInfoW SPI_GETICONTITLELOGFONT, 0, ByVal StrPtr(S), 0
  cTheme_GetIconFontName = Mid$(S, 15, InStr(15, S, vbNullChar) - 15)
End Function
 
Private Sub cTheme_FillThemeColorArray(DstThemeColorArray() As Long)
Dim i As Long
  For i = 0 To UBound(mThemeColorArray)
    DstThemeColorArray(i) = mThemeColorArray(i)
  Next i
End Sub

Private Function cTheme_GetThemeColor(WhichColor As enmThemeColor) As Long
  cTheme_GetThemeColor = mThemeColorArray(WhichColor)
End Function

Private Function cTheme_GetScrollerSize() As Long
  cTheme_GetScrollerSize = mScrollerSize
End Function
 
Private Sub cTheme_DrawTo(CC As cCairoContext, W As cWidgetBase, _
                          ByVal What As enmThemeDrawingType, ByVal State As enmThemeDrawingState, _
                          ByVal x As Double, ByVal y As Double, ByVal dx As Double, ByVal dy As Double, _
                          Optional ByVal Radius As Double, Optional ByVal Direction As enmThemeDrawingDirection)

Dim i As Long, Offs As Single, Size As Single, Tmp As Single
Dim Lines As Long, Angle As Long, ZoomFac As Double
Dim Pat As cCairoPattern, Color As Long, Alpha As Single
  
  Alpha = W.AlphaInherited
  CC.Matrix.CalculateScale ZoomFac, ZoomFac
 
  CC.Save
    Select Case What
      Case thmTypeButtonFace
        CC.SetLineWidth 2
        CC.RoundedRect x, y, dx, dy, Radius, True
        Color = W.BackColor
        If Color = -1 Then Color = cTheme_GetThemeColor(thmBackColor)
        If CBool(State And thmStateHovered) And W.Enabled Then Color = W.HoverColor: Offs = 0.05

        Select Case Direction
          Case thmDirectionDown:  Set Pat = Cairo.CreateLinearPattern(x, y, x, dy)
          Case thmDirectionRight: Set Pat = Cairo.CreateLinearPattern(x, y, dx, y)
          Case thmDirectionLeft:  Set Pat = Cairo.CreateLinearPattern(dx, y, x, y)
        End Select
        If State And thmStatePressed Then
          Pat.AddColorStop 0, Color, Alpha * 0.7, 0.82 + Offs
          Pat.AddColorStop 1, Color, Alpha * 0.85, 1 + Offs
        Else
          Pat.AddColorStop 0, Color, Alpha, 1 + Offs
          Pat.AddColorStop 0.49, Color, Alpha, 0.96 + Offs
          Pat.AddColorStop 0.49, Color, Alpha, 0.92 + Offs
          Pat.AddColorStop 1, Color, Alpha, 0.9 + Offs
        End If
        CC.Fill True, Pat

        CC.SetSourceColor vbWhite, IIf(State And thmStatePressed, 0.45, 0.35) * Alpha
        CC.Stroke (State And thmStatePressed)
        If State And thmStatePressed Then CC.Fill , Pat
 
        
      Case thmTypeShine
          CC.RoundedRect x, y, dx, dy, Radius, True
          Select Case Direction
            Case thmDirectionDown:  Set Pat = Cairo.CreateLinearPattern(x, y, x, dy)
            Case thmDirectionRight: Set Pat = Cairo.CreateLinearPattern(x, y, dx, y)
            Case thmDirectionLeft:  Set Pat = Cairo.CreateLinearPattern(dx, y, x, y)
          End Select
            If State And thmStatePressed Then
              Pat.AddColorStop 0, W.SelectionColor, 0.16 * Alpha
              Pat.AddColorStop 0.21, vbWhite, 0.24 * Alpha
              Pat.AddColorStop 1, vbWhite, 0.62 * Alpha
            Else
              Pat.AddColorStop 0, vbWhite, 0.45 * Alpha
              Pat.AddColorStop 1, vbWhite, 0.35 * Alpha
            End If
          CC.Fill , Pat

      Case thmTypeArrow
        CC.TranslateDrawings (x + dx * 0.5) - 0.04, (y + dy * 0.5) + 0.22
        Angle = GetAngle(Direction)
        
        CC.RotateDrawingsDeg Angle
          CC.SetLineWidth 0.4
'            CC.MoveTo -Int(dx * 0.28) + 0.15, 0
'            CC.LineTo Int(dx * 0.28) + 0.45, -Int(dy * 0.5) - 0.45
'            CC.LineTo Int(dx * 0.28) + 0.45, Int(dy * 0.5) + 0.45
            CC.MoveTo -(dx * 0.29), 0
            CC.LineTo (dx * 0.29), -(dy * 0.58)
            CC.LineTo (dx * 0.29), (dy * 0.58)
          CC.ClosePath
        CC.RotateDrawingsDeg -Angle
        
        Color = W.ForeColor
        If State And thmStateHovered And W.Enabled Then Color = W.FocusColor
        Set Pat = Cairo.CreateLinearPattern(-dx * 0.1, -dy * 0.15, dx * 0.3, dy * 0.3)
          Pat.AddColorStop 0, Color, Alpha, 0.8
          Pat.AddColorStop 1, vbWhite, Alpha * 0.7
        CC.Fill True, Pat
        CC.SetSourceColor &HF0F0F0, Alpha * 0.8
        CC.Stroke
        
      Case thmTypeScrollBarGripperMarks
        If dx <= 10 Or dy <= 10 Then CC.Restore: Exit Sub

        If Direction = thmDirectionVert Then
          CC.TranslateDrawings (x + dx - 0.1), 0
          CC.RotateDrawingsDeg 90
          Tmp = x: x = y: y = Tmp
          Tmp = dx: dx = dy: dy = Tmp
        End If
        
        Lines = dx \ 15
        If Lines > 3 Then Lines = 3
        x = Int(x + (dx - 3 * Lines) \ 2 + 2)
        
        CC.SetLineWidth 3
        CC.SetLineCap Cairo_LINE_CAP_ROUND
        CC.SetSourceColor IIf(State And thmStateHovered, W.HoverColor, vbWhite), Alpha * 0.8, 0.95
        Offs = Int(dy * 0.51) + 0.1
        Size = Int(dy * 0.16)
        For i = 0 To Lines - 1
          CC.DrawLine x + 3 * i - 0.5, Offs - Size - 0.2, x + 3 * i - 0.5, Offs + Size + 1.2
          CC.Stroke
        Next i
        
        CC.SetLineWidth 1
        CC.SetLineCap Cairo_LINE_CAP_SQUARE
        CC.SetSourceColor W.ForeColor, Alpha * 0.6, 0.96
        For i = 0 To Lines - 1
          If Direction = thmDirectionVert Then
            CC.DrawLine x + 3 * i - 0.5, Offs - Size - 0.5, x + 3 * i - 0.5, Offs + Size + 1
          Else
            CC.DrawLine x + 3 * i - 0.5, Offs - Size - 0.5, x + 3 * i - 0.5, Offs + Size + 1
          End If
          CC.Stroke
        Next i
        CC.SetSourceColor W.ForeColor, Alpha * 0.6, 1.3
        For i = 0 To Lines - 1
          If Direction = thmDirectionVert Then
            CC.DrawLine x + 3 * i - 0.5, Offs - Size - 0.5, x + 3 * i + 0.35, Offs + Size + 1
          Else
            CC.DrawLine x + 3 * i + 0.5, Offs - Size - 0.5, x + 3 * i - 0.35, Offs + Size + 1
          End If
          CC.Stroke
        Next i
      
      Case thmTypeDottedRectangle
        CC.SetLineWidth 1 / ZoomFac
        x = CLng(x * ZoomFac) / ZoomFac '- 0.05 '+ 0.05
        y = CLng(y * ZoomFac) / ZoomFac '+ 0.05
        Offs = 0.50005 / ZoomFac
        Size = 1.50005 / ZoomFac
        Radius = (Int(Radius * ZoomFac)) / ZoomFac '- 0.3
        dx = CLng((dx + 0.1) * ZoomFac) / ZoomFac - 0.1
        dy = CLng((dy + 0.1) * ZoomFac) / ZoomFac
        CC.RoundedRect x, y, dx, dy, Radius, True
          CC.SetDashes 0.85 / ZoomFac, Offs, Size
          CC.SetSourceColor vbBlack, Alpha * 0.7
        CC.Stroke
        
        CC.RoundedRect x, y, dx, dy, Radius, True
          CC.SetSourceColor vbWhite, Alpha * 0.6
          CC.SetDashes 1.75 / ZoomFac, Offs, Size
        CC.Stroke
        
      Case thmTypeOptionBox
        CC.SetLineWidth 1
        CC.SetLineCap Cairo_LINE_CAP_BUTT
        Color = IIf(W.MouseOver, W.ShadeColor(W.FocusColor, 0.4), W.BorderColor)
        
        CC.ARC x + Radius, y + Radius, Radius
          If W.MouseOver Then
            CC.SetSourceColor W.HoverColor, Alpha '* 0.9
            CC.Fill True
          End If
          CC.SetSourceColor Color, Alpha * 0.75
        CC.Stroke
        
        CC.ARC x + Radius, y + Radius, Radius - 2.27
          CC.SetSourceColor Color, Alpha * 0.4
        CC.Stroke
        
        Set Pat = Cairo.CreateLinearPattern(x, y, x + Radius, y + dy)
          Color = IIf(W.MouseOver, W.HoverColor, &HE8E8E8)
          Pat.AddColorStop 0, Color, Alpha * 0.8, 0.5
          Pat.AddColorStop 0.75, Color, Alpha * 0.8, 1.1
      
          CC.ARC x + Radius + 0.25, y + Radius + 0.25, Radius - 2.47
        CC.Fill , Pat

        If State = thmStateChecked Then
          Set Pat = Cairo.CreateLinearPattern(x + 3.5, y + 5, x + 6, y + dy - 4)
            Pat.AddColorStop 0, vbWhite, Alpha * 0.9
            Pat.AddColorStop 0.5, W.FocusColor, Alpha
            Pat.AddColorStop 1, W.FocusColor, Alpha, 0.6
        
            CC.ARC x + Radius + 0.027, y + Radius + 0.027, 3.1
          CC.Fill True, Pat
          CC.SetSourceColor vbBlack, 0.8
          CC.Stroke
        End If
        
      Case thmTypeCheckBox
        CC.SetLineWidth 1
        CC.SetLineCap Cairo_LINE_CAP_BUTT
        Color = IIf(W.MouseOver, W.ShadeColor(W.FocusColor, 0.4), W.BorderColor)
        
        CC.Rectangle x, y, dx, dy, True
          If W.MouseOver Then
            CC.SetSourceColor W.HoverColor, Alpha '* 0.9
            CC.Fill True
          End If
          CC.SetSourceColor Color, Alpha * 0.7
        CC.Stroke
        
        CC.Rectangle x + 2, y + 2, dx - 4, dy - 4, True
          CC.SetSourceColor Color, Alpha * 0.35
        CC.Stroke
        
        Set Pat = Cairo.CreateLinearPattern(x, y, x + Radius, y + dy)
          Color = IIf(W.MouseOver, W.HoverColor, &HE8E8E8)
          Pat.AddColorStop 0, Color, Alpha * 0.8, 0.5
          Pat.AddColorStop 0.75, Color, Alpha * 0.8, 1.1
      
          CC.Rectangle x + 2.9, y + 2.9, dx - 4.9, dy - 4.9, True
        CC.Fill , Pat

        If State = thmStateChecked Then
          CC.SetLineCap Cairo_LINE_CAP_ROUND
          CC.SetLineWidth 2
            CC.SetSourceColor W.FocusColor, 0.85, 0.27
            x = x + 2: dx = dx - 4
            y = y + 2: dy = dy - 4
            CC.MoveTo x + dx - 1.75, y + 1
            CC.LineTo x + dx * 0.395, y + dy - 1.9
            CC.LineTo x + 1.75, y + dy * 0.55
          CC.Stroke True
          CC.SetLineWidth 1
            CC.SetSourceColor W.FocusColor, 0.35
          CC.Stroke
        ElseIf State = thmStateCheckUndefined Then
          Set Pat = Cairo.CreateLinearPattern(x + 2, y + 3, x + 6, y + dy - 4)
            Pat.AddColorStop 0, vbWhite, Alpha * 0.8
            Pat.AddColorStop 0.5, W.FocusColor, Alpha, 0.9 + IIf(W.MouseOver, 0.3, 0)
            Pat.AddColorStop 1, W.FocusColor, Alpha, 0.65 + IIf(W.MouseOver, 0.3, 0)
        
            CC.Rectangle x + 2, y + 2, dx - 4, dy - 4, True
          CC.Fill True, Pat
          CC.SetSourceColor vbBlack, 0.4
          CC.Stroke
        End If
        
      Case thmTypeBorder
          CC.RoundedRect x, y, dx, dy, Radius, True
          Color = W.BorderColor
          If State = thmStatePressed Then Color = W.SelectionColor
          If W.Focused Then Color = W.FocusColor
          CC.SetSourceColor Color, IIf(W.MouseOver, 0.85, 0.75) * Alpha
        CC.Stroke
        
      Case thmTypeBottomRightEdge
        CC.Rectangle x, y, dx, dy, True
        Set Pat = Cairo.CreateLinearPattern(x, y, x + dx, y + dy)
            Color = W.BackColor
            If Color = -1 Then Color = cTheme_GetThemeColor(thmBackColor)
            Pat.AddColorStop 0, Color, Alpha, 1.2
            Pat.AddColorStop 0.5, Color, Alpha
            Pat.AddColorStop 0.5, Color, Alpha, 0.9
            Pat.AddColorStop 1, Color, Alpha, 0.95
        CC.Fill , Pat
        
'        CC.TranslateDrawings x + dx, y + dy
'        CC.RotateDrawingsDeg -45
'        If dx > 0 And dy > 0 Then
'          Tmp = IIf(dx > dy, dy / dx, dx / dy)
'          CC.ScaleDrawings Tmp, Tmp
'        End If
'        CC.SetLineWidth 0.8
'        For i = 0 To 2
'          Offs = -(4 + 2 * i)
'          Size = 1 + 2 * i
''          CC.SetSourceColor W.BorderColor, Alpha, 1.15
'          CC.SetSourceColor vbRed, Alpha, 1.15
'          CC.DrawLine -Size, Offs, Size, Offs
'          CC.Stroke
'        Next i

      Case thmTypeListSelection
        Color = W.ShadeColor(IIf(W.DisabledColor = -1, cTheme_GetThemeColor(thmDisabledColor), W.DisabledColor), 1.5)
        If W.Focused And W.Enabled Then Color = W.SelectionColor
        
        Set Pat = Cairo.CreateLinearPattern(x, y + 0.9, x, y + dy - 0.9)
            Pat.AddColorStop 0, Color, Alpha * 0.57, 1.25
            Pat.AddColorStop 0.04, Color, Alpha * 0.58, 1.24
            Pat.AddColorStop 0.12, Color, Alpha * 0.43, 1.11
            Pat.AddColorStop 0.55, Color, Alpha * 0.46, 1.02
            Pat.AddColorStop 0.94, Color, Alpha * 0.56, 0.96
            Pat.AddColorStop 1, vbWhite, Alpha * 0.6
        CC.RoundedRect x, y, dx, dy, Radius, True
          CC.Fill True, Pat
          If W.Focused And W.Enabled Then Color = IIf(W.FocusColor = -1, cTheme_GetThemeColor(thmFocusColor), W.FocusColor)
          CC.SetSourceColor Color, Alpha * 0.5, 0.6
        CC.Stroke
                    
      Case thmTypeSeparatorLine
        x = CLng(x * ZoomFac) / ZoomFac '- 0.05 '+ 0.05
        y = CLng(y * ZoomFac) / ZoomFac '+ 0.05
        If dy > dx Then dx = 1 Else dy = 1
        dx = CLng((dx + 0.1) * ZoomFac) / ZoomFac - 0.1
        dy = CLng((dy + 0.1) * ZoomFac) / ZoomFac
 
        CC.SetLineWidth 2.5 / ZoomFac
        CC.SetLineCap Cairo_LINE_CAP_ROUND
        CC.SetSourceColor vbWhite, 0.3 * Alpha
        CC.MoveTo x + 0.5, y + 0.5
        CC.RelLineTo IIf(dy > dx, 0, dx - 3), IIf(dy > dx, dy - 3, 0)
        CC.Stroke
        
        CC.SetLineWidth 1 / ZoomFac
        Set Pat = Cairo.CreateLinearPattern(x, y, x + IIf(dy > dx, 0, dx), y + IIf(dy > dx, dy, 0))
            Pat.AddColorStop 0, W.BorderColor, Alpha * 0.6 ', 1.27
            Pat.AddColorStop 1, W.BorderColor, 0
        CC.Rectangle x, y, dx, dy, True
        CC.Stroke , Pat
 
      Case thmTypeToolButtonFace
      Case thmTypeMenuButtonFace
      Case thmTypeToolBarFace
      Case thmTypeStatusBarFace
      Case thmTypeProgressFace
      Case thmTypeHeaderFace

      Case thmTypeBorderShadow
      Case thmTypeSliderTrack
 
      Case thmTypeToolBarGrippersMarks
      Case thmTypeResizerGrippersMarks

    End Select
  CC.Restore
End Sub

Private Function GetAngle(Direction As enmThemeDrawingDirection) As Long
  Select Case Direction
    Case thmDirectionDown:  GetAngle = 270
    Case thmDirectionUp:    GetAngle = 90
    Case thmDirectionRight: GetAngle = 180
    Case thmDirectionLeft:  GetAngle = 0
    Case thmDirectionNE:    GetAngle = 135
    Case thmDirectionSE:    GetAngle = 225
    Case thmDirectionSW:    GetAngle = 315
    Case thmDirectionNW:    GetAngle = 45
  End Select
End Function
