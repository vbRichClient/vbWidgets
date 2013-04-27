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

Private CairoSurface As cCairoSurface, SVG As cSVG, SVGFileName As String

'****---- Start of cwImplementation-Conventions ----****
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
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
'****---- End of cwImplementation-Conventions ----****


Public Sub SetImage(ImageSrfOrFileName As Variant, Optional ByVal BufferSVGs As Boolean, _
                                      Optional ByVal SVGBufferWidth As Long, _
                                      Optional ByVal SVGBufferHeight As Long)
    If VarType(ImageSrfOrFileName) = vbString Then
        If InStr(ImageSrfOrFileName, ".svg") Then
            Set CairoSurface = Nothing
                   
            If SVGFileName <> ImageSrfOrFileName Then 'this ensures, that the Parsing-Part of reading an SVG is done only on FileName-Changes
              SVGFileName = ImageSrfOrFileName
              SVG.ParseContent SVGFileName
              If BufferSVGs Then
                If SVGBufferWidth > 0 And SVGBufferHeight > 0 Then
                  Set CairoSurface = Cairo.CreateSurface(SVGBufferWidth + 2, SVGBufferHeight + 2)
                Else
                  Set CairoSurface = Cairo.CreateSurface(W.Width + 2, W.Height + 2)
                End If
                SVG.RenderFromDOM CairoSurface.CreateContext, 0, 0, CairoSurface.Width, CairoSurface.Height
              End If
            End If
        Else
            Set CairoSurface = Cairo.CreateSurface(W.Width, W.Height, ImageSurface, ImageSrfOrFileName)
        End If
    ElseIf IsObject(ImageSrfOrFileName) Then
        Set CairoSurface = ImageSrfOrFileName
    End If
End Sub

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
'
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  If Len(W.ImageKey) Then
    CC.RenderSurfaceContent W.ImageKey, 0, 0, dx_Aligned, dy_Aligned
    Exit Sub
  End If
  
  If Len(SVGFileName) Then
    If CairoSurface Is Nothing Then 'unbuffered SVG-mode (slower, but accurate in any zoom-level and size)
      SVG.RenderFromDOM CC, 0, 0, dx_Aligned, dy_Aligned
    Else 'buffered mode, with one pixel overscan on all edges
      CC.RenderSurfaceContent CairoSurface, -1, -1, dx_Aligned + 2, dy_Aligned + 2
    End If
  Else
    If Not CairoSurface Is Nothing Then
      CC.RenderSurfaceContent CairoSurface, 0, 0, dx_Aligned, dy_Aligned
    End If
  End If
End Sub

