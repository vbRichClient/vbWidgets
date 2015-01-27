VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQREncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'a simple Binding for the nice libqrencode-project of Kentaro Fukuchi (http://fukuchi.org/works/qrencode/) -
'the lib comes under LGPL-license which allows also commercial usage ...
'I've compiled it into the latest release of vb_cairo_sqlite.dll in a VB-friendly StdCall-version
'Aside from vb_cairo_sqlite, this Class has no other (Code)Module-dependencies (Olaf Schmidt, Jan. 2015)

Public Enum QRencodeMode
  QR_MODE_NUM = 0    'Numeric mode
  QR_MODE_AN         'Alphabet-numeric mode
  QR_MODE_8          '8-bit data mode
  QR_MODE_KANJI      'Kanji (shift-jis) mode
  QR_MODE_STRUCTURE  'Internal use only
  QR_MODE_ECI        'ECI mode
  QR_MODE_FNC1FIRST  'FNC1, first position
  QR_MODE_FNC1SECOND 'FNC1, second position
End Enum
 
Public Enum QRecLevel 'Error-Correction-Level
  QR_ECLEVEL_L = 0 'lowest
  QR_ECLEVEL_M
  QR_ECLEVEL_Q
  QR_ECLEVEL_H     'highest
End Enum

Private Type QRcode
  Version As Long
  Width As Long 'Width describes the squared Width (and Height) of the returned symbol
  pResult As Long 'pointer to a data-array with a ByteLength of Width ^ 2
' each Byte in the above result contains the following Info in its 8 Bits
' * MSB 76543210 LSB
' *     |||||||`- 1=black/0=white
' *     ||||||`-- data and ecc code area
' *     |||||`--- format information
' *     ||||`---- version information
' *     |||`----- timing pattern
' *     ||`------ alignment pattern
' *     |`------- finder pattern and separator
' *     `-------- non-data modules (format, timing, etc.)
End Type
 
Private Declare Function QRcode_encodeData Lib "vb_cairo_sqlite" (ByVal LenData As Long, ByVal pData As Long, ByVal Version As Long, ByVal EcLevel As QRecLevel) As Long
Private Declare Function QRcode_encodeDataMQR Lib "vb_cairo_sqlite" (ByVal LenData As Long, ByVal pData As Long, ByVal Version As Long, ByVal EcLevel As QRecLevel) As Long
Private Declare Sub QRcode_free Lib "vb_cairo_sqlite" (ByVal hQR As Long)
 
Public Function QREncode(Data() As Byte, Optional Version_1to40 As Long, _
                                         Optional ByVal EcLevel As QRecLevel = QR_ECLEVEL_M, _
                                         Optional ByVal PixelSize As Long = 4, _
                                         Optional ByVal SrfType As SurfaceCreateEnum = ImageSurface) As cCairoSurface
                                         
Dim hQR As Long, LenData As Long, QR As QRcode, Result() As Byte
  On Error Resume Next
    LenData = UBound(Data) - LBound(Data) + 1
  On Error GoTo 0
  If LenData = 0 Then Set QREncode = Cairo.CreateSurface(1, 1, SrfType): Exit Function
  
  If Version_1to40 < 0 Then Version_1to40 = 0 Else If Version_1to40 > 40 Then Version_1to40 = 40
  hQR = QRcode_encodeData(LenData, VarPtr(Data(LBound(Data))), Version_1to40, EcLevel)
  If hQR Then New_c.MemCopy VarPtr(QR), hQR, LenB(QR) Else Exit Function
  
  ReDim Result(0 To QR.Width - 1, 0 To QR.Width - 1)
  New_c.MemCopy VarPtr(Result(0, 0)), QR.pResult, QR.Width ^ 2
  QRcode_free hQR
  
  Version_1to40 = QR.Version 'report the resulting Version back to the caller in our ByRef-Param
  Set QREncode = RenderResultToSurface(Result, PixelSize, SrfType)
End Function
 
Public Function QREncodeMQR(Data() As Byte, Optional ByVal Version_1to4 As Long = 4, _
                                            Optional ByVal EcLevel As QRecLevel = QR_ECLEVEL_M, _
                                            Optional ByVal PixelSize As Long = 4, _
                                            Optional ByVal SrfType As SurfaceCreateEnum = ImageSurface) As cCairoSurface
                                            
Dim hQR As Long, LenData As Long, QR As QRcode, Result() As Byte
  On Error Resume Next
    LenData = UBound(Data) - LBound(Data) + 1
  On Error GoTo 0
  If LenData = 0 Then Set QREncodeMQR = Cairo.CreateSurface(1, 1, SrfType): Exit Function
  'other than in normal QREncoding above, for MicroQR the Version is not "auto-expanded" -
  'one has to provide the Version in the right size (for the amount of Data to be taken up) oneself
  If Version_1to4 < 1 Then Version_1to4 = 1 Else If Version_1to4 > 4 Then Version_1to4 = 4
  hQR = QRcode_encodeDataMQR(LenData, VarPtr(Data(LBound(Data))), Version_1to4, EcLevel)
  If hQR Then New_c.MemCopy VarPtr(QR), hQR, LenB(QR) Else Exit Function
  
  ReDim Result(0 To QR.Width - 1, 0 To QR.Width - 1)
  New_c.MemCopy VarPtr(Result(0, 0)), QR.pResult, QR.Width ^ 2
  QRcode_free hQR
  
  Set QREncodeMQR = RenderResultToSurface(Result, PixelSize, SrfType)
End Function

Private Function RenderResultToSurface(Result() As Byte, ByVal PixelSize&, _
                 Optional ByVal SrfType As SurfaceCreateEnum = ImageSurface) As cCairoSurface
Dim x As Long, y As Long, Width As Long
  Width = UBound(Result) + 1
  If PixelSize < 1 Then PixelSize = 1 Else If PixelSize > 15 Then PixelSize = 15
  Set RenderResultToSurface = Cairo.CreateSurface(PixelSize * Width, PixelSize * Width, SrfType)
  With RenderResultToSurface.CreateContext
    .SetSourceColor vbWhite: .Paint 'ensure white BackGround
    
    .ScaleDrawings PixelSize, PixelSize 'scale the output appropriately
    
    .SetSourceColor vbBlack 'since white background was ensured above...
        For y = 0 To Width - 1: For x = 0 To Width - 1
          If Result(x, y) And 1 Then .Rectangle x, y, 1, 1 '...only black rects are drawn now
        Next x, y
    .Fill 'fill the black rectangles
  End With
End Function
 


