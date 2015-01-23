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
Attribute VB_Exposed = False
Option Explicit

'a simple Binding for the nice libqrencode-project of Kentaro Fukuchi (http://fukuchi.org/works/qrencode/) -
'the lib comes under LGPL-license which allows commercial usage of the compiled Binary...
'I've compiled it in a StdCall-version, which this Class expects to be located in a Bin-SubFolder of the App-Path
'Aside from the above library, this Class has no other (Code)Module-dependencies (Olaf Schmidt, Jan. 2015)

Private Declare Function LoadLibraryW& Lib "kernel32" (ByVal lpLibFileName&)
Private Declare Function FreeLibrary& Lib "kernel32" (ByVal hLib&)

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
' each Byte in the above result contains the following Info in its 8 Bits:
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
 
Private Declare Function QRcode_encodeData Lib "libQREncodeStdCall" (ByVal LenData As Long, ByVal pData As Long, ByVal Version As Long, ByVal EcLevel As QRecLevel) As Long
Private Declare Function QRcode_encodeDataMQR Lib "libQREncodeStdCall" (ByVal LenData As Long, ByVal pData As Long, ByVal Version As Long, ByVal EcLevel As QRecLevel) As Long
Private Declare Sub QRcode_free Lib "libQREncodeStdCall" (ByVal hQR As Long)

Private hLib As Long

Private Sub Class_Initialize()
  hLib = LoadLibraryW(StrPtr(App.Path & "\Bin\libQREncodeStdCall.dll"))
  If hLib = 0 Then Err.Raise vbObjectError, , "couldn't load (or find): " & App.Path & "\Bin\libQREncodeStdCall.dll"
End Sub
Private Sub Class_Terminate()
  If hLib Then FreeLibrary hLib
End Sub

Public Function QREncode(Data() As Byte, Optional Version_1to40 As Long, _
                                         Optional ByVal EcLevel As QRecLevel = QR_ECLEVEL_M, _
                                         Optional ByVal PixelSize As Long = 4) As cCairoSurface
                                         
Dim hQR As Long, LenData As Long, QR As QRcode, Result() As Byte
  On Error Resume Next
    LenData = UBound(Data) - LBound(Data) + 1
  On Error GoTo 0
  If LenData = 0 Then Set QREncode = Cairo.CreateSurface(1, 1): Exit Function
  
  If Version_1to40 < 0 Then Version_1to40 = 0 Else If Version_1to40 > 40 Then Version_1to40 = 40
  hQR = QRcode_encodeData(LenData, VarPtr(Data(LBound(Data))), Version_1to40, EcLevel)
  If hQR Then New_c.MemCopy VarPtr(QR), hQR, LenB(QR) Else Exit Function
  
  ReDim Result(0 To QR.Width - 1, 0 To QR.Width - 1)
  New_c.MemCopy VarPtr(Result(0, 0)), QR.pResult, QR.Width ^ 2
  QRcode_free hQR
  
  Version_1to40 = QR.Version 'report the resulting Version back to the caller in our ByRef-Param
  Set QREncode = RenderResultToSurface(Result, PixelSize)
End Function

Public Function QREncodeMQR(Data() As Byte, Optional ByVal Version_1to4 As Long = 4, _
                                            Optional ByVal EcLevel As QRecLevel = QR_ECLEVEL_M, _
                                            Optional ByVal PixelSize As Long = 4) As cCairoSurface
                                            
Dim hQR As Long, LenData As Long, QR As QRcode, Result() As Byte
  On Error Resume Next
    LenData = UBound(Data) - LBound(Data) + 1
  On Error GoTo 0
  If LenData = 0 Then Set QREncodeMQR = Cairo.CreateSurface(1, 1): Exit Function
  'other than in normal QREncoding above, for MicroQR the Version is not "auto-expanded" -
  'one has to provide the Version in the right size (for the amount of Data to be taken up) oneself
  If Version_1to4 < 1 Then Version_1to4 = 1 Else If Version_1to4 > 4 Then Version_1to4 = 4
  hQR = QRcode_encodeDataMQR(LenData, VarPtr(Data(LBound(Data))), Version_1to4, EcLevel)
  If hQR Then New_c.MemCopy VarPtr(QR), hQR, LenB(QR) Else Exit Function
  
  ReDim Result(0 To QR.Width - 1, 0 To QR.Width - 1)
  New_c.MemCopy VarPtr(Result(0, 0)), QR.pResult, QR.Width ^ 2
  QRcode_free hQR
  
  Set QREncodeMQR = RenderResultToSurface(Result, PixelSize)
End Function

Private Function RenderResultToSurface(Result() As Byte, ByVal PixelSize&) As cCairoSurface
Dim x As Long, y As Long, Width As Long
  Width = UBound(Result) + 1
  If PixelSize < 1 Then PixelSize = 1 Else If PixelSize > 15 Then PixelSize = 15
  Set RenderResultToSurface = Cairo.CreateSurface(PixelSize * Width, PixelSize * Width)
  With RenderResultToSurface.CreateContext
    .ScaleDrawings PixelSize, PixelSize
    For y = 0 To Width - 1: For x = 0 To Width - 1
      .Rectangle x, y, 1, 1
      .SetSourceColor IIf(Result(x, y) And 1, vbBlack, vbWhite)
      .Fill
    Next x, y
  End With
End Function

