VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQRDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'a simple Binding for the nice Quirc-project of Daniel Beer (https://github.com/dlbeer/quirc) -
'the lib comes under a liberal license, which allows commercial usage (see Bin-Folder for details)
'I've compiled it in a StdCall-version, which this Class expects to be located in a Bin-SubFolder of the App-Path
'Aside from the above library, this Class has no other (Code)Module-dependencies (Olaf Schmidt, Jan. 2015)

Private Declare Function LoadLibraryW& Lib "kernel32" (ByVal lpLibFileName&)
Private Declare Function FreeLibrary& Lib "kernel32" (ByVal hLib&)

Private Const QUIRC_MAX_BITMAP& = 3917
Private Const QUIRC_MAX_PAYLOAD& = 8896

Private Type tQUIRC_POINT
  x As Long
  y As Long
End Type

Private Type tQUIRC_CODE
  Corners(0 To 3) As tQUIRC_POINT ' The four corners of the QR-code, from top left, clockwise

' The number of cells across in the QR-code. The cell bitmap is a bitmask giving the actual values of cells.
' If the cell at (x, y) is black, then the following bit is set: CellBitmap(i * 8) AND (1 * 2 ^ (i AND 7))

  Size As Long ' <- where i = (y * size) + x.
  CellBitmap(0 To QUIRC_MAX_BITMAP - 1) As Byte
End Type

Private Type tQUIRC_DATA
  Version As Long
  EccLevel As eQUIRC_ECC_LEVEL
  Mask As Long
  DataType As QUIRC_DATA_TYPE 'the highest-valued data type found in the QR code.
  Data(0 To QUIRC_MAX_PAYLOAD - 1) As Byte
  DataLen As Long
End Type

Private Enum eQUIRC_DECODE_ERROR
  QUIRC_SUCCESS
  QUIRC_ERROR_INVALID_GRID_SIZE
  QUIRC_ERROR_INVALID_VERSION
  QUIRC_ERROR_FORMAT_ECC
  QUIRC_ERROR_DATA_ECC
  QUIRC_ERROR_UNKNOWN_DATA_TYPE
  QUIRC_ERROR_DATA_OVERFLOW
  QUIRC_ERROR_DATA_UNDERFLOW
End Enum

Public Enum eQUIRC_ECC_LEVEL
  QUIRC_ECC_LEVEL_M
  QUIRC_ECC_LEVEL_L
  QUIRC_ECC_LEVEL_H
  QUIRC_ECC_LEVEL_Q
End Enum

Public Enum QUIRC_DATA_TYPE
  QUIRC_DATA_TYPE_NUMERIC = 1
  QUIRC_DATA_TYPE_ALPHA = 2
  QUIRC_DATA_TYPE_BYTE = 4
  QUIRC_DATA_TYPE_KANJI = 8
End Enum

' Retrieve an initialized QR-code recognizer.
Private Declare Function quirc_new Lib "QuircStdCall" () As Long

' Resize the QR-code recognizer. The size of an image must be specified before codes can be analyzed.
' This function returns 0 on success, or -1 if sufficient memory could not be allocated.
Private Declare Function quirc_resize Lib "QuircStdCall" (ByVal hQ As Long, ByVal Width As Long, ByVal Height As Long) As Long

' quirc_begin() must first be called to obtain access to a buffer into which the input image should be placed.
' Optionally, the current width and height may be returned.
Private Declare Function quirc_begin Lib "QuircStdCall" (ByVal hQ As Long, Optional Width As Long, Optional Height As Long) As Long

' After filling the buffer, quirc_end() should be called to process the image for QR-code recognition.
' The locations and content of each code may be obtained using accessor functions described below.
Private Declare Sub quirc_end Lib "QuircStdCall" (ByVal hQ As Long)

' Return the number of QR-codes identified in the last processed image.
Private Declare Function quirc_count Lib "QuircStdCall" (ByVal hQ As Long) As Long

' Extract the QR-code specified by the given index.
Private Declare Sub quirc_extract Lib "QuircStdCall" (ByVal hQ As Long, ByVal Index As Long, Code As tQUIRC_CODE)

' Decode a QR-code, returning the payload data.
Private Declare Function quirc_decode Lib "QuircStdCall" (Code As tQUIRC_CODE, Data As tQUIRC_DATA) As eQUIRC_DECODE_ERROR

' Destroy a QR-code recognizer.
Private Declare Sub quirc_destroy Lib "QuircStdCall" (ByVal hQ As Long)
 
'Class-internal Helper-Vars
Private hLib&, mResultsCount&, mResultCodes() As tQUIRC_CODE, mResults() As tQUIRC_DATA, mResultErrors$()

Private Sub Class_Initialize()
  hLib = LoadLibraryW(StrPtr(App.Path & "\Bin\QuircStdCall.dll"))
  If hLib = 0 Then Err.Raise vbObjectError, , "couldn't load (or find): " & App.Path & "\Bin\QuircStdCall.dll"
End Sub
Private Sub Class_Terminate()
  If hLib Then FreeLibrary hLib
End Sub

Public Sub DecodeFromSurface(Src As cCairoSurface)
Dim i As Long, x&, y&, hQ As Long, pImg As Long, SrfPxl() As Byte, QRPxl() As Byte
  mResultsCount = 0
  If Src Is Nothing Then Exit Sub
  
  With Src.CreateSimilar.CreateContext 'creata a temporary Surface from the Source
    .SetSourceColor vbBlack
    .Paint 'pre-fill the Tmp-Surface with Black
 
    .Operator = CAIRO_OPERATOR_HSL_LUMINOSITY 'grayscale-conversion operator
    .Paint , Src.CreateSurfacePattern 'draw from the Source-Surface
 
    'now that we have greyscaled Surface-Content - we start to copy over into an 8BPP-ByteArray
    ReDim QRPxl(0 To Src.Width - 1, 0 To Src.Height - 1)
    If .Surface.BindToArray(SrfPxl) Then
      For y = 1 To UBound(SrfPxl, 2) - 1
        i = 1
        For x = 4 To UBound(SrfPxl, 1) - 4 Step 4
          If SrfPxl(x, y) < 128 Then QRPxl(i, y) = 1 'a QRInput-Arr needs a '1' for black - and a '0' for white
          i = i + 1
        Next x
      Next y
      .Surface.ReleaseArray SrfPxl
    Else
      Err.Raise vbObjectError, , "couldn't create GreyScale-Data from Surface"
    End If
  End With
  
  'our 8BPP QRPxlArr is prepared, so we can start the analysis on it now (incorporating the quirc-lib)
  hQ = quirc_new()
  
    If hQ = 0 Then Err.Raise vbObjectError, , "can't create quirc-handle"
    
    If quirc_resize(hQ, Src.Width, Src.Height) < 0 Then
       quirc_destroy hQ
       Err.Raise vbObjectError, , "can't allocate memory for the quirc-internal image"
    End If
    
    pImg = quirc_begin(hQ, 0, 0) 'prepare for image-analysis
      New_c.MemCopy pImg, VarPtr(QRPxl(0, 0)), Src.Width * Src.Height 'copy 8BPP-GreyData over
    quirc_end hQ 'perform image-analysis-process (and end it internally)
    
    
    mResultsCount = quirc_count(hQ) 'get the amount of QR-Areas we've found on that image
     
    If mResultsCount Then 'store Results in Class-internal Result-Arrays
      ReDim mResultCodes(0 To mResultsCount - 1)
      ReDim mResults(0 To mResultsCount - 1)
      ReDim mResultErrors(0 To mResultsCount - 1)
      
      For i = 0 To UBound(mResults)
        quirc_extract hQ, i, mResultCodes(i)
        mResultErrors(i) = GetErrorString(quirc_decode(mResultCodes(i), mResults(i)))
      Next i
    End If
  
  quirc_destroy hQ
End Sub

'index-like access into the Results (since there's possibly more than one single QR-Code per Image)
Public Property Get QRResultsCount() As Long
  QRResultsCount = mResultsCount
End Property
Public Property Get QRVersion(ByVal IndexZeroBased As Long) As Long
  QRVersion = mResults(IndexZeroBased).Version
End Property
Public Property Get QREccLevel(ByVal IndexZeroBased As Long) As String
  QREccLevel = Choose(mResults(IndexZeroBased).EccLevel + 1, "M", "L", "H", "Q")
End Property
Public Property Get QRMask(ByVal IndexZeroBased As Long) As Long
  QRMask = mResults(IndexZeroBased).Mask
End Property
Public Property Get QRDataType(ByVal IndexZeroBased As Long) As QUIRC_DATA_TYPE
  QRDataType = mResults(IndexZeroBased).DataType
End Property
Public Property Get QRData(ByVal IndexZeroBased As Long) As Byte()
  If QRDataLen(IndexZeroBased) <= 0 Or QRDataLen(IndexZeroBased) > QUIRC_MAX_PAYLOAD Then QRData = vbNullString: Exit Property
  With New_c.Stream
    .WriteFromPtr VarPtr(mResults(IndexZeroBased).Data(0)), QRDataLen(IndexZeroBased)
    .SetPosition 0
    .ReadToByteArr QRData
  End With
End Property
Public Property Get QRDataLen(ByVal IndexZeroBased As Long) As Long
  If mResultsCount Then QRDataLen = mResults(IndexZeroBased).DataLen
End Property
Public Property Get QRErrString(ByVal IndexZeroBased As Long) As String
  QRErrString = mResultErrors(IndexZeroBased)
End Property


'just a small helper, to translate the Err_Enum into Err_Strings
Private Function GetErrorString(ByVal E As eQUIRC_DECODE_ERROR) As String
  Select Case E
    Case QUIRC_ERROR_INVALID_GRID_SIZE: GetErrorString = "Invalid grid size"
    Case QUIRC_ERROR_INVALID_VERSION:   GetErrorString = "Invalid version"
    Case QUIRC_ERROR_FORMAT_ECC:        GetErrorString = "Format data ECC failure"
    Case QUIRC_ERROR_DATA_ECC:          GetErrorString = "ECC failure"
    Case QUIRC_ERROR_UNKNOWN_DATA_TYPE: GetErrorString = "Unknown data type"
    Case QUIRC_ERROR_DATA_OVERFLOW:     GetErrorString = "Data overflow"
    Case QUIRC_ERROR_DATA_UNDERFLOW:    GetErrorString = "Data underflow"
  End Select
End Function

