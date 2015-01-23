Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
  AddImagesFromFolder New_c.FSO.GetDirList(App.Path & "\Res")

  fMain.Show 'we start with a normal (ANSI-)VB-Form
End Sub

Sub AddImagesFromFolder(DL As cDirList, Optional ByVal DesiredSize As Long = 32)
Dim i As Long, Key As String, R&, G&, B&
  For i = 0 To DL.FilesCount - 1
    Key = Left$(DL.FileName(i), InStrRev(DL.FileName(i), ".") - 1)
    Cairo.ImageList.AddImage Key, DL.Path & DL.FileName(i)
  Next i
End Sub

 
'not used in the App, but available to test the two QR-Classes in the Immediate-Window (without the need to start the App)
Sub QREncDecTest()
  'encoding-direction (into a Cairo-ImageSurface)
  Dim QREnc As New cQREncode, BUTF8() As Byte, QRSrf As cCairoSurface
    BUTF8 = New_c.Crypt.VBStringToUTF8("Hello-QRCode")
    Set QRSrf = QREnc.QREncode(BUTF8, 1, QR_ECLEVEL_H)
  '  QRSrf.WriteContentToPngFile App.Path & "\Hello-QRCode.png"
   
  'decoding-direction (from a Cairo-ImageSurface)
  Dim QRDec As New cQRDecode
    QRDec.DecodeFromSurface QRSrf
    Debug.Print "QRResultsCount"; QRDec.QRResultsCount
    Debug.Print "QRErrString", "["; QRDec.QRErrString(0); "]"
    Debug.Print "QRDataLen", QRDec.QRDataLen(0)
    Debug.Print "QRDataType", QRDec.QRDataType(0)
    Debug.Print "QREccLevel", " "; QRDec.QREccLevel(0)
    Debug.Print "QRMask", QRDec.QRMask(0)
    Debug.Print "QRVersion", QRDec.QRVersion(0)
    Debug.Print "QRDataUTF8Dec  "; New_c.Crypt.UTF8ToVBString(QRDec.QRData(0))
End Sub
