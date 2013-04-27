Attribute VB_Name = "modWidgets"
Option Explicit

Public Const VSplitCursor_Png$ = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAABnRSTlMAwADAAMCNeLu6AAAAb0lEQVR42u3WSwqAMAwEUEe8V66enGxcdCNSmkCtLpysSgl59BcKd99Wxr60uoB/AEcxz8wAkIyI6/j7FQgQ0KLdy9skADN7AOhWrxsJMKheNBIgIkgOEtL3nG/RwKh0i9Ihdw31IgGvAdDPTsB0nEm6NMFxeZ+IAAAAAElFTkSuQmCC"
Public Const HSplitCursor_Png$ = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAABnRSTlMAwADAAMCNeLu6AAAAh0lEQVR42u2UsQ7AIAhEoel/8evyZedg4mKbcBibDtzkQHjAIdpak5O6jmYvQAEKkAOY2UGAmakqxSAAI7uIUIw7PoeRfb4BrDHuHupAn5SIeQUgq+iI1k7TIjyIaMuDiLY8SMRwHgCYNQII+kR8NHcfNcazc4DJoHaMPnbsBv/vXBegAN8DOhNVk1H7kjSuAAAAAElFTkSuQmCC"

Declare Function GetInstanceEx Lib "DirectCom" (StrPtr_FName As Long, StrPtr_ClassName As Long, ByVal UseAlteredSearchPath As Boolean) As Object

Public New_c As cConstructor, Cairo As cCairo

Public Sub Main()
  On Error Resume Next
    Set New_c = GetInstanceEx(StrPtr(App.Path & "\vbRichClient5.dll"), StrPtr("cConstructor"), True)
  If New_c Is Nothing Then
    Err.Clear
    Set New_c = New cConstructor
  End If
  
  Set Cairo = New_c.Cairo
  
  Set Cairo.Theme = New cThemeWin7
'  Cairo.FontOptions = CAIRO_ANTIALIAS_DEFAULT
End Sub

