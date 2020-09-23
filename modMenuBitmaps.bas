Attribute VB_Name = "modMenuBitMap"
Option Explicit
Public Const TSelect As Long = 1
Public Const TPencil As Long = 2
Public Const TLine As Long = 3
Public Const TRect As Long = 4
Public Const TFRect As Long = 5
Public Const TCirc As Long = 6
Public Const TFCirc As Long = 7
Public Const TSelColor As Long = 8
Public Const TFlood As Long = 9
Public Const TText As Long = 10
Public Const TErase As Long = 11
Public Const MF_BYPOSITION = &H400 ' assigns the new item by zero-based position
Public Const HALFTONE = 4
Private Type PALETTEENTRY
 peRed   As Byte
 peGreen As Byte
 peBlue  As Byte
 peFlags As Byte
End Type
Public Type LOGPALETTE
 palVersion       As Integer
 palNumEntries    As Integer
 palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal FillType As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function GetNearestPaletteIndex Lib "gdi32.dll" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'publics for frmText
Public FontDesc As String
Public gFontName As String
Public gFontBold As Boolean
Public gFontItalic As Boolean
Public gFontSize As Double
Public gCancelled As Boolean
Public gText As String
Public Pal(255) As Long
Public Function RedV(ByVal clr As Long) As Byte
 RedV = clr And &HFF&
End Function
Public Function GrnV(ByVal clr As Long) As Byte
 GrnV = (clr And &HFF00&) \ 256
End Function
Public Function BluV(ByVal clr As Long) As Byte
 BluV = (clr And &HFF0000) \ 65536
End Function
Public Function FileExists(ByVal sFile As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sFile)
 FileExists = (Err.Number = 0) And ((eAttr And vbDirectory) = 0)
 On Error GoTo 0
End Function

Public Sub SetMenuItemBMP(ByVal FrmHWnd As Long, _
 ByVal MainIndex As Long, ByVal SubIndex As Long, _
 Pic As IPictureDisp)
'A little explanation here
' MainIndex - the zero based top level Menu Item
'   for instance, in this app that would
'   be File - 0
'      Edit - 1
'      View Menu Bitmap - 2
' SubIndex - the zero based index within MainIndex
'   so within the File MainIndex
'      New - 0
'      Open - 1
'      Save - 2 , etc.
' for more complex menu structures
' you would just continue drilling down
' hMenu = GetMenu(FrmHWnd)
' hSubMenu = GetSubMenu(hMenu, MainIndex)
' hSubMenu2 = GetSubMenu(hSubMenu, Menu2FirstIndex)

 Dim hMenu As Long
 Dim hSubMenu As Long
 hMenu = GetMenu(FrmHWnd)
 hSubMenu = GetSubMenu(hMenu, MainIndex) '1st top level
 SetMenuItemBitmaps hSubMenu, SubIndex, MF_BYPOSITION, Pic, 0&
End Sub
Public Sub GetPal()
 'load up our default palette
 Dim fHdl As Long
 fHdl = FreeFile
 Open App.Path & "\defpal.bin" For Binary As fHdl
 Get #fHdl, , Pal
 Close fHdl
End Sub

