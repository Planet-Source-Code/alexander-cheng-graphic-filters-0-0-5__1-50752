VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl PicEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   3510
   ScaleWidth      =   3510
   Begin VB.CommandButton cmdResume 
      Caption         =   "O"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   3240
      Width           =   255
   End
   Begin VB.HScrollBar VH 
      Height          =   255
      Left            =   0
      Max             =   50
      Min             =   -50
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.VScrollBar BV 
      Height          =   3255
      Left            =   3240
      Max             =   50
      Min             =   -50
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox PicDisp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   3240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicProcess 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image picBackup 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "PicEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private pDATA() As Byte
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Private mProgress As Long         '% filter progress
Private Speed(0 To 765) As Long   'Speed up values

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type BITMAPINFOHEADER   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Public PictureLoaded As Boolean 'Boolean value to indicated whether pic is loaded or not
Private hdcNew As Long
Private oldhand As Long
Private ret As Long
Private BytesPerScanLine As Long
Private PadBytesPerScanLine As Long

Private R As RECT 'Ellipse Region
Private mRGN As Long
Private Ellipse_PatternBrush As Long
Private SelectionMode As Long ' 1 = Rectangular, 2 = Ellipse

Private m_MouseDown As Boolean 'Mouse is down
Private sX As Long
Private sY As Long
Private eX As Long
Private eY As Long

Event Resize(Width As Long, Height As Long)
Event ImageLoaded(bmType As Long, bmWidth As Long, bmHeight As Long, bmBitsPixel As Integer)

Private Sub GetRGB(ByVal Col As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
  R = Col Mod 256
  G = ((Col And &HFF00&) \ 256&) Mod 256&
  B = (Col And &HFF0000) \ 65536
End Sub

Private Sub cmdResume_Click()
    BV.Value = 0: VH.Value = 0
End Sub

Private Sub UserControl_InitProperties()
    sX = 0: sY = 0
    eX = 0: eY = 0
End Sub

Sub OpenImage(Filename As String)
    On Error GoTo ErrorOO
    
    PicDisp.Picture = LoadPicture(Filename)
    PicProcess.Picture = LoadPicture(Filename)
    
With picBackup
    .Width = PicProcess.Width
    .Height = PicProcess.Height
    .Picture = PicProcess.Picture
End With

    GetBMPInfo 'call all APIs to retrive BITMAP info
    PictureLoaded = True
    eX = PicProcess.Width: eY = PicProcess.Height
    RaiseEvent ImageLoaded(PicInfo.bmType, PicInfo.bmWidth, PicInfo.bmHeight, DIBInfo.bmiHeader.biBitCount)
Exit Sub
ErrorOO:
    PictureLoaded = False

End Sub

Sub SaveImage(ByVal Filename As String)
    On Error GoTo SaveError
    SavePicture PicDisp.Image, Filename
    Exit Sub
SaveError:
    MsgBox "An error has occured!"
End Sub

Private Sub GetBMPInfo()
    GetObject PicProcess.Image, Len(PicInfo), PicInfo ' reads bmp info
    hdcNew = CreateCompatibleDC(PicProcess.hdc) ' creates DC
    oldhand = SelectObject(hdcNew, PicProcess.Image) 'make DC capable
    
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  
    ReDim iDATA(3, PicInfo.bmWidth - 1, PicInfo.bmHeight - 1) As Byte
    ReDim bDATA(3, PicInfo.bmWidth - 1, PicInfo.bmHeight - 1) As Byte
    ReDim pDATA(3, PicInfo.bmWidth - 1, PicInfo.bmHeight - 1) As Byte
    GetDIBits hdcNew, PicProcess.Image, 0, PicInfo.bmHeight, iDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
    GetDIBits hdcNew, PicProcess.Image, 0, PicInfo.bmHeight, bDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
    GetDIBits PicDisp.hdc, PicDisp.Image, 0, PicInfo.bmHeight, pDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
End Sub

Sub UndoLast()
    PicDisp.Picture = Nothing
    PicProcess.Picture = Nothing
    SetDIBits PicDisp.hdc, PicDisp.Image, 0, PicInfo.bmHeight, pDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
    SetDIBits PicProcess.hdc, PicProcess.Image, 0, PicInfo.bmHeight, pDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
    GetDIBits PicDisp.hdc, PicDisp.Image, 0, PicInfo.bmHeight, iDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
    GetDIBits PicDisp.hdc, PicDisp.Image, 0, PicInfo.bmHeight, bDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
End Sub

Private Sub BackUp_Previous()
    Erase pDATA
    ReDim pDATA(3, PicInfo.bmWidth - 1, PicInfo.bmHeight - 1) As Byte
    ret = GetDIBits(PicDisp.hdc, PicDisp.Image, 0, PicInfo.bmHeight, pDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS)
End Sub

Sub ReleaseAllStuff()
    SelectObject hdcNew, oldhand
    DeleteDC hdcNew
    DeleteObject mRGN
    DeleteObject Ellipse_PatternBrush
    PictureLoaded = False
    Erase iDATA
    Erase bDATA
    Erase pDATA
End Sub

Sub UndoAll()
    PicProcess.Picture = Nothing
    PicDisp.Picture = Nothing
    PicProcess.Picture = picBackup.Picture
    PicDisp.Picture = picBackup.Picture
    ret = GetDIBits(hdcNew, PicProcess.Image, 0, PicInfo.bmHeight, iDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS)
    ret = GetDIBits(hdcNew, PicProcess.Image, 0, PicInfo.bmHeight, bDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS)
End Sub
Private Sub PicDisp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    R.Top = 0: R.Left = 0: R.Bottom = PicDisp.ScaleHeight: R.Right = PicDisp.ScaleWidth
    mRGN = CreateRectRgn(R.Left, R.Top, R.Right, R.Bottom)
    ret = SelectObject(PicDisp.hdc, mRGN)

    If Button = 1 Then
        sX = X
        sY = Y
        eX = X
        eY = Y
        SelectionMode = 1
    End If
    
    If Button = 2 Then
        R.Top = Y
        R.Left = X
        R.Bottom = Y
        R.Right = X
        SelectionMode = 2
    End If
    
m_MouseDown = True
End Sub

Private Sub PicDisp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_MouseDown And Button = 1 Then
        With PicDisp
        .ForeColor = vbYellow
        .AutoRedraw = True
        .DrawMode = vbInvert
        PicDisp.Line (sX, sY)-(eX, eY), vbYellow, B
        eX = X: eY = Y
        PicDisp.Line (sX, sY)-(eX, eY), vbYellow, B
        .Refresh
        SelectionMode = 1
        End With
    End If
    
    If m_MouseDown And Button = 2 Then
        With PicDisp
        .ForeColor = vbYellow
        .AutoRedraw = True
        .DrawMode = vbXorPen
        End With
        Ellipse PicDisp.hdc, R.Left, R.Top, R.Right, R.Bottom
        R.Bottom = Y: R.Right = X
        Ellipse PicDisp.hdc, R.Left, R.Top, R.Right, R.Bottom
        PicDisp.Refresh
        SelectionMode = 2
    End If
    
End Sub

Private Sub PicDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If m_MouseDown And Button = 1 Then
    With PicDisp
        PicDisp.Line (sX, sY)-(eX, eY), vbYellow, B
        .DrawMode = vbCopyPen
    End With
End If

If m_MouseDown And Button = 2 Then
    R.Bottom = Y: R.Right = X
    Ellipse PicDisp.hdc, R.Left, R.Top, R.Right, R.Bottom
    With PicDisp
    .Refresh
    .DrawMode = vbCopyPen
    End With
End If

m_MouseDown = False

End Sub

Private Sub PicDisp_Resize()
'move 2 bars, bar BV and var VH
    With BV
        .Height = PicDisp.Height
        .Move PicDisp.Width, 0
    End With
    With VH
        .Width = PicDisp.Width
        .Move 0, PicDisp.Height
    End With
'=======================
    cmdResume.Move PicDisp.Width, PicDisp.Height
    
    UserControl.Width = PicDisp.Width + BV.Width
    UserControl.Height = PicDisp.Height + VH.Height
    RaiseEvent Resize(UserControl.Width, UserControl.Height)
End Sub

Sub Aqua()
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim Med(1 To 4) As Long
  Dim Dev(1 To 4) As Long
  Dim I As Long, j As Long
  Dim sDev As Long, vDev As Long
  
  For Y = 2 To UBound(iDATA, 3) - 2
    For X = 2 To UBound(iDATA, 2) - 2
      For I = 0 To 2
        Med(1) = CLng(bDATA(I, X - 2, Y - 2)) + CLng(bDATA(I, X - 1, Y - 2)) + CLng(bDATA(I, X, Y - 2)) + _
                 CLng(bDATA(I, X - 2, Y - 1)) + CLng(bDATA(I, X - 1, Y - 1)) + CLng(bDATA(I, X, Y - 1)) + _
                 CLng(bDATA(I, X - 2, Y)) + CLng(bDATA(I, X - 1, Y)) + CLng(bDATA(I, X, Y))
        Med(2) = CLng(bDATA(I, X + 2, Y - 2)) + CLng(bDATA(I, X + 1, Y - 2)) + CLng(bDATA(I, X, Y - 2)) + _
                 CLng(bDATA(I, X + 2, Y - 1)) + CLng(bDATA(I, X + 1, Y - 1)) + CLng(bDATA(I, X, Y - 1)) + _
                 CLng(bDATA(I, X + 2, Y)) + CLng(bDATA(I, X + 1, Y)) + CLng(bDATA(I, X, Y))
        Med(3) = CLng(bDATA(I, X - 2, Y + 2)) + CLng(bDATA(I, X - 1, Y + 2)) + CLng(bDATA(I, X, Y + 2)) + _
                 CLng(bDATA(I, X - 2, Y + 1)) + CLng(bDATA(I, X - 1, Y + 1)) + CLng(bDATA(I, X, Y + 1)) + _
                 CLng(bDATA(I, X - 2, Y)) + CLng(bDATA(I, X - 1, Y)) + CLng(bDATA(I, X, Y))
        Med(4) = CLng(bDATA(I, X + 2, Y + 2)) + CLng(bDATA(I, X + 1, Y + 2)) + CLng(bDATA(I, X, Y + 2)) + _
                 CLng(bDATA(I, X + 2, Y + 1)) + CLng(bDATA(I, X + 1, Y + 1)) + CLng(bDATA(I, X, Y + 1)) + _
                 CLng(bDATA(I, X + 2, Y)) + CLng(bDATA(I, X + 1, Y)) + CLng(bDATA(I, X, Y))
        Med(1) = Med(1) \ 9
        Med(2) = Med(2) \ 9
        Med(3) = Med(3) \ 9
        Med(4) = Med(4) \ 9
        Dev(1) = Abs(CLng(bDATA(I, X - 2, Y - 2)) - Med(1)) + Abs(CLng(bDATA(I, X - 1, Y - 2)) - Med(1)) + Abs(CLng(bDATA(I, X, Y - 2)) - Med(1)) + _
                 Abs(CLng(bDATA(I, X - 2, Y - 1)) - Med(1)) + Abs(CLng(bDATA(I, X - 1, Y - 1)) - Med(1)) + Abs(CLng(bDATA(I, X, Y - 1)) - Med(1)) + _
                 Abs(CLng(bDATA(I, X - 2, Y)) - Med(1)) + Abs(CLng(bDATA(I, X - 1, Y)) - Med(1)) + Abs(CLng(bDATA(I, X, Y)) - Med(1))
        Dev(2) = Abs(CLng(bDATA(I, X + 2, Y - 2)) - Med(2)) + Abs(CLng(bDATA(I, X + 1, Y - 2)) - Med(2)) + Abs(CLng(bDATA(I, X, Y - 2)) - Med(2)) + _
                 Abs(CLng(bDATA(I, X + 2, Y - 1)) - Med(2)) + Abs(CLng(bDATA(I, X + 1, Y - 1)) - Med(2)) + Abs(CLng(bDATA(I, X, Y - 1)) - Med(2)) + _
                 Abs(CLng(bDATA(I, X + 2, Y)) - Med(2)) + Abs(CLng(bDATA(I, X + 1, Y)) - Med(2)) + Abs(CLng(bDATA(I, X, Y)) - Med(2))
        Dev(3) = Abs(CLng(bDATA(I, X - 2, Y + 2)) - Med(3)) + Abs(CLng(bDATA(I, X - 1, Y + 2)) - Med(3)) + Abs(CLng(bDATA(I, X, Y + 2)) - Med(3)) + _
                 Abs(CLng(bDATA(I, X - 2, Y + 1)) - Med(3)) + Abs(CLng(bDATA(I, X - 1, Y + 1)) - Med(3)) + Abs(CLng(bDATA(I, X, Y + 1)) - Med(3)) + _
                 Abs(CLng(bDATA(I, X - 2, Y)) - Med(3)) + Abs(CLng(bDATA(I, X - 1, Y)) - Med(3)) + Abs(CLng(bDATA(I, X, Y)) - Med(3))
        Dev(4) = Abs(CLng(bDATA(I, X + 2, Y + 2)) - Med(4)) + Abs(CLng(bDATA(I, X + 1, Y + 2)) - Med(4)) + Abs(CLng(bDATA(I, X, Y + 2)) - Med(4)) + _
                 Abs(CLng(bDATA(I, X + 2, Y + 1)) - Med(4)) + Abs(CLng(bDATA(I, X + 1, Y + 1)) - Med(4)) + Abs(CLng(bDATA(I, X, Y + 1)) - Med(4)) + _
                 Abs(CLng(bDATA(I, X + 2, Y)) - Med(4)) + Abs(CLng(bDATA(I, X + 1, Y)) - Med(4)) + Abs(CLng(bDATA(I, X, Y)) - Med(4))
        vDev = 99999
        sDev = 0
        For j = 1 To 4
          If Dev(j) < vDev Then
            vDev = Dev(j)
            sDev = j
          End If
        Next j
        iDATA(I, X, Y) = Med(sDev)
      Next I
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Sharpen(ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim mf As Long, dF As Long

  mf = 24 + Factor
  dF = 8 + Factor
  For Y = 1 To PicProcess.ScaleHeight - 2
    For X = 1 To PicProcess.ScaleWidth - 2
      B = CLng(iDATA(0, X, Y - 1)) + CLng(iDATA(0, X - 1, Y)) + _
          CLng(iDATA(0, X + 1, Y)) + CLng(iDATA(0, X, Y + 1)) + _
          CLng(iDATA(0, X + 1, Y + 1)) + CLng(iDATA(0, X - 1, Y + 1)) + _
          CLng(iDATA(0, X + 1, Y - 1)) + CLng(iDATA(0, X - 1, Y - 1))
      B = (mf * CLng(iDATA(0, X, Y)) - 2 * B) \ dF
      G = CLng(iDATA(1, X, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + _
          CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X, Y + 1)) + _
          CLng(iDATA(1, X + 1, Y + 1)) + CLng(iDATA(1, X - 1, Y + 1)) + _
          CLng(iDATA(1, X + 1, Y - 1)) + CLng(iDATA(1, X - 1, Y - 1))
      G = (mf * CLng(iDATA(1, X, Y)) - 2 * G) \ dF
      R = CLng(iDATA(2, X, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + _
          CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X, Y + 1)) + _
          CLng(iDATA(2, X + 1, Y + 1)) + CLng(iDATA(2, X - 1, Y + 1)) + _
          CLng(iDATA(2, X + 1, Y - 1)) + CLng(iDATA(2, X - 1, Y - 1))
      R = (mf * CLng(iDATA(2, X, Y)) - 2 * R) \ dF
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
DrawOnDevice
End Sub
Sub NegativeImage()
  Dim X As Long, Y As Long
  For Y = 1 To PicProcess.ScaleHeight - 1
    For X = 1 To PicProcess.ScaleWidth - 1
      iDATA(0, X, Y) = 255 - iDATA(0, X, Y)
      iDATA(1, X, Y) = 255 - iDATA(1, X, Y)
      iDATA(2, X, Y) = 255 - iDATA(2, X, Y)
    Next X
    DoEvents
  Next Y
  DoEvents
DrawOnDevice
End Sub

Sub Engrave(ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 0 To PicProcess.ScaleHeight - 2
    For X = 0 To PicProcess.ScaleWidth - 2
      B = Abs(CLng(iDATA(0, X + 1, Y + 1)) - CLng(iDATA(0, X, Y)) + cB)
      G = Abs(CLng(iDATA(1, X + 1, Y + 1)) - CLng(iDATA(1, X, Y)) + cG)
      R = Abs(CLng(iDATA(2, X + 1, Y + 1)) - CLng(iDATA(2, X, Y)) + cR)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Emboss(ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  

  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 0 To PicInfo.bmHeight - 2
    For X = 0 To PicInfo.bmWidth - 2
      B = Abs(CLng(iDATA(0, X, Y)) - CLng(iDATA(0, X + 1, Y + 1)) + cB)
      G = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + cG)
      R = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + cR)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub Smooth()
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long

  For Y = 2 To PicInfo.bmHeight - 2
    For X = 2 To PicInfo.bmWidth - 2
      B = CLng(iDATA(0, X, Y)) + _
        CLng(iDATA(0, X - 1, Y)) + CLng(iDATA(0, X, Y - 1)) + _
        CLng(iDATA(0, X, Y + 1)) + CLng(iDATA(0, X + 1, Y))
      B = B \ 5
      G = CLng(iDATA(1, X, Y)) + _
        CLng(iDATA(1, X - 1, Y)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y))
      G = G \ 5
      R = CLng(iDATA(2, X, Y)) + _
        CLng(iDATA(2, X - 1, Y)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y))
      R = R \ 5
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub EdgeEnhance(ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim mf As Long, dF As Long

  mf = 9 + Factor
  dF = 1 + Factor
  For Y = 1 To UBound(iDATA, 3) - 1
    For X = 1 To UBound(iDATA, 2) - 1
      B = CLng(iDATA(0, X - 1, Y - 1)) + CLng(iDATA(0, X - 1, Y)) + _
        CLng(iDATA(0, X - 1, Y + 1)) + CLng(iDATA(0, X, Y - 1)) + _
        CLng(iDATA(0, X, Y + 1)) + CLng(iDATA(0, X + 1, Y - 1)) + _
        CLng(iDATA(0, X + 1, Y)) + CLng(iDATA(0, X + 1, Y + 1))
      B = (mf * CLng(iDATA(0, X, Y)) - B) \ dF
      G = CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + _
        CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + _
        CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1))
      G = (mf * CLng(iDATA(1, X, Y)) - G) \ dF
      R = CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + _
        CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + _
        CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1))
      R = (mf * CLng(iDATA(2, X, Y)) - R) \ dF
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Erode()
  Dim X As Long, Y As Long
  Dim V As Long
  Dim I As Long
  Dim vMin As Long
  
  For Y = 1 To UBound(iDATA, 3) - 1
    For X = 1 To UBound(iDATA, 2) - 1
      For I = 0 To 2
        vMin = 255
        V = CLng(bDATA(I, X - 1, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y - 1))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(I, X - 1, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(I, X - 1, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y + 1))
        If V < vMin Then vMin = V
        
        iDATA(I, X, Y) = vMin
      Next I
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub GreyScale()
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      B = iDATA(0, X, Y)
      G = iDATA(1, X, Y)
      R = iDATA(2, X, Y)
      B = Speed(R + G + B)
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = B
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub Contour(ByVal BackCol As Long)
'BackUp_Previous
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  GetRGB BackCol, cR, cG, cB
  For Y = 1 To UBound(iDATA, 3) - 1
    For X = 1 To UBound(iDATA, 2) - 1
      B = CLng(bDATA(0, X - 1, Y - 1)) + CLng(bDATA(0, X - 1, Y)) + _
          CLng(bDATA(0, X - 1, Y + 1)) + CLng(bDATA(0, X, Y - 1)) + _
          CLng(bDATA(0, X, Y + 1)) + CLng(bDATA(0, X + 1, Y - 1)) + _
          CLng(bDATA(0, X + 1, Y)) + CLng(bDATA(0, X + 1, Y + 1))
      G = CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X - 1, Y + 1)) + CLng(bDATA(1, X, Y - 1)) + _
          CLng(bDATA(1, X, Y + 1)) + CLng(bDATA(1, X + 1, Y - 1)) + _
          CLng(bDATA(1, X + 1, Y)) + CLng(bDATA(1, X + 1, Y + 1))
      R = CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X - 1, Y + 1)) + CLng(bDATA(2, X, Y - 1)) + _
          CLng(bDATA(2, X, Y + 1)) + CLng(bDATA(2, X + 1, Y - 1)) + _
          CLng(bDATA(2, X + 1, Y)) + CLng(bDATA(2, X + 1, Y + 1))
      B = 8 * CLng(bDATA(0, X, Y)) - B + cB
      G = 8 * CLng(bDATA(1, X, Y)) - G + cG
      R = 8 * CLng(bDATA(2, X, Y)) - R + cR
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub Relief()
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  
  For Y = 1 To UBound(iDATA, 3) - 1
    For X = 1 To UBound(iDATA, 2) - 1
      B = 2 * CLng(bDATA(0, X - 1, Y - 1)) + CLng(bDATA(0, X - 1, Y)) + _
          CLng(bDATA(0, X, Y - 1)) - CLng(bDATA(0, X, Y + 1)) - _
          CLng(bDATA(0, X + 1, Y)) - 2 * CLng(bDATA(0, X + 1, Y + 1))
      G = 2 * CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X, Y - 1)) - CLng(bDATA(1, X, Y + 1)) - _
          CLng(bDATA(1, X + 1, Y)) - 2 * CLng(bDATA(1, X + 1, Y + 1))
      R = 2 * CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X, Y - 1)) - CLng(bDATA(2, X, Y + 1)) - _
          CLng(bDATA(2, X + 1, Y)) - 2 * CLng(bDATA(2, X + 1, Y + 1))
      B = (CLng(bDATA(0, X, Y)) + B) \ 2 + 50
      G = (CLng(bDATA(1, X, Y)) + G) \ 2 + 50
      R = (CLng(bDATA(2, X, Y)) + R) \ 2 + 50
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub Brightness(ByVal Factor As Long)

  Dim X As Long, Y As Long
  Dim sF As Single
  
  sF = (Factor + 100) / 100
  For X = 0 To 255
    Speed(X) = X * sF
    If Speed(X) > 255 Then Speed(X) = 255
    If Speed(X) < 0 Then Speed(X) = 0
  Next X
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      iDATA(0, X, Y) = Speed(bDATA(0, X, Y))
      iDATA(1, X, Y) = Speed(bDATA(1, X, Y))
      iDATA(2, X, Y) = Speed(bDATA(2, X, Y))
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub OldStylePhoto()
    Dim X As Long, Y As Long
    Dim R As Long, G As Long, B As Long
    Dim TvALUE As Long
    
    For Y = 0 To UBound(iDATA, 3)
        For X = 0 To UBound(iDATA, 2)
        
        B = CLng(iDATA(0, X, Y)) * 0.114
        G = CLng(iDATA(1, X, Y)) * 0.587
        R = CLng(iDATA(2, X, Y)) * 0.299
        TvALUE = R + G + B
        TvALUE = RGB(TvALUE + 64, TvALUE + 24, TvALUE + 8)
        GetRGB TvALUE, R, G, B
        iDATA(0, X, Y) = B
        iDATA(1, X, Y) = G
        iDATA(2, X, Y) = R
        Next X
    Next Y
DrawOnDevice
End Sub

Sub Diffuse(ByVal Factor As Long)
'BackUp_Previous
  Dim X As Long, Y As Long
  Dim aX As Long, aY As Long
  Dim R As Long, G As Long, B As Long
  Dim hF As Long

  hF = Factor / 2
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      aX = Rnd * Factor - hF
      aY = Rnd * Factor - hF
      If X + aX < 1 Then aX = 0
      If X + aX > UBound(iDATA, 2) Then aX = 0
      If Y + aY < 1 Then aY = 0
      If Y + aY > UBound(iDATA, 3) Then aY = 0
      iDATA(0, X, Y) = iDATA(0, X + aX, Y + aY)
      iDATA(1, X, Y) = iDATA(1, X + aX, Y + aY)
      iDATA(2, X, Y) = iDATA(2, X + aX, Y + aY)
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Pixelize(ByVal PixSize As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim pX As Long, pY As Long
  Dim sX As Long, sY As Long
  Dim mC As Long
  
  B = 0: G = 0: R = 0
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      If ((X - 1) Mod PixSize) = 0 Then
        sX = ((X - 1) \ PixSize) * PixSize + 1
        sY = ((Y - 1) \ PixSize) * PixSize + 1
        B = 0: G = 0: R = 0: mC = 0
        For pX = sX To sX + PixSize - 1
          For pY = sY To sY + PixSize - 1
            If (pX <= UBound(iDATA, 2)) And (pY <= UBound(iDATA, 3)) Then
              B = B + CLng(bDATA(0, pX, pY))
              G = G + CLng(bDATA(1, pX, pY))
              R = R + CLng(bDATA(2, pX, pY))
              mC = mC + 1
            End If
          Next pY
        Next pX
        If mC > 0 Then
          B = B \ mC
          G = G \ mC
          R = R \ mC
        End If
      End If
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub SwapBank(ByVal Modo As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long

  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      B = CLng(bDATA(0, X, Y))
      G = CLng(bDATA(1, X, Y))
      R = CLng(bDATA(2, X, Y))
      Select Case Modo
        Case 1: 'RGB -> BRG
          iDATA(0, X, Y) = G
          iDATA(1, X, Y) = R
          iDATA(2, X, Y) = B
        Case 2: 'RGB -> GBR
          iDATA(0, X, Y) = R
          iDATA(1, X, Y) = B
          iDATA(2, X, Y) = G
        Case 3: 'RGB -> RBG
          iDATA(0, X, Y) = G
          iDATA(1, X, Y) = B
          iDATA(2, X, Y) = R
        Case 4: 'RGB -> BGR
          iDATA(0, X, Y) = R
          iDATA(1, X, Y) = G
          iDATA(2, X, Y) = B
        Case 5: 'RGB -> GRB
          iDATA(0, X, Y) = B
          iDATA(1, X, Y) = R
          iDATA(2, X, Y) = G
      End Select
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Saturation(ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim V As Long
  Dim sF As Single
    
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  
  sF = Factor / 100
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      B = CLng(bDATA(0, X, Y))
      G = CLng(bDATA(1, X, Y))
      R = CLng(bDATA(2, X, Y))
      V = Speed(B + G + R)
      B = B + sF * (B - V)
      G = G + sF * (G - V)
      R = R + sF * (R - V)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub GammaCorrection(ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim dB As Double, dG As Double, dR As Double
  Dim sF As Single
  Dim Max As Double, Min As Double, MM As Double
  Dim H As Double, S As Double, I As Double
  Dim cB As Double, cG As Double, cR As Double
  Dim Flo As Long
    
  sF = Factor / 100
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      'get data
      B = CLng(bDATA(0, X, Y))
      G = CLng(bDATA(1, X, Y))
      R = CLng(bDATA(2, X, Y))
      dB = B / 255
      dG = G / 255
      dR = R / 255
      'correct gamma
      dB = dB ^ (1 / sF)
      dG = dG ^ (1 / sF)
      dR = dR ^ (1 / sF)
      'set data
      B = dB * 255
      G = dG * 255
      R = dR * 255
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub

Sub Dilate()
  Dim X As Long, Y As Long
  Dim V As Long
  Dim I As Long
  Dim vMax As Long
  
  For Y = 1 To UBound(iDATA, 3) - 1
    For X = 1 To UBound(iDATA, 2) - 1
      For I = 1 To 3
        vMax = 0
        V = CLng(bDATA(I, X - 1, Y - 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X, Y - 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X + 1, Y - 1))
        If V > vMax Then vMax = V
        
        V = CLng(bDATA(I, X - 1, Y))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X, Y))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X + 1, Y))
        If V > vMax Then vMax = V
        
        V = CLng(bDATA(I, X - 1, Y + 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X, Y + 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(I, X + 1, Y + 1))
        If V > vMax Then vMax = V
        
        iDATA(I, X, Y) = vMax
      Next I
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Sub AddNoise(ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim R As Long, G As Long, B As Long
  Dim V As Long
  
  For Y = 0 To UBound(iDATA, 3)
    For X = 0 To UBound(iDATA, 2)
      B = CLng(bDATA(0, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      G = CLng(bDATA(1, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      R = CLng(bDATA(2, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(0, X, Y) = B
      iDATA(1, X, Y) = G
      iDATA(2, X, Y) = R
    Next X
    DoEvents
  Next Y
  DoEvents
  DrawOnDevice
End Sub
Function SelectAll()
DeleteObject mRGN
Select Case SelectionMode
    Case 1
        R.Top = 0: R.Left = 0: R.Bottom = PicDisp.ScaleHeight: R.Right = PicDisp.ScaleWidth
        mRGN = CreateRectRgn(R.Left, R.Top, R.Right, R.Bottom)
        ret = SelectObject(PicDisp.hdc, mRGN)
        sX = 0
        sY = 0
        eX = PicProcess.ScaleWidth
        eY = PicProcess.ScaleHeight
    Case 2
        R.Top = 0: R.Left = 0: R.Bottom = PicDisp.ScaleHeight: R.Right = PicDisp.ScaleWidth
        mRGN = CreateEllipticRgn(R.Left, R.Top, R.Right, R.Bottom)
        ret = SelectObject(PicDisp.hdc, mRGN)
    End Select
End Function

Sub FitToParent_VB6()
    Extender.Move (Parent.Width - (8 * Screen.TwipsPerPixelY) - UserControl.Width) \ 2, _
                             (Parent.Height - (30 * Screen.TwipsPerPixelY) - UserControl.Height) \ 2
End Sub

Private Sub DrawOnDevice()
    BackUp_Previous

    If PictureLoaded Then
        PicProcess.Picture = Nothing
        SetDIBits PicProcess.hdc, PicProcess.Image, 0, PicInfo.bmHeight, iDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS
        If SelectionMode = 2 Then 'if it is Elliptical selection
            mRGN = CreateEllipticRgnIndirect(R)
            ret = SelectObject(PicDisp.hdc, mRGN)
            Ellipse_PatternBrush = CreatePatternBrush(PicProcess.Image)
            ret = FillRgn(PicDisp.hdc, mRGN, Ellipse_PatternBrush)
            DeleteObject mRGN
            DeleteObject Ellipse_PatternBrush
        End If
        'BitBlt necessary rectangular region
        ret = BitBlt(PicDisp.hdc, sX, sY, eX - sX, eY - sY, _
                PicProcess.hdc, sX, sY, SRCCOPY)
    End If
    
    Erase iDATA
    Erase bDATA
    ReDim iDATA(3, PicProcess.ScaleWidth - 1, PicProcess.ScaleHeight - 1) As Byte
    ReDim bDATA(3, PicProcess.ScaleWidth - 1, PicProcess.ScaleHeight - 1) As Byte
    ret = GetDIBits(hdcNew, PicDisp.Image, 0, PicInfo.bmHeight, iDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS)
    ret = GetDIBits(hdcNew, PicDisp.Image, 0, PicInfo.bmHeight, bDATA(0, 0, 0), DIBInfo, DIB_RGB_COLORS)
End Sub

Public Property Get PStartX() As Long
    PStartX = sX
End Property

Public Property Let PStartX(ByVal vNewValue As Long)
    sX = vNewValue
End Property

Public Property Get PStartY() As Long
    PStartY = sY
End Property

Public Property Let PStartY(ByVal vNewValue As Long)

End Property

Public Property Get PLastX() As Long
    PLastX = eX
End Property

Public Property Let PLastX(ByVal vNewValue As Long)

End Property

Public Property Get PLastY() As Long
    PLastY = eY
End Property

Public Property Let PLastY(ByVal vNewValue As Long)

End Property


Public Property Get PicWidth() As Long
    PicWidth = PicInfo.bmWidth
End Property

Public Property Let PicWidth(ByVal vNewValue As Long)

End Property

Public Property Get PicHeight() As Long
    PicHeight = PicInfo.bmHeight
End Property

Public Property Let PicHeight(ByVal vNewValue As Long)

End Property

Public Property Get PicBitCount() As Long
    PicBitCount = PicInfo.bmBitsPixel
End Property

Public Property Let PicBitCount(ByVal vNewValue As Long)

End Property

'=========================================================================
'scrolling code
Private Sub Tops()
    Dim reading As Long
    Dim a As Long
    Dim L As Long
    
    reading = BV.Value
    a = PicDisp.Height \ 2
    L = -(a * reading) \ 100
    PicDisp.Top = L
End Sub

Private Sub Lefts()
    Dim reading As Long
    Dim a As Long
    Dim L As Long
    
    reading = VH.Value
    a = PicDisp.Width \ 2
    L = -(a * reading) \ 100
    PicDisp.Left = L
End Sub
Private Sub VH_Change()
    Call Lefts
End Sub

Private Sub VH_Scroll()
    Call Lefts
End Sub

Private Sub BV_Change()
    Call Tops
End Sub

Private Sub BV_Scroll()
    Call Tops
End Sub
'==========================================================================

