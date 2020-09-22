VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm mdiEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Octa"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   6750
   Icon            =   "mdiEdit.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FilterChecker 
      Interval        =   10
      Left            =   4920
      Top             =   1440
   End
   Begin MSComctlLib.StatusBar ctlStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4815
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog C2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSelectA 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUL 
         Caption         =   "Undo Last"
      End
      Begin VB.Menu mnuUA 
         Caption         =   "Undo All"
      End
   End
   Begin VB.Menu mnuFilterM 
      Caption         =   "F&ilters"
      Begin VB.Menu mnuAddNoise 
         Caption         =   "Add Noise"
      End
      Begin VB.Menu mnuAqua 
         Caption         =   "Aqua"
      End
      Begin VB.Menu mnuBrighter 
         Caption         =   "Brighter"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDilate 
         Caption         =   "Dilate"
      End
      Begin VB.Menu mnuContour 
         Caption         =   "Contour"
      End
      Begin VB.Menu mnuDarker 
         Caption         =   "Darker"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuDiffuse 
         Caption         =   "Diffuse"
      End
      Begin VB.Menu MnuDiffuseM 
         Caption         =   "Diffuse More"
      End
      Begin VB.Menu mnuErode 
         Caption         =   "Erode"
      End
      Begin VB.Menu mnuEngrave 
         Caption         =   "Engrave"
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "Emboss"
      End
      Begin VB.Menu mnuEEn 
         Caption         =   "Edge Enhance"
      End
      Begin VB.Menu mnuGCo 
         Caption         =   "Gamma Correction"
      End
      Begin VB.Menu mnuGscale 
         Caption         =   "Gray Scale"
      End
      Begin VB.Menu mnuISa 
         Caption         =   "Increase Saturation"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuDSa 
         Caption         =   "Decrease Saturation"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSmooth 
         Caption         =   "Smooth"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSharpen 
         Caption         =   "Sharpen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSharpenMore 
         Caption         =   "Sharpen More"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuNI 
         Caption         =   "Negative Image"
      End
      Begin VB.Menu mnuPixelize 
         Caption         =   "Pixelize"
      End
      Begin VB.Menu mnuRelief 
         Caption         =   "Relief"
      End
      Begin VB.Menu mnuYSP 
         Caption         =   "Yellow Photo Effect"
      End
      Begin VB.Menu mnuRGBBGR 
         Caption         =   "RGB -> BGR"
      End
      Begin VB.Menu MnuRGBBRG 
         Caption         =   "RGB -> BRG"
      End
      Begin VB.Menu mnuRGBGBR 
         Caption         =   "RGB -> GBR"
      End
      Begin VB.Menu mnuRGBGRB 
         Caption         =   "RGB -> GRB"
      End
      Begin VB.Menu mnuRGBRBG 
         Caption         =   "RGB -> RBG"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "mdiEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FilterChecker_Timer()
    If ActiveForm Is Nothing Then
        mnuFilterM.Enabled = False
        mnuSave.Enabled = False
        mnuSelectA.Enabled = False
        mnuEdit.Enabled = False
        FilterChecker.Enabled = False
        Exit Sub
    End If
        mnuFilterM.Enabled = True
        mnuEdit.Enabled = True
        mnuSave.Enabled = True
        mnuSelectA.Enabled = True
        FilterChecker.Enabled = False
End Sub

Private Sub mnuAddNoise_Click()
    ActiveForm.PicEdit1.AddNoise 50
End Sub

Private Sub mnuAqua_Click()
    ActiveForm.PicEdit1.Aqua
End Sub

Private Sub mnuDarker_Click()
    ActiveForm.PicEdit1.Brightness -20
End Sub

Private Sub mnuDiffuse_Click()
    ActiveForm.PicEdit1.Diffuse 6
End Sub

Private Sub MnuDiffuseM_Click()
    ActiveForm.PicEdit1.Diffuse 12
End Sub

Private Sub mnuDilate_Click()
    ActiveForm.PicEdit1.Dilate
End Sub

Private Sub mnuDSa_Click()
    ActiveForm.PicEdit1.Saturation -20
End Sub

Private Sub mnuErode_Click()
    ActiveForm.PicEdit1.Erode
End Sub

Private Sub mnuISa_Click()
    ActiveForm.PicEdit1.Saturation 15
End Sub

Private Sub MnuRGBBRG_Click()
    ActiveForm.PicEdit1.SwapBank 1
End Sub

Private Sub mnuRGBGBR_Click()
    ActiveForm.PicEdit1.SwapBank 2
End Sub

Private Sub mnuRGBRBG_Click()
    ActiveForm.PicEdit1.SwapBank 3
End Sub

Private Sub mnuRGBBGR_Click()
    ActiveForm.PicEdit1.SwapBank 4
End Sub

Private Sub mnuRGBGRB_Click()
    ActiveForm.PicEdit1.SwapBank 5
End Sub

Private Sub mnuPixelize_Click()
    ActiveForm.PicEdit1.Pixelize 3
End Sub

Private Sub mnuSHis_Click()
    ActiveForm.PicEdit1.StretchHistogram
End Sub

Private Sub mnuUA_Click()
    ActiveForm.PicEdit1.UndoAll
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuBrighter_Click()
    ActiveForm.PicEdit1.Brightness 10
End Sub

Private Sub mnuContour_Click()
    ActiveForm.PicEdit1.Contour RGB(255, 255, 255)
End Sub

Private Sub mnuGscale_Click()
    ActiveForm.PicEdit1.GreyScale
End Sub

Private Sub mnuPixel_Click()
    ActiveForm.PicEdit1.Pixelize 2
End Sub

Private Sub mnuRelief_Click()
    ActiveForm.PicEdit1.Relief
End Sub

Private Sub mnuSelectA_Click()
    ActiveForm.PicEdit1.SelectAll
End Sub

Private Sub mnuSharpen_Click()
    ActiveForm.PicEdit1.Sharpen 2
End Sub

Private Sub mnuSharpenMore_Click()
    ActiveForm.PicEdit1.Sharpen 3
End Sub

Private Sub mnuEngrave_Click()
    ActiveForm.PicEdit1.Engrave RGB(0, 120, 120)
End Sub

Private Sub mnuNI_Click()
    ActiveForm.PicEdit1.NegativeImage
End Sub
Private Sub mnuEmboss_Click()
    ActiveForm.PicEdit1.Emboss RGB(0, 120, 120)
End Sub

Private Sub mnuSmooth_Click()
    ActiveForm.PicEdit1.Smooth
End Sub

Private Sub mnuEEn_Click()
    ActiveForm.PicEdit1.EdgeEnhance 1
End Sub

Private Sub mnuUL_Click()
    ActiveForm.PicEdit1.UndoLast
End Sub

Private Sub mnuYSP_Click()
    ActiveForm.PicEdit1.OldStylePhoto
End Sub

Private Sub mnuGCo_Click()
    On Error GoTo GammaError
    Dim GFactor As Long
    GFactor = InputBox("Gamma Factor")
    If GFactor <= 0 Or Not (IsNumeric(GFactor)) Then Exit Sub
    ActiveForm.PicEdit1.GammaCorrection GFactor
    Exit Sub
GammaError:
MsgBox "Only type in Numbers."
End Sub



'===================================================================
Private Sub mnuSave_Click()

Dim Filename As String
    
    With C2
        .DialogTitle = "Save Bitmap"
        .Filter = "Bitmap|*.bmp"
        .Filename = ""
    End With
    
    C2.ShowSave
    
    If C2.Filename = "" Then Exit Sub
    Filename = C2.Filename
    ActiveForm.PicEdit1.SaveImage Filename
    
End Sub

Private Sub MDIForm_Load()
    With Me
        .Height = 9600
        .Width = 12800
    End With
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim Filename As String
    Dim frm As frmEdit
    Static Count As Long
    Count = Count + 1
    Set frm = New frmEdit
    frm.Width = 3615
    frm.Height = 3945
    
    C2.DialogTitle = "Open"
    C2.Filename = ""
    C2.Filter = "Image Files(bmp,jpg,gif)|*.bmp;*.jpg;*.gif"
    If FilterChecker.Enabled = False Then FilterChecker.Enabled = True
    C2.ShowOpen
    If C2.Filename = "" Then
        Unload frm
        Count = Count - 1
        Exit Sub
    End If
    Filename = C2.Filename
    frm.PicEdit1.OpenImage Filename
    frm.Caption = Filename
End Sub

