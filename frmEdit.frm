VERSION 5.00
Begin VB.Form frmEdit 
   Appearance      =   0  'Flat
   Caption         =   "Waiting..."
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   3855
   Begin Project1.PicEdit PicEdit1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    PicEdit1.ReleaseAllStuff
    mdiEdit.ctlStatusBar.SimpleText = ""
    mdiEdit.FilterChecker.Enabled = True
End Sub

Private Sub Form_Resize()
    PicEdit1.FitToParent_VB6
End Sub

Private Sub PicEdit1_ImageLoaded(bmType As Long, bmWidth As Long, bmHeight As Long, bmBitsPixel As Integer)
    mdiEdit.ctlStatusBar.SimpleText = PicEdit1.PicHeight & "*" & PicEdit1.PicWidth & "  " & PicEdit1.PicBitCount & " bit"
End Sub

Private Sub PicEdit1_Resize(Width As Long, Height As Long)
    Me.Height = (PicEdit1.PictureDisplay.ScaleHeight + 30) * Screen.TwipsPerPixelY + 495
    Me.Width = (PicEdit1.PictureDisplay.ScaleWidth + 8) * Screen.TwipsPerPixelX + 495
End Sub
