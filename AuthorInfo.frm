VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Version Info"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
      Begin VB.Label Label5 
         Caption         =   "*Added new filters, and add scroll bar resume buttom."
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Version : 0.05"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Memory Status"
      Height          =   1575
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1440
         Top             =   1920
      End
      Begin VB.Label lblMemAva 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Available"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblMemTotal 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   120
      Picture         =   "AuthorInfo.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Alexander Cheng @ 2003"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
        dwLength As Long 'MEMORYSTATUSµ²ºc¤j¤p
        dwMemoryLoad As Long '¨t²Î°O¾ÐÅé¤u§@­t²üªº¦ô­p ¤¶©ó0~100¶¡ _
        ³o­Ó­È¥u¨Ñ¤ñ¸û¥Î 95/98 NTºtºâªk³£¤£¦P ±N¨Ó¤]¥i¯à§ï
        dwTotalPhys As Long ' ¹êÅé°O¾ÐÅé¤j¤p
        dwAvailPhys As Long ' ³Ñ¾lªº¹êÅé°O¾ÐÅé
        dwTotalPageFile As Long '°O¾ÐÅé­¶¥iÀx¦sªº¦ì¤¸²Õ¼Æ
        dwAvailPageFile As Long '³Ñ¾l°O¾ÐÅé­¶¤j¤p
        dwTotalVirtual As Long '¨C­Ó³B²zµ{§Ç¥i¥Î¦ì§}¤j¤p
        dwAvailVirtual As Long
End Type
Const mbmb = 1048576
Dim lpBuffer As MEMORYSTATUS
Dim lTotalMem As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    GlobalMemoryStatus lpBuffer
    lblMemAva.Caption = Format(lpBuffer.dwAvailPhys / mbmb, "#.##") & "-MB"
    lblMemTotal.Caption = Format(lpBuffer.dwTotalPhys / mbmb, "#.##") & "-MB"
End Sub
