VERSION 5.00
Begin VB.Form frmCrop 
   Caption         =   "Click to Select"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "frmCrop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton cmdReset 
         Caption         =   "ReSelect"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Use this Selection"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCanx 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.PictureBox picCrop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   2
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   0
      MouseIcon       =   "frmCrop.frx":014A
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   0
      Width           =   2355
   End
End
Attribute VB_Name = "frmCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GotIt As Boolean

Private Sub cmdCanx_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 Set frmMain.picBMP.Picture = picCrop.Image
 Unload Me
End Sub

Private Sub cmdReset_Click()
 GotIt = False
 cmdOK.Enabled = False
 cmdReset.Enabled = False
End Sub

Private Sub Form_Resize()
 Frame1.Move 0, ScaleHeight - Frame1.Height
End Sub

Private Sub picSrc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 cmdOK.Enabled = True
 If Not GotIt Then
  Set picCrop.Picture = LoadPicture
  BitBlt picCrop.hdc, 0, 0, 13, 13, _
    picSrc.hdc, x - 0, y - 0, vbSrcCopy
  picCrop.Refresh
 End If
End Sub

Private Sub picSrc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 GotIt = True
 cmdReset.Enabled = True
End Sub
