VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4785
   Begin VB.TextBox txtText 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "A"
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdCanx 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Set Font"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFont 
      Caption         =   "MS Sans Serif Regular 8 Point"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanx_Click()
 gCancelled = True
 Unload Me
End Sub

Private Sub cmdFont_Click()
 With frmMain.CD
  .CancelError = True
  .Flags = cdlCFBoth
  On Error GoTo Canx
  .ShowFont
  FontDesc = GetFontDesc(frmMain.CD)
  lblFont.Caption = FontDesc
  txtText.FontName = .FontName
  txtText.FontBold = .FontBold
  txtText.FontItalic = .FontItalic
  txtText.FontSize = .FontSize
  gFontName = .FontName
  gFontBold = .FontBold
  gFontItalic = .FontItalic
  gFontSize = .FontSize
 End With
Canx:
End Sub
Private Function GetFontDesc(CD As CommonDialog)
 With CD
  GetFontDesc = .FontName
  If .FontBold = False And .FontItalic = False Then
   GetFontDesc = GetFontDesc & " Regular "
  ElseIf .FontBold = True And .FontItalic = True Then
   GetFontDesc = GetFontDesc & " Bold Italic "
  ElseIf .FontBold Then
   GetFontDesc = GetFontDesc & " Bold "
  ElseIf .FontItalic Then
   GetFontDesc = GetFontDesc & " Italic "
  End If
  GetFontDesc = GetFontDesc & .FontSize & " Points"
 End With
End Function

Private Sub cmdOK_Click()
 gText = txtText.Text
 Unload Me
End Sub

Private Sub Form_Load()
 gCancelled = False
 If Len(FontDesc) = 0 Then
  'set default
  FontDesc = "MS Sans Serif Regular 8 Point"
  lblFont.Caption = FontDesc
  gFontName = "MS Sans Serif"
  gFontBold = False
  gFontItalic = False
  gFontSize = 8
 Else
  'reload previously used font
  lblFont.Caption = FontDesc
  txtText.FontName = gFontName
  txtText.FontBold = gFontBold
  txtText.FontItalic = gFontItalic
  txtText.FontSize = gFontSize
 End If
End Sub
