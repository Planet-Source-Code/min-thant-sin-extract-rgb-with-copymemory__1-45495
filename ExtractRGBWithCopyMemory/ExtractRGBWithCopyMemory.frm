VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Extract RGB"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   75
      ScaleHeight     =   184
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   1
      Top             =   75
      Width           =   2640
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   2775
      MousePointer    =   2  'Cross
      Picture         =   "ExtractRGBWithCopyMemory.frx":0000
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   0
      Top             =   75
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "Move the mouse over the color pic while pressing left mouse button"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2775
      TabIndex        =   2
      Top             =   3000
      Width           =   2565
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////
'Extract RGB with CopyMemory by Min Thant Sin
'Questions? Comments? Suggestions?
'Feel free to e-mail me at < minsin999@hotmail.com >
'//////////////////////////////////////////////////////////////////////////////////////

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub picSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ChangeColor CLng(X), CLng(Y)
End Sub

Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    ChangeColor CLng(X), CLng(Y)
      
End Sub

Sub ChangeColor(ByVal X As Long, ByVal Y As Long)
    Dim ColorValue As Long
    Dim Red As Integer, Green As Integer, Blue As Integer
      
    ColorValue = GetPixel(picSource.hDC, X, Y)
      
    'Extract red, green, blue color components.
    Call ExtractRGBFromColor(ColorValue, Red, Green, Blue)
      
    picDest.BackColor = RGB(Red, Green, Blue)
End Sub

Sub ExtractRGBFromColor(ByVal ColorValue As Long, Red As Integer, Green As Integer, Blue As Integer)
      If ColorValue < 0 Then Exit Sub
      
      Dim bArray(1 To 4) As Byte
      
      Call CopyMemory(bArray(1), ColorValue, Len(ColorValue))
      
      Red = bArray(1)
      Green = bArray(2)
      Blue = bArray(3)
End Sub
