VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   Icon            =   "dmclock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   143
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3315
      Top             =   5835
   End
   Begin VB.PictureBox picdst 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   -15
      Picture         =   "dmclock.frx":0CCA
      ScaleHeight     =   330
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   210
      Width           =   3015
      Begin VB.Image imgdown 
         Height          =   105
         Left            =   1785
         Picture         =   "dmclock.frx":0D10
         Top             =   165
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image imgup 
         Height          =   105
         Left            =   1920
         Picture         =   "dmclock.frx":0DFA
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.PictureBox srcpic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      Picture         =   "dmclock.frx":0EE4
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   0
      Top             =   5895
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Image imgupdown 
      Height          =   105
      Left            =   1800
      Picture         =   "dmclock.frx":4256
      Top             =   30
      Width           =   120
   End
   Begin VB.Label lblclose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   3
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private tUpDown As Boolean

Public ViewOption As Integer

Private Sub UpDown()
    Select Case tUpDown
        Case True
            tUpDown = False
            imgupdown.Picture = imgdown.Picture
            frmmain.Height = 195
        Case False
            tUpDown = True
            imgupdown.Picture = imgup.Picture
            frmmain.Height = 540
    End Select
End Sub

Function MoveForm(mHwnd As Long)
    ReleaseCapture
    SendMessage mHwnd, WM_NCLBUTTONDOWN, HTCAPTION, True
End Function

Sub UpdateDisplay(sTime As String)
Dim I As Long
Dim CH1, CH2 As String
    picdst.Cls
    For I = 1 To Len(sTime)
        CH1 = Mid(sTime, I, 1)
        CH2 = Mid(sTime, I, 2)
        If IsNumeric(CH1) Then
            BitBlt picdst.hDC, I * 16 - 16, 0, 16, 21, srcpic.hDC, Val(CH1) * 16, 0, vbSrcCopy
        ElseIf (CH1 = ":" Or CH1 = "/") Then
             BitBlt picdst.hDC, I * 16 - 16, 0, 16, 21, srcpic.hDC, 192, 0, vbSrcCopy
        ElseIf CH2 = "AM" Then
            BitBlt picdst.hDC, I * 16 - 32, 0, 16, 21, srcpic.hDC, 160, 0, vbSrcCopy
        ElseIf CH2 = "PM" Then
            BitBlt picdst.hDC, I * 16 - 32, 0, 16, 21, srcpic.hDC, 176, 0, vbSrcCopy
        End If
    Next
    picdst.Refresh
    CH1 = ""
    CH2 = ""
    I = 0
    
End Sub
Private Sub Form_Load()
    frmmain.Width = 2160
    ViewOption = 0
    lbltitle.Caption = "Clock View"
    UpDown
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing ' unload the form from memory
    Set frmmenu = Nothing
    ViewOption = 0
End Sub

Private Sub imgupdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpDown
End Sub

Private Sub lblclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload frmmenu ' unload the menu form
    Unload frmmain ' unload this form
End Sub

Private Sub lbltitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveForm frmmain.hwnd
    End If
    
End Sub

Private Sub picdst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu frmmenu.mnuView
    End If
    
End Sub

Private Sub picdst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveForm frmmain.hwnd
    End If
End Sub

Private Sub Timer1_Timer()
    If ViewOption = 0 Then
        'lbltitle.Width = 144 ' Set the labels width
        UpdateDisplay Format(Time, "hh:mm:ss AM/PM") ' Update display with formated time
    End If
    
    If ViewOption = 1 Then
       ' lbltitle.Width = 144 ' Set the labels width
        UpdateDisplay Format(Date, "dd:mm:yyyy") ' Update display with formated time
    End If
End Sub
