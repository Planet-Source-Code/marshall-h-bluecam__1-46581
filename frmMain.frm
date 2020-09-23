VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlueCam"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9960
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOn 
      Caption         =   "Filter On"
      Height          =   270
      Left            =   1440
      TabIndex        =   10
      Top             =   3885
      Width           =   945
   End
   Begin VB.PictureBox picTmpOutput 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   2085
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1515
      Top             =   4695
   End
   Begin VB.FileListBox lstPictures 
      Height          =   1260
      Left            =   5010
      Pattern         =   "*.bmp;*.jpg;*.gif"
      TabIndex        =   9
      Top             =   3855
      Width           =   4845
   End
   Begin VB.VScrollBar scrRange 
      Height          =   1290
      LargeChange     =   5
      Left            =   2775
      Max             =   0
      Min             =   255
      SmallChange     =   2
      TabIndex        =   8
      Top             =   3840
      Value           =   50
      Width           =   270
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   150
      TabIndex        =   7
      Top             =   4290
      Width           =   1125
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   3840
      Width           =   1125
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Format"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   4755
      Width           =   1125
   End
   Begin VB.PictureBox picPalette 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   990
      Left            =   3150
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   4
      Top             =   3840
      Width           =   1755
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   270
      Left            =   3150
      ScaleHeight     =   210
      ScaleWidth      =   1710
      TabIndex        =   3
      Top             =   4890
      Width           =   1770
   End
   Begin VB.PictureBox picBackGround 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3660
      Left            =   5010
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   120
      Width           =   4860
   End
   Begin VB.PictureBox picOutput 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3645
      Left            =   135
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------
'-          Blue Screen Cam        -
'-----------------------------------
'-         (c) 2003 Marshall       -
'-----------------------------------
'
Private Sub cmdFormat_Click()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    'Setup a capture window
    mCapHwnd = capCreateCaptureWindow("WebCap", 0, 0, 0, 320, 240, Me.hwnd, 0)
    'Connect to capture device
    DoEvents: SendMessage mCapHwnd, WM_CAP_CONNECT, 0, 0
    SendMessage mCapHwnd, WM_CAP_SET_PREVIEW, 0, 0
    tmrMain.Enabled = True
End Sub

Private Sub cmdStop_Click()
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    tmrMain.Enabled = False
    'Make sure to disconnect from capture source
    DoEvents: SendMessage mCapHwnd, WM_CAP_DISCONNECT, 0, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStop.Enabled = False Then
        'Make sure to disconnect from capture source - if it is connected upon termination the program can become unstable
        DoEvents: SendMessage mCapHwnd, WM_CAP_DISCONNECT, 0, 0
    End If
End Sub

Private Sub lstPictures_Click()
    picBackGround.Picture = LoadPicture(lstPictures.FileName)
End Sub

Private Sub picOutput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picColor.BackColor = GetPixel(picOutput.hdc, X, Y)
End Sub

Private Sub picPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelColor = True
End Sub

Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If SelColor Then picColor.BackColor = GetPixel(picPalette.hdc, X, Y)
End Sub

Private Sub picPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelColor = False
End Sub

Private Sub tmrMain_Timer()
On Error Resume Next
    'Get current frame
    SendMessage mCapHwnd, WM_CAP_GET_FRAME, 0, 0
    
    'Copy current frame to Clipboard
    SendMessage mCapHwnd, WM_CAP_COPY, 0, 0
    
    'Put Clipboard's data to picOutput
    picTmpOutput.Picture = Clipboard.GetData
    
    If chkOn.Value = 0 Then GoTo SkipFilter
    
    'Replace black pixels
    For Y = 0 To picTmpOutput.ScaleHeight
    For X = 0 To picTmpOutput.ScaleWidth
        CurrentColor = GetPixel(picTmpOutput.hdc, X, Y)
        Range = RGB(scrRange.Value, scrRange.Value, scrRange.Value)
        Color = picColor.BackColor
        
        If CurrentColor > Color - Range And CurrentColor < Color + Range Then
            SetPixel picTmpOutput.hdc, X, Y, GetPixel(picBackGround.hdc, X, Y)
        End If
    Next
    Next

SkipFilter:
    picOutput.PaintPicture picTmpOutput.Image, 0, 0
End Sub
