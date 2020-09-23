Attribute VB_Name = "modWebcam"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long

Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Public mCapHwnd As Long

Public SelColor As Boolean

Public Const WM_CAP_CONNECT As Long = 1034
Public Const WM_CAP_DISCONNECT As Long = 1035
Public Const WM_CAP_GET_FRAME As Long = 1084
Public Const WM_CAP_COPY As Long = 1054

Public Const WM_USER = &H400
Public Const WM_CAP_START = WM_USER
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46
Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
