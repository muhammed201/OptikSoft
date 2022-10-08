VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Similasyon"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   9960
      Picture         =   "Form1.frx":511A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   4680
   End
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4680
   End
   Begin VB.PictureBox picCapture 
      BackColor       =   &H80000006&
      Height          =   3495
      Left            =   4800
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Webcam'a baðlan"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Resmi Çek"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.ListBox lstDevices 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label sayi 
      BackColor       =   &H80000009&
      Caption         =   "1. Resimi Çekiniz"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   6480
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   5970
      Left            =   3840
      Picture         =   "Form1.frx":5584
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Aygýt Seçiniz"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const WM_CAP As Integer = &H400

Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM As Integer = 1

Dim iDevice As Long  ' Current device ID
Dim hHwnd As Long ' Handle to preview window

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean

Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
    (ByVal lpszWindowName As String, ByVal dwStyle As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Integer, ByVal hWndParent As Long, _
    ByVal nID As Long) As Long

Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, _
    ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, _
    ByVal cbVer As Long) As Boolean
    Private X As Byte

Private Sub cmdSave_Click()
    Dim bm As Image
    Dim ism As String
    X = X + 1
    sayi.Caption = X + 1 & ". Resimi Çekiniz"

    '
    ' Copy image to clipboard
    '
    
    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
    ClosePreviewWindow

    picCapture.Picture = Clipboard.GetData
    
    'CommonDialog1.CancelError = True
    'CommonDialog1.FileName = "Webcam1"
    'CommonDialog1.Filter = "Bitmap |*.bmp"
     OpenPreviewWindow '
    
    On Error GoTo NoSave
    'CommonDialog1.ShowSave
    ism = "similasyon" & "(" & X & ").jpg"
    SavePicture picCapture.Image, App.Path & "\" & ism  'CommonDialog1.FileName
    If X >= 4 Then
    X = 0
    cmdSave.Enabled = False
    cmdStart.Enabled = True
    sayi.Caption = "Çekim Bitmiþtir"
    Unload Me
    goster.Show
    End If
    
NoSave:
    
End Sub

Private Sub cmdStart_Click()
    ShockwaveFlash1.Visible = False
    iDevice = lstDevices.ListIndex
    OpenPreviewWindow
End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo hata
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True
ShockwaveFlash1.Movie = App.Path & "/flash/banner.swf"
    LoadDeviceList
    
    If lstDevices.ListCount > 0 Then
        lstDevices.Selected(0) = True
    Else
    cmdStart.Enabled = False
        lstDevices.AddItem ("No Device Available")
    End If
    cmdSave.Enabled = False
hata:
    Resume Next

End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer As String
    Dim iReturn As Boolean
    Dim X As Long
    
    X = 0
    strName = Space(100)
    strVer = Space(100)
    '
    ' Load name of all avialable devices into the lstDevices
    '

    Do
        '
        '   Get Driver name and version
        '
        iReturn = capGetDriverDescriptionA(X, strName, 100, strVer, 100)

        '
        ' If there was a device add device name to the list
        '
        If iReturn Then lstDevices.AddItem Trim$(strName)
        X = X + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()
    '
    ' Open Preview window in picturebox
    '
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, _
        480, picCapture.hwnd, 0)

    '
    ' Connect to device
    '
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
        '
        'Set the preview scale
        '
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0

        '
        'Set the preview rate in milliseconds
        '
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0

        '
        'Start previewing the image from the camera
        '
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0

        '
        ' Resize window to fit in picturebox
        '
        SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, _
                SWP_NOMOVE Or SWP_NOZORDER

        cmdSave.Enabled = True
        
        cmdStart.Enabled = False
    Else
        '
        ' Error connecting to device close window
        '
        DestroyWindow hHwnd

        cmdSave.Enabled = False
    End If
 End Sub

Private Sub ClosePreviewWindow()
    '
    ' Disconnect from device
    '
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0
        '
    ' close window
    '

    DestroyWindow hHwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
        X = 0
        ClosePreviewWindow
        trmUnload.Enabled = True
        AnaMenu.Show
        t = 0
    
End Sub
Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 245 Then
    trmLoad.Enabled = False
        Else
    t = t + 10
  End If
End Sub

Private Sub trmUnload_Timer()
SetTransparent hwnd, t
  If trmLoad.Enabled = True Then trmLoad.Enabled = False
  If t <= 0 Then
    trmUnload.Enabled = False
    Unload Me
        Else
    t = t - 4
  End If
  End Sub





