VERSION 5.00
Object = "{D0CD41D3-0FA6-4C72-8799-854FB0C7CC6F}#1.0#0"; "AResizeLite.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form KURLAR 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Döviz Kurlarý"
   ClientHeight    =   8490
   ClientLeft      =   900
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "doviz.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Lütfen Döviz Kurlarýný Öðrenmek Ýstediðiniz Tarihi Giriniz"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      Begin VB.Timer trmUnload 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer trmLoad 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   600
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Height          =   855
         Left            =   9960
         Picture         =   "doviz.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin MSACAL.Calendar takvim 
         Height          =   2295
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   16249582
         Year            =   2006
         Month           =   12
         Day             =   26
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Tur"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Tur"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Tur"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         TabIndex        =   5
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Türkiye Cumhuriyeti Merkez Bankasý Döviz Kurlarý "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   7005
      End
   End
   Begin SHDocVwCtl.WebBrowser adres 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   10186
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ActiveResizeLiteCtl.ActiveResizeLite ActiveResizeLite1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   2
      ScreenHeight    =   768
      ScreenWidth     =   1024
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1024
      FormHeightDT    =   9000
      FormWidthDT     =   12000
      FormScaleHeightDT=   8490
      FormScaleWidthDT=   11880
   End
End
Attribute VB_Name = "KURLAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
trmUnload.Enabled = True
End Sub

Private Sub Form_Load()
takvim.Value = Now
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True
Me.Caption = takvim.Value & " Türkiye Cumhuriyeti Merkez Bankasý Döviz Kurlarý Bilgileri ~~~~~Optik Soft 1.0"
Label2.Caption = takvim.Value & " Tarihli Dövizlerin Kur Bilgileri"
adres.Navigate "http://www.tcmb.gov.tr/kurlar/today"
End Sub

Private Sub Form_Unload(Cancel As Integer)
trmUnload.Enabled = True
AnaMenu.Show
End Sub

Private Sub takvim_Click()
Dim yil, ay
Dim gun1
yil = Year(takvim.Value)
ay = Month(takvim.Value)
If Len(ay) < 2 Then
ay = 0 & ay
End If
gun1 = Mid(takvim.Value, 1, 2)
If takvim.Value > Now Then
MsgBox "Bugünün Tarihinden Büyük Giremezsiniz", vbApplicationModal, "Optik Soft 1.0~Kur Bilgileri"
adres.Navigate "http://www.tcmb.gov.tr/kurlar/today"
takvim.Value = Now
Exit Sub
End If

Label2.Caption = takvim.Value & "Tarihli Dövizlerin Kur Bilgileri"
Me.Caption = takvim.Value & " Türkiye Cumhuriyeti Merkez Bankasý Döviz Kurlarý Bilgileri ~~~~~Optik Soft 1.0"
adres.Navigate "http://www.tcmb.gov.tr/kurlar/" & yil & ay & "/" & gun1 & ay & yil & ".html"
End Sub

Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 240 Then
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
    t = t - 10
  End If
  End Sub


