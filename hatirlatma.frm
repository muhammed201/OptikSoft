VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RSEFileRead.ocx"
Begin VB.Form hatirlatma 
   Caption         =   "Form2"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form2"
   ScaleHeight     =   6390
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Hatýrlatma"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      Begin VB.CommandButton Command2 
         Caption         =   "Ýptal"
         Height          =   495
         Left            =   5040
         TabIndex        =   6
         Top             =   5280
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tamam"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hatýrlatýlacak Notunuzu Giriniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   4695
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4215
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7435
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"hatirlatma.frx":0000
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hatýrlatma Tarihini Giriniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   4455
         Begin MSACAL.Calendar Calendar1 
            Height          =   4095
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   7223
            _StockProps     =   1
            BackColor       =   255
            Year            =   2006
            Month           =   12
            Day             =   28
            DayLength       =   1
            MonthLength     =   1
            DayFontColor    =   16711680
            FirstDay        =   7
            GridCellEffect  =   1
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483624
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   0
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Tur"
               Size            =   9.75
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
               Weight          =   700
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
      End
   End
End
Attribute VB_Name = "hatirlatma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
If Calendar1.Value < Now Then
MsgBox "Geçmiþ için hatýrlatma yatamazsýnýz :( ", vbExclamation
Calendar1.Value = Now
Exit Sub
End If
End Sub

Private Sub Command1_Click()
Call veri_ac(False, False)
Call tablo_ac("select * from hatirlatma")
tablo.AddNew
tablo!tarih = Calendar1.Value
tablo!dipnot = RichTextBox1.Text
tablo.Update
MsgBox "Hatýrlatmanýz kaydedilmiþtir", vbInformation
RichTextBox1 = Empty
Calendar1.Value = Now
tablo.Close
veri.Close
End Sub

Private Sub Command2_Click()
End Sub
