VERSION 5.00
Begin VB.Form buyuk 
   Caption         =   "Ön Ýzleme"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H007A83CB&
      Caption         =   "Seçmiþ Olduðunuz Ýmajýn Büyültülmüþ hali"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      Begin VB.Image Image2 
         Height          =   6720
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Picture         =   "buyuk.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "buyuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Local Error GoTo hata
Image2.Picture = LoadPicture(badres)
badres = Null
hata:
If Err = 76 Then
MsgBox "Görüntülenecek imaj yok", vbCritical
Else
Resume Next
End If
End Sub

