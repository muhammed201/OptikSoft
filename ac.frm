VERSION 5.00
Begin VB.Form ac 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call veri_ac(False, False)
Call tablo_ac("select * from izin")
If tablo!aktif = True Then
Unload Me
Giris.Show
Else
Unload Me
AnaMenu.Show
End If

End Sub
