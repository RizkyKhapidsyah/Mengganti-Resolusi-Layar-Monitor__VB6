VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengganti Resolusi Layar Monitor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  'Ganti '800,600' dengan resolusi yang Anda inginkan.
  'Anda dapat mengganti "color pallete" ke 32 - Bit
  'dengan mengganti '16' di bawah dengan '32'
  'ChangeScreenSettings 640, 480, 16 - Bit  'Contoh
  'apabila menggunakan 640x480, 16 Bit
  ChangeScreenSettings 800, 600, 16 - Bit  'Contoh yang
  'diaplikasikan saat ini.
End Sub

Public Sub ChangeScreenSettings(lWidth As Integer, _
lHeight As Integer, lColors As Integer)

Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long
lIndex = 0
Do
lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
If lTemp = 0 Then Exit Do
lIndex = lIndex + 1
With tDevMode
If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight _
And .dmBitsPerPel = lColors Then
lTemp = ChangeDisplaySettings(tDevMode, _
CDS_UPDATEREGISTRY)
Exit Do
End If
End With
Loop
Select Case lTemp
Case DISP_CHANGE_SUCCESSFUL
     MsgBox "Setting tampilan baru telah berhasil", _
            vbInformation

Case DISP_CHANGE_RESTART
     MsgBox "Komputer harus di-restart agar mode grafik dapat berfungsi!", vbQuestion

Case DISP_CHANGE_FAILED
     MsgBox "Driver dari tampilan gagal memilih mode grafik!", vbCritical

Case DISP_CHANGE_BADMODE
     MsgBox "Mode grafik tidak mendukung!", vbCritical

Case DISP_CHANGE_NOTUPDATED
     MsgBox "Tidak dapat menulis setting ke dalam registry", vbCritical

Case DISP_CHANGE_BADFLAGS
     MsgBox "Anda memasukkan data yang tidak valid!", vbCritical

End Select
End Sub


