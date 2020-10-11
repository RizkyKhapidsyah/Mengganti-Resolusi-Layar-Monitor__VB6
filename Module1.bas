Attribute VB_Name = "Module1"
'Program akan mengkonfirmasikan ke Anda jika komputer 'harus di-restart untuk menyimpan perubahan.

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const DISP_CHANGE_FAILED = -1
Public Const DISP_CHANGE_BADMODE = -2
Public Const DISP_CHANGE_NOTUPDATED = -3
Public Const DISP_CHANGE_BADFLAGS = -4
Public Const DISP_CHANGE_BADPARAM = -5
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H2
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Type DEVMODE
dmDeviceName As String * CCDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" _
Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As _
Long, ByVal iModeNum As Long, lpDevMode As Any) As _
Boolean

Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

