Attribute VB_Name = "ModuloReportes"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo para Reportes EXCEL
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long



Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" _
    Alias "capCreateCaptureWindowA" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hwndParent As Long, _
    ByVal nID As Long) As Long

Public Declare Function DestroyWindow Lib "USER32" (ByVal hndw As Long) As Boolean

Public Declare Function Imagen Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function video Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Public Auxiliar As Long
Public Const CONNECT As Long = 1034
Public Const DISCONNECT As Long = 1035
Public Const GET_FRAME As Long = 1084
Public Const COPY As Long = 1054

Public Const ws_child = &H40000000
Public Const ws_visible = &H10000000
Public Const WM_USER = 1024
Public Const wm_cap_driver_connect = WM_USER + 10
Public Const wm_cap_set_preview = WM_USER + 50
Public Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Public Const WM_CAP_DRIVER_DISCONNECT = WM_USER + 11
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_USER + 41
Public Const WM_CAP_DLG_VIDEOCONFIG = WM_USER + 42
Public Const WM_CAP_SET_SCALE = WM_USER + 53
Global Const WM_CAP_EDIT_COPY = WM_USER + 30
  




Public hwdc As Long
Public startcap As Integer
  



'para tener la ventana activa
Declare Function SetForegroundWindow Lib "USER32" (ByVal hwnd As Long) As Long

Dim mo_ExcelApplication As Excel.Application

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Const wxSinApellido As String = "__________"


Function GalenhosExcelApplication() As Excel.Application
    If mo_ExcelApplication Is Nothing Then
        Set mo_ExcelApplication = New Excel.Application
    End If
    Set GalenhosExcelApplication = mo_ExcelApplication
End Function
Function GalenhosKillExcelApplication()
    Set mo_ExcelApplication = Nothing
End Function


  

  

  















