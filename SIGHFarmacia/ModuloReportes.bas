Attribute VB_Name = "ModuloReportes"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Modulo para Reportes
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

'para tener la ventana activa
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Dim mo_ExcelApplication As Excel.Application

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Const WxLOTEpaquete As String = "PQTELOTE"
Public Const WxFVENCIMIENTOpaquete As String = "31/12/2020"
Public Const WxREGSANITARIOpaquete As String = "PQTE1234567890"



Function GalenhosExcelApplication() As Excel.Application
    If mo_ExcelApplication Is Nothing Then
        Set mo_ExcelApplication = New Excel.Application
    End If
    Set GalenhosExcelApplication = mo_ExcelApplication
End Function
Function GalenhosKillExcelApplication()
    Set mo_ExcelApplication = Nothing
End Function

