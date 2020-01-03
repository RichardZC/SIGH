Attribute VB_Name = "ModuloReportes"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo para Reportes en excel
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Const WxLOTEpaquete As String = "PQTELOTE"
Public Const WxFVENCIMIENTOpaquete As String = "31/12/2020"
Public Const WxREGSANITARIOpaquete As String = "PQTE1234567890"


Dim mo_ExcelApplication As Excel.Application

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
'para tener la ventana activa
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Const SETTINGS_PROGID = "biopdf.PDFSettings"
'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function GetProfileString Lib "KERNEL32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long
Declare Function WriteProfileString Lib "KERNEL32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
    
Sub SeteaOtraImpresoraDefault(lcNuevaImpresora As String)
    If lcNuevaImpresora <> "" Then
        Dim Di As Long, L As Long, lcImpresora
        lcImpresora = lcNuevaImpresora & ",winspool,Ne05"
        Di = WriteProfileString("WINDOWS", "DEVICE", lcImpresora)
       ' L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    End If
End Sub

Function GalenhosExcelApplication() As Excel.Application
    If mo_ExcelApplication Is Nothing Then
        Set mo_ExcelApplication = New Excel.Application
    End If
    Set GalenhosExcelApplication = mo_ExcelApplication
End Function
Function GalenhosKillExcelApplication()
    Set mo_ExcelApplication = Nothing
End Function

Function SePuedeImprimirPDF(lcArchivoPDF As String, lbConVistaPrevia As Boolean) As Boolean
    SePuedeImprimirPDF = False
End Function

Function ImpresoraDefault() As String
    Dim buffer As String
    Dim ret As Integer
    buffer = Space(255)
    ret = GetProfileString("Windows", ByVal "device", "", _
                                 buffer, Len(buffer))
    If ret Then
        ImpresoraDefault = UCase(Left(buffer, _
                                   InStr(buffer, ",") - 1))
    End If
End Function


