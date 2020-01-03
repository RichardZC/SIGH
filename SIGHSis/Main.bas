Attribute VB_Name = "ModuloInicial"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo para impresion en Excel
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ExcelApplication As Excel.Application

Public Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Type ImpresionDOS  ' Define el tipo definido por el usuario.
  codigo As String * 20
  NombreProducto As String * 40
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

'
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Declare Function WriteProfileString Lib "KERNEL32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
Const SETTINGS_PROGID = "biopdf.PDFSettings"

Function ValidaConfiguracionRegional()
  Dim sFormat  As String
  Dim sMsg As String
  Dim bHora As Boolean
  Dim bDecimal As Boolean
  Dim bMiles As Boolean
  Dim bFecha  As Boolean
  Dim bMilesMon As Boolean
  Dim bDecimalMon As Boolean

  ValidaConfiguracionRegional = False
  
  bHora = True
  bFecha = True
  bDecimal = True
  bMiles = True
  bDecimalMon = True
  bMilesMon = True

  'Obtiene la configuración regional de la fecha
  sFormat = LCase(sighentidades.FormatoFechaCorta)
  If sFormat <> sighentidades.DevuelveFechaSoloFormato_DMY Then
    bFecha = False
    sMsg = sMsg + "Formato de fecha dice: [" + sFormat + "] debe decir: [dd/MM/yyyy]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de fecha: [" + sFormat + "]" + Chr(13)
  End If

  sFormat = sighentidades.SeparadorDecimal
  If sFormat <> "." Then
    bDecimal = False
    sMsg = sMsg + "Formato de separador decimal dice: [" + sFormat + "] debe decir: [.]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador decimal: [" + sFormat + "]" + Chr(13)
  End If

  sFormat = sighentidades.SeparadorDeMiles
  If sFormat <> "," Then
    bMiles = False
    sMsg = sMsg + "Formato de separador miles dice: [" + sFormat + "] debe decir: [,]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador miles: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighentidades.SeparadorDecimalDeMonedas
  If sFormat <> "." Then
    bDecimalMon = False
    sMsg = sMsg + "Formato de separador decimal de monedas dice: [" + sFormat + "] debe decir: [.]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador decimal de monedas: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighentidades.SeparadorDeMilesDeMonedas
  If sFormat <> "," Then
    bMilesMon = False
    sMsg = sMsg + "Formato de separador de miles de monedas dice: [" + sFormat + "] debe decir: [,]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador de miles de monedas: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighentidades.FormatoDeHoras
  If sFormat <> "hh:mm:ss tt" Then
    bHora = False
    sMsg = sMsg + "Formato de horas dice: [" + sFormat + "] debe decir: [hh:mm:ss tt]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de horas: [" + sFormat + "]" + Chr(13)
  End If
    
  Dim iResp As Integer
  Dim lngX As Long
  Dim iX As Integer
    
  If Not (bHora And bDecimal And bMiles And bFecha And bDecimalMon And bMilesMon) Then
    iResp = MsgBox("Algunos de los valores de la configuración regional " + Chr(13) + _
            " no coincide con los valores requeridos por el sistema " + Chr(13) + _
            sMsg + Chr(13) + _
            "Estos valores se modificaran a continuación, si tiene otra aplicación que use un formato diferente recuerde que puede modificar estos valores en el panel de control >> Configuración Regional." + Chr(13) + _
            "y vuelva a reingresar al sistema.", vbInformation, "Configuración del Sistema")
                    
            sighentidades.FormatoFechaCorta = sighentidades.DevuelveFechaSoloFormato_DMY
            sighentidades.SeparadorDecimal = "."
            sighentidades.SeparadorDeMiles = ","
            sighentidades.SeparadorDecimalDeMonedas = "."
            sighentidades.SeparadorDeMilesDeMonedas = ","
            sighentidades.FormatoDeHoras = "hh:mm:ss tt"
                    
  End If
  ValidaConfiguracionRegional = True
End Function

Function GalenhosExcelApplication() As Excel.Application
  If mo_ExcelApplication Is Nothing Then Set mo_ExcelApplication = New Excel.Application
  Set GalenhosExcelApplication = mo_ExcelApplication
End Function

Function GalenhosKillExcelApplication()
  Set mo_ExcelApplication = Nothing
End Function
Public Sub SeleccionaMask(ByVal MaskB As MaskEdBox)
  MaskB.SelStart = 0
  MaskB.SelLength = Len(MaskB.Text)
End Sub


Function SePuedeImprimirPDF(lcArchivoPDF As String, lbConVistaPrevia As Boolean) As Boolean
    SePuedeImprimirPDF = False
End Function


Sub SeteaOtraImpresoraDefault(lcNuevaImpresora As String)
    If lcNuevaImpresora <> "" Then
        Dim Di As Long, L As Long, lcImpresora
        lcImpresora = lcNuevaImpresora & ",winspool,Ne05"
        Di = WriteProfileString("WINDOWS", "DEVICE", lcImpresora)
       ' L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    End If
End Sub
