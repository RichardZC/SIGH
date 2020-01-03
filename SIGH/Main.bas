Attribute VB_Name = "ModuloInicial"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Modulo para Reportes
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------

Option Explicit

Dim mo_ExcelApplication As Excel.Application

Public Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Type ImpresionDOS  ' Define el tipo definido por el usuario.
  Codigo As String * 20
  NombreProducto As String * 40
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'para tener la ventana activa
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'

Public Sub Main()
  ' we need to call InitCommonControls before we
  ' can use XP visual styles.  Here I'm using
  ' InitCommonControlsEx, which is the extended
  ' version provided in v4.72 upwards (you need
  ' v6.00 or higher to get XP styles)
  'On Error Resume Next
  ' this will fail if Comctl not available
  '  - unlikely now though!
  Dim iccex As tagInitCommonControlsEx
  With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_USEREX_CLASSES
  End With
  InitCommonControlsEx iccex
   
  If Not ValidaConfiguracionRegional() Then End
   
  '*******************************

    Dim sTiempo As String
    sTiempo = Now
    Splash.Show 'vbModal
    Splash.Refresh
    Dim oLoginForm As New Login
    oLoginForm.Show vbModal
    
    If Not oLoginForm.Autenticado Then End
  
 
  

    Dim oFormPrincipal As New Principal
    Set Principal.LoginForm = oLoginForm
    Unload Splash
    Principal.Show
  
    
  
    
  
  '*******************************
End Sub

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
  sFormat = LCase(sighEntidades.FormatoFechaCorta)
  If sFormat <> sighEntidades.DevuelveFechaSoloFormato_DMY Then
    bFecha = False
    sMsg = sMsg + "Formato de fecha dice: [" + sFormat + "] debe decir: [dd/MM/yyyy]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de fecha: [" + sFormat + "]" + Chr(13)
  End If

  sFormat = sighEntidades.SeparadorDecimal
  If sFormat <> "." Then
    bDecimal = False
    sMsg = sMsg + "Formato de separador decimal dice: [" + sFormat + "] debe decir: [.]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador decimal: [" + sFormat + "]" + Chr(13)
  End If

  sFormat = sighEntidades.SeparadorDeMiles
  If sFormat <> "," Then
    bMiles = False
    sMsg = sMsg + "Formato de separador miles dice: [" + sFormat + "] debe decir: [,]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador miles: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighEntidades.SeparadorDecimalDeMonedas
  If sFormat <> "." Then
    bDecimalMon = False
    sMsg = sMsg + "Formato de separador decimal de monedas dice: [" + sFormat + "] debe decir: [.]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador decimal de monedas: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighEntidades.SeparadorDeMilesDeMonedas
  If sFormat <> "," Then
    bMilesMon = False
    sMsg = sMsg + "Formato de separador de miles de monedas dice: [" + sFormat + "] debe decir: [,]" + Chr(13)
  Else
    sMsg = sMsg + "Formato de separador de miles de monedas: [" + sFormat + "]" + Chr(13)
  End If
    
  sFormat = sighEntidades.FormatoDeHoras
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
                    
            sighEntidades.FormatoFechaCorta = sighEntidades.DevuelveFechaSoloFormato_DMY
            sighEntidades.SeparadorDecimal = "."
            sighEntidades.SeparadorDeMiles = ","
            sighEntidades.SeparadorDecimalDeMonedas = "."
            sighEntidades.SeparadorDeMilesDeMonedas = ","
            sighEntidades.FormatoDeHoras = "hh:mm:ss tt"
  End If
  ValidaConfiguracionRegional = True
End Function

Function GalenhosExcelApplication() As Excel.Application
    If mo_ExcelApplication Is Nothing Then
        Set mo_ExcelApplication = New Excel.Application
    End If
    Set GalenhosExcelApplication = mo_ExcelApplication
End Function

Function GalenhosKillExcelApplication()
  Set mo_ExcelApplication = Nothing
End Function

Public Sub SeleccionaMask(ByVal MaskB As MaskEdBox)
  MaskB.SelStart = 0
  MaskB.SelLength = Len(MaskB.Text)
End Sub


Sub PVcomboBoxUbicaPosicion(lcCodigo As String, cmbComboPV As PVComboBox)
    Dim lnFor As Integer
    If lcCodigo <> "" Then
        For lnFor = 0 To (cmbComboPV.ListCount - 1)
            cmbComboPV.ListIndex = lnFor
            If cmbComboPV.SubItem(cmbComboPV.ListIndex, 0) = lcCodigo Then
               Exit For
            End If
        Next
    Else
        cmbComboPV.ListIndex = -1
    End If
End Sub

Function PVcomboBoxDevuelveEleccion(cmbComboPV As PVComboBox) As String
           Dim oCampos() As String
           If cmbComboPV.ListIndex < 0 Then
               PVcomboBoxDevuelveEleccion = ""
           Else
               oCampos = Split(cmbComboPV.List(cmbComboPV.ListIndex), "|")
               PVcomboBoxDevuelveEleccion = oCampos(0)
           End If
End Function
