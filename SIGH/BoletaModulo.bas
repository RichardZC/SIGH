Attribute VB_Name = "BoletaModulo"
'mgaray
'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function GetProfileString Lib "KERNEL32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long

    
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Login
Public WxLcVersionSisGalenPlus As String
Public Const WxSkin11 As String = "winaqua.skn"

Public WxLnCabRucX As Long, WxLnCabRucY As Long, WxLnCabRucX_F As Long, WxLnCabRucY_F  As Long
Public WxLnCabDireccionX As Long, WxLnCabDireccionY As Long, WxLnCabDireccionX_F As Long, WxLnCabDireccionY_F As Long


'servicios
Public WxLnDireccionEESSX As Long, WxLnDireccionEESSY As Long
Public WxLnRucEESSX As Long, WxLnRucEESSY As Long
Public WxLnTelefonoEESSX As Long, WxLnTelefonoEESSY As Long
Public WxLnRzSocialPacienteX As Long, WxLnRzSocialPacienteY As Long

Public WxLnNumeroSerieX As Long, WxLnNumeroSerieY As Long
Public WxLnEstadoX As Long, WxLnEstadoY As Long
Public WxLnTipoX As Long, WxLnTipoY As Long
Public WxLnRzSocialX As Long, WxLnRzSocialY As Long
Public WxLnFechaX As Long, WxLnFechaY As Long
Public WxLnServicioX As Long, WxLnServicioY As Long
Public WxLnObservacionesX As Long, WxLnObservacionesY As Long
'nuevo
Public WxLnPaqueteX As Long, WxLnPaqueteY As Long
Public WxLnFarmaceuticoX As Long, WxLnFarmaceuticoY As Long
Public WxLnDNIX As Long, WxLnDNIY As Long
'-------
Public WxLnHistoriaX As Long, WxLnHistoriaY As Long 'Ticketera
Public WxLnNItemY As Long, WxLnCodigoY As Long, WxLnProductoY As Long, WxLnCantidadY As Long, WxLnPrecioY As Long, WxLnImporteY As Long
Public WxLnCajeroX As Long, WxLnCajeroY As Long
Public WxLnCajaX As Long, WxLnCajaY As Long
Public WxLnAdelantosX As Long, WxLnAdelantosY As Long
Public WxLnTotalPagarX As Long, WxLnTotalPagarY As Long
Public WxLnTotalItemsX As Long, WxLnTotalItemsY As Long
Public WxLnCuentaX As Long, WxLnCuentaY As Long
Public WxLnFUAX As Long, WxLnFUAY As Long
Public WxLnExoneracionesX As Long, WxLnExoneracionesY As Long
Public WxLnTotalEnLetrasX As Long, WxLnTotalEnLetrasY As Long
Public WxLnTotalX As Long, WxLnTotalY As Long
Public WxLnSubTotalX As Long, WxLnSubTotalY As Long 'Factura
Public WxLnIGVX As Long, WxLnIGVY As Long 'Factura
Public WxLnCabeceraAlto As Long, WxLnPieAlto As Long
Public WxLnTerminalX As String, WxLnTerminalY As String
Public WxLnSerieImpresoraX As String, WxLnSerieImpresoraY As String
Public WxLnNombrePaqueteX As Long, WxLnNombrePaqueteY As Long
Public WxLnDniPacienteX As Long, WxLnDniPacienteY As Long
Public WxLnUsuarioDespachoX As Long, WxLnUsuarioDespachoY As Long

'farmacia
Public WxLnDireccionEESSX_F As Long, WxLnDireccionEESSY_F As Long
Public WxLnRucEESSX_F As Long, WxLnRucEESSY_F As Long
Public WxLnTelefonoEESSX_F As Long, WxLnTelefonoEESSY_F As Long
Public WxLnRzSocialPacienteX_F As Long, WxLnRzSocialPacienteY_F As Long
Public WxLnModeloImpresoraX_F  As String, WxLnModeloImpresoraY_F  As String
Public WxLnSerieImpresoraX_F  As String, WxLnSerieImpresoraY_F  As String

Public WxLnNumeroSerieX_F As Long, WxLnNumeroSerieY_F As Long
Public WxLnEstadoX_F As Long, WxLnEstadoY_F As Long
Public WxLnTipoX_F As Long, WxLnTipoY_F As Long
Public WxLnRzSocialX_F As Long, WxLnRzSocialY_F As Long
Public WxLnFechaX_F As Long, WxLnFechaY_F As Long
Public WxLnServicioX_F As Long, WxLnServicioY_F As Long
Public WxLnObservacionesX_F As Long, WxLnObservacionesY_F As Long
'nuevo
Public WxLnPaqueteX_F As Long, WxLnPaqueteY_F As Long
Public WxLnFarmaceuticoX_F As Long, WxLnFarmaceuticoY_F As Long
Public WxLnDNIX_F As Long, WxLnDNIY_F As Long
'---
Public WxLnHistoriaX_F As Long, WxLnHistoriaY_F As Long 'Ticketera
Public WxLnCodigoY_F As Long, WxLnProductoY_F As Long, WxLnCantidadY_F As Long, WxLnPrecioY_F As Long, WxLnImporteY_F As Long
Public WxLnCajeroX_F As Long, WxLnCajeroY_F As Long
Public WxLnCajaX_F As Long, WxLnCajaY_F As Long
Public WxLnAdelantosX_F As Long, WxLnAdelantosY_F As Long
Public WxLnTotalPagarX_F As Long, WxLnTotalPagarY_F As Long
Public WxLnCuentaX_F As Long, WxLnCuentaY_F As Long
Public WxLnExoneracionesX_F As Long, WxLnExoneracionesY_F As Long
Public WxLnTotalEnLetrasX_F As Long, WxLnTotalEnLetrasY_F As Long
Public WxLnTotalX_F As Long, WxLnTotalY_F As Long
Public WxLnSubTotalX_F As Long, WxLnSubTotalY_F As Long 'Factura
Public WxLnIGVX_F As Long, WxLnIGVY_F As Long 'Factura
Public WxLnCabeceraAlto_F As Long, WxLnPieAlto_F As Long
Public WxLnTerminalX_F As String, WxLnTerminalY_F As String
'JR 10052016
Public WxLnNombrePaqueteX_F As Long, WxLnNombrePaqueteY_F As Long
Public WxLnDniPacienteX_F As Long, WxLnDniPacienteY_F As Long
Public WxLnUsuarioDespachoX_F As Long, WxLnUsuarioDespachoY_F As Long
'caja
Public wxParametro102 As String, wxParametro208 As String, wxParametro285 As String
Public wxParametro286 As String, wxParametro288 As String, wxParametro527 As String
Public wxParametro211 As String, wxParametro237 As String, wxParametro221 As String
Public wxIdTipoComprobanteDefault As Long, wxParametro532 As String, wxParametro533 As String
Public wxIdTipoComprobante2 As Long, lcParametro523 As String, lcParametro524 As String
Public WxLnProductoWidhtY_F As Long, WxLnProductoWidhtY As Long
Public WxLnTotalLetrasWidhtY_F As Long, WxLnTotalLetrasWidhtY As Long
Public wxParametro346 As String, wxParametro379 As String, wxParametro377 As String   'SUNAT
Public wxParametro500 As String, wxParametro501 As String '18/05/2016
Public wxParametro538 As String, wxParametro7 As String, wxParametro543 As String
Public wxParametro548 As String, wxParametro549 As String, wxParametro534 As String
Public wxParametro557 As String, wxParametro558 As String
Public wxHuboCambioAfactura As Boolean
'ce
Public Const Lx_Lab As String = "LAB", LxDx As String = "DX", LxCPT As String = "CPT"
Public wxParametro258      As String, wxParametro274     As String, wxParametro275     As String
Public wxParametro276     As String, wxParametro281     As String, wxParametro282     As String
Public wxParametro296     As String, wxParametro302     As String, wxParametroJAMO     As String
Public wxParametro312 As String, wxParametro289    As String, wxParametro329 As String
Public wxParametro306 As String, wxParametro216 As String, wxParametro518 As String
Public wxParametro287 As String, wxParametro333 As String, wxParametro336 As String
Public wxParametro502 As String, wxParametro514 As String, wxParametro506 As String
Public wxParametro511 As String, wxParametro513 As String, wxParametro512 As String
Public wxParametro517 As String, wxParametro522 As String, wxParametro539 As String
Public wxParametro540 As String, wxParametro541 As String, wxParametro542 As String
Public wxParametro555 As String
'hosp/emerg
Public wxParametro202 As String, wxParametro203 As String, wxParametro204 As String
Public wxParametro210 As String, wxParametro212 As String, wxParametro215 As String
Public wxParametro231         As String, wxParametro232 As String, wxParametro525 As String
Public wxParametro233 As String, wxParametro259 As String, wxParametro353 As String
Public wxParametro290 As String, wxParametro291 As String, wxParametro292 As String
Public wxParametro357 As String, wxParametro526 As String, wxParametro546 As String, wxParametro547 As String

'solo Emergencia
Public wxParametro316 As String, wxParametro317 As String, wxParametro521 As String
Public wxParametro530 As String, wxParametro536 As String, wxParametro552 As String
Public wxParametro559 As String
'Fua
Public wxParametro205 As String, wxParametro206 As String, wxParametro207 As String, wxParametro242 As String
Public wxParametro280 As String, wxParametro303 As String, wxParametro304 As String, wxParametro305 As String
Public wxParametro310 As String, wxParametro320 As String, wxParametro322 As String, wxParametro323 As String, wxParametro324 As String
Public wxParametro326 As String, wxParametro301 As String, wxParametro553 As String
Public wxParametro327 As String, wxParametro328 As String, wxParametro358 As String, wxParametro359 As String, wxParametro362 As String
Public wxParametroSIS As String, wxParametro339 As String, wxParametro380 As String, wxParametro381 As String, wxParametro386 As String
Public wxParametro387 As String, wxParametro389 As String, wxParametro382 As String, wxParametro395 As String, wxParametro396 As String
'Ce/hosp/Emeg
Public Const wxSinApellido As String = "__________"
Public WxDEFAULT_BUSQ_PACIENTE As Integer, WxDEFAULT_BUSQ_CE As Integer
Public WxDEFAULT_BUSQ_EMERGENCIA As Integer, WxDEFAULT_BUSQ_HOSPITALIZ As Integer
Public wxParametro351 As String, wxParametro354 As String, wxParametro545 As String, wxParametro554 As String
Public wxparametro563 As String, wxparametro564 As String, wxparametro565 As String, wxparametro566 As String
'GLCC 02/11/20 CAMBIO36 INICIO
'YA NO GUARDA EL NUMERO 9
'Public Const wxNueve = "9"
'GLCC 02/11/20 CAMBIO36 FIN
'mgaray variables de pagina
Public WxLnNombreHoja As String
Public WxLnTipoReporteador As Integer
Public WxLnMargenIzquierdoX As Long, WxLnMargenDerechoX As Long, WxLnMargenSuperiorY As Long, WxLnMargenInferiorY As Long

Public WxLnNombreHoja_F As String
Public WxLnTipoReporteador_F As Integer
Public WxLnMargenIzquierdoX_F As Long, WxLnMargenDerechoX_F As Long, WxLnMargenSuperiorY_F As Long, WxLnMargenInferiorY_F As Long

'Fcachay variables Boleta/Ticket Farmacia - Ventas
Public WxLnNombreEESSX As Long, WxLnNombreEESSY As Long
Public WxLnTipoFormatoX As Long, WxLnTipoFormatoY As Long
Public WxLnFarmaciaX As Long, WxLnFarmaciaY As Long
Public WxLnPacienteX As Long, WxLnPacienteY As Long
Public WxLnDiagPrincipalX As Long, WxLnDiagPrincipalY As Long
Public WxLnNroCuentaX As Long, WxLnNroCuentaY As Long
Public WxLnServicioHospX As Long, WxLnServicioHospY As Long
Public WxLnNroMovmientoX As Long, WxLnNroMovmientoY As Long
Public WxLnFechaMovimientoX As Long, WxLnFechaMovimientoY As Long

Declare Function WriteProfileString Lib "KERNEL32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long

Public Const wxHoraMadrugada As String = "14:55"
Const SETTINGS_PROGID = "biopdf.PDFSettings"

Sub SeteaOtraImpresoraDefault(lcNuevaImpresora As String)
    If lcNuevaImpresora <> "" Then
        Dim Di As Long, L As Long, lcImpresora
        lcImpresora = lcNuevaImpresora & ",winspool,Ne05"
        Di = WriteProfileString("WINDOWS", "DEVICE", lcImpresora)
        L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    End If
End Sub

Function SeteaOtraImpresoraDiferenteAlDefault(Nombre_Impresora As String)
    If Nombre_Impresora <> "" Then
        Dim Prt As Printer
        ' Establece la impresora que se utilizará para imprimir
        For Each Prt In Printers
        'LO UNICO QUE GUARDAMOS EN UN ARCHIVO ES EL NOMBRE
        'Y LO PONEMOS EN LA PROPIEDAD DEVICENAME DEL PRINTER
            If Prt.DeviceName = Nombre_Impresora Then
                Set Printer = Prt
            End If
        Next
    End If
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

Function SePuedeImprimirPDF(lcArchivoPDF As String, lbConVistaPrevia As Boolean) As Boolean
    SePuedeImprimirPDF = False
'    On Error GoTo ErrPDF
'    Dim sPrinterName As String, sImpresoraDefault As String
'    Dim settings As Object
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    sPrinterName = Trim(lcBuscaParametro.SeleccionaFilaParametro(572))
'    Set lcBuscaParametro = Nothing
'    SeteaOtraImpresoraDefault sPrinterName
'    Set settings = CreateObject(SETTINGS_PROGID)
'    settings.PrinterName = sPrinterName
'    settings.SetValue "Output", lcArchivoPDF
'    settings.SetValue "ConfirmOverwrite", "no"
'    settings.SetValue "ShowSaveAS", "never"
'    settings.SetValue "ShowSettings", "never"
'    If lbConVistaPrevia = True Then
'       settings.SetValue "ShowPDF", "yes"
'    Else
'       settings.SetValue "ShowPDF", "no"
'    End If
'    settings.SetValue "RememberLastFileName", "no"
'    settings.SetValue "RememberLastFolderName", "no"
'    settings.WriteSettings True
'    SePuedeImprimirPDF = True
'ErrPDF:
End Function



Function sGetINI(sIniFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sIniFile)
    sGetINI = Left$(sTemp, nLength)
End Function
'Carga archivo Setup_caja.ini (fin)


'sunat
Sub CargaSetup_Caja(lcRutaINI As String, lnIdTipoComprobanteDefault As Long, lbBoletaFacturaFormatoTicket As Boolean)
 On Error GoTo ErrINI
' If lnIdTipoComprobanteDefault = 3 And wxParametro527 = "S" Then 'es una BOLETA pero se imprimirá como ticket
'    lnIdTipoComprobanteDefault = 4
' End If
 Dim sIniFile As String
 
' Select Case lnIdTipoComprobanteDefault
' Case 3      '"Boleta"
'    sIniFile = lcRutaINI & "\setup_caja_boleta.ini"
' Case 2      '"Factura"
'     sIniFile = lcRutaINI & "\setup_caja_factura.ini"
' Case 1      ' "Recibo"
'     sIniFile = lcRutaINI & "\setup_caja_recibo.ini"
' Case 4       '"Ticket"
'     sIniFile = lcRutaINI & "\setup_caja_ticket.ini"
' End Select

sIniFile = getPathIniFile(lcRutaINI, lnIdTipoComprobanteDefault, lbBoletaFacturaFormatoTicket)

 
' sIniFile = lcRutaINI & "\setup_caja.ini"
 If Dir$(sIniFile) <> "" Then
    'Servicios
    WxLnDireccionEESSX = Val(sGetINI(sIniFile, "Variables", "DireccionEESSX", "?"))
    WxLnDireccionEESSY = Val(sGetINI(sIniFile, "Variables", "DireccionEESSY", "?"))
    WxLnRucEESSX = Val(sGetINI(sIniFile, "Variables", "RucEESSX", "?"))
    WxLnRucEESSY = Val(sGetINI(sIniFile, "Variables", "RucEESSY", "?"))
    WxLnTelefonoEESSX = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSX", "?"))
    WxLnTelefonoEESSY = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSY", "?"))
    WxLnRzSocialPacienteX = Val(sGetINI(sIniFile, "Variables", "RzSocialPacienteX", "?"))
    WxLnRzSocialPacienteY = Val(sGetINI(sIniFile, "Variables", "RzSocialPacienteY", "?"))
    
    WxLnNumeroSerieX = Val(sGetINI(sIniFile, "Variables", "NumeroSerieX", "?"))
    WxLnNumeroSerieY = Val(sGetINI(sIniFile, "Variables", "NumeroSerieY", "?"))
    WxLnEstadoX = Val(sGetINI(sIniFile, "Variables", "EstadoX", "?"))
    WxLnEstadoY = Val(sGetINI(sIniFile, "Variables", "EstadoY", "?"))
    WxLnTipoX = Val(sGetINI(sIniFile, "Variables", "TipoX", "?"))
    WxLnTipoY = Val(sGetINI(sIniFile, "Variables", "TipoY", "?"))
    WxLnRzSocialX = Val(sGetINI(sIniFile, "Variables", "RzSocialX", "?"))
    WxLnRzSocialY = Val(sGetINI(sIniFile, "Variables", "RzSocialY", "?"))
    WxLnFechaX = Val(sGetINI(sIniFile, "Variables", "FechaX", "?"))
    WxLnFechaY = Val(sGetINI(sIniFile, "Variables", "FechaY", "?"))
    WxLnServicioX = Val(sGetINI(sIniFile, "Variables", "ServicioX", "?"))
    WxLnServicioY = Val(sGetINI(sIniFile, "Variables", "ServicioY", "?"))
    WxLnObservacionesX = Val(sGetINI(sIniFile, "Variables", "ObservacionesX", "?"))
    WxLnObservacionesY = Val(sGetINI(sIniFile, "Variables", "ObservacionesY", "?"))
    WxLnHistoriaX = Val(sGetINI(sIniFile, "Variables", "HistoriaX", "?"))
    WxLnHistoriaY = Val(sGetINI(sIniFile, "Variables", "HistoriaY", "?"))
    '
    WxLnCodigoY = Val(sGetINI(sIniFile, "Variables", "CodigoY", "?"))
    WxLnProductoY = Val(sGetINI(sIniFile, "Variables", "ProductoY", "?"))
    WxLnProductoWidhtY = Val(sGetINI(sIniFile, "Variables", "ProductoWidhtY", "?"))
    WxLnCantidadY = Val(sGetINI(sIniFile, "Variables", "CantidadY", "?"))
    WxLnPrecioY = Val(sGetINI(sIniFile, "Variables", "PrecioY", "?"))
    WxLnImporteY = Val(sGetINI(sIniFile, "Variables", "ImporteY", "?"))
    '
    WxLnCajeroX = Val(sGetINI(sIniFile, "Variables", "CajeroX", "?"))
    WxLnCajeroY = Val(sGetINI(sIniFile, "Variables", "CajeroY", "?"))
    WxLnCajaX = Val(sGetINI(sIniFile, "Variables", "CajaX", "?"))
    WxLnCajaY = Val(sGetINI(sIniFile, "Variables", "CajaY", "?"))
    WxLnAdelantosX = Val(sGetINI(sIniFile, "Variables", "AdelantosX", "?"))
    WxLnAdelantosY = Val(sGetINI(sIniFile, "Variables", "AdelantosY", "?"))
    WxLnTotalPagarX = Val(sGetINI(sIniFile, "Variables", "TotalPagarX", "?"))
    WxLnTotalPagarY = Val(sGetINI(sIniFile, "Variables", "TotalPagarY", "?"))
    WxLnCuentaX = Val(sGetINI(sIniFile, "Variables", "CuentaX", "?"))
    WxLnCuentaY = Val(sGetINI(sIniFile, "Variables", "CuentaY", "?"))
    WxLnExoneracionesX = Val(sGetINI(sIniFile, "Variables", "ExoneracionesX", "?"))
    WxLnExoneracionesY = Val(sGetINI(sIniFile, "Variables", "ExoneracionesY", "?"))
    WxLnTotalEnLetrasX = Val(sGetINI(sIniFile, "Variables", "TotalEnLetrasX", "?"))
    WxLnTotalEnLetrasY = Val(sGetINI(sIniFile, "Variables", "TotalEnLetrasY", "?"))
    WxLnTotalLetrasWidhtY = Val(sGetINI(sIniFile, "Variables", "TotalLetrasWidhtY", "?"))
    WxLnTotalX = Val(sGetINI(sIniFile, "Variables", "TotalX", "?"))
    WxLnTotalY = Val(sGetINI(sIniFile, "Variables", "TotalY", "?"))
    WxLnSubTotalX = Val(sGetINI(sIniFile, "Variables", "SubTotalX", "?"))
    WxLnSubTotalY = Val(sGetINI(sIniFile, "Variables", "SubTotalY", "?"))
    WxLnIGVX = Val(sGetINI(sIniFile, "Variables", "IGVX", "?"))
    WxLnIGVY = Val(sGetINI(sIniFile, "Variables", "IGVY", "?"))
    WxLnTerminalX = Val(sGetINI(sIniFile, "Variables", "TerminalX", "?"))
    WxLnTerminalY = Val(sGetINI(sIniFile, "Variables", "TerminalY", "?"))
    WxLnSerieImpresoraX = Val(sGetINI(sIniFile, "Variables", "SerieImpresoraX", "?"))
    WxLnSerieImpresoraY = Val(sGetINI(sIniFile, "Variables", "SerieImpresoraY", "?"))
    
    'JR 04042016
    WxLnNombrePaqueteX = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteX", "?"))
    WxLnNombrePaqueteY = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteY", "?"))
    WxLnDniPacienteX = Val(sGetINI(sIniFile, "Variables", "DniPacienteX", "?"))
    WxLnDniPacienteY = Val(sGetINI(sIniFile, "Variables", "DniPacienteY", "?"))
    WxLnUsuarioDespachoX = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoX", "?"))
    WxLnUsuarioDespachoY = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoY", "?"))
    
     'JR 10052016 (6L)
    WxLnNombrePaqueteX_F = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteX_F", "?"))
    WxLnNombrePaqueteY_F = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteY_F", "?"))
    WxLnDniPacienteX_F = Val(sGetINI(sIniFile, "Variables", "DniPacienteX_F", "?"))
    WxLnDniPacienteY_F = Val(sGetINI(sIniFile, "Variables", "DniPacienteY_F", "?"))
    WxLnUsuarioDespachoX_F = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoX_F", "?"))
    WxLnUsuarioDespachoY_F = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoY_F", "?"))
    
'    WxLnTextoVacioX = Val(sGetINI(sIniFile, "Variables", "TextoVacioX", "?"))
'    WxLnTextoVacioY = Val(sGetINI(sIniFile, "Variables", "TextoVacioY", "?"))
    '
    WxLnCabeceraAlto = Val(sGetINI(sIniFile, "Variables", "CabeceraAlto", "?"))
    WxLnPieAlto = Val(sGetINI(sIniFile, "Variables", "PieAlto", "?"))
    
   'mgaray
    WxLnNombreHoja = sGetINI(sIniFile, "Variables", "nombreHoja", "")
    WxLnTipoReporteador = Val(sGetINI(sIniFile, "Variables", "tipoReporteador", "0"))
    WxLnMargenIzquierdoX = Val(sGetINI(sIniFile, "Variables", "margenIzquierdoX", "100"))
    WxLnMargenDerechoX = Val(sGetINI(sIniFile, "Variables", "margenDerechoX", "100"))
    WxLnMargenSuperiorY = Val(sGetINI(sIniFile, "Variables", "margenSuperiorY", "100"))
    WxLnMargenInferiorY = Val(sGetINI(sIniFile, "Variables", "margenInferiorY", "100"))
    
    
    'Farmacia
    WxLnDireccionEESSX_F = Val(sGetINI(sIniFile, "Variables", "DireccionEESSX_F", "?"))
    WxLnDireccionEESSY_F = Val(sGetINI(sIniFile, "Variables", "DireccionEESSY_F", "?"))
    WxLnRucEESSX_F = Val(sGetINI(sIniFile, "Variables", "RucEESSX_F", "?"))
    WxLnRucEESSY_F = Val(sGetINI(sIniFile, "Variables", "RucEESSY_F", "?"))
    WxLnTelefonoEESSX_F = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSX_F", "?"))
    WxLnTelefonoEESSY_F = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSY_F", "?"))
    WxLnRzSocialPacienteX_F = Val(sGetINI(sIniFile, "Variables", "RzSocialPacienteX_F ", "?"))
    WxLnRzSocialPacienteY_F = Val(sGetINI(sIniFile, "Variables", "RzSocialPacienteY_F ", "?"))
    
    WxLnNumeroSerieX_F = Val(sGetINI(sIniFile, "Variables", "NumeroSerieX_F", "?"))
    WxLnNumeroSerieY_F = Val(sGetINI(sIniFile, "Variables", "NumeroSerieY_F", "?"))
    WxLnEstadoX_F = Val(sGetINI(sIniFile, "Variables", "EstadoX_F", "?"))
    WxLnEstadoY_F = Val(sGetINI(sIniFile, "Variables", "EstadoY_F", "?"))
    WxLnTipoX_F = Val(sGetINI(sIniFile, "Variables", "TipoX_F", "?"))
    WxLnTipoY_F = Val(sGetINI(sIniFile, "Variables", "TipoY_F", "?"))
    WxLnRzSocialX_F = Val(sGetINI(sIniFile, "Variables", "RzSocialX_F", "?"))
    WxLnRzSocialY_F = Val(sGetINI(sIniFile, "Variables", "RzSocialY_F", "?"))
    WxLnFechaX_F = Val(sGetINI(sIniFile, "Variables", "FechaX_F", "?"))
    WxLnFechaY_F = Val(sGetINI(sIniFile, "Variables", "FechaY_F", "?"))
    WxLnServicioX_F = Val(sGetINI(sIniFile, "Variables", "ServicioX_F", "?"))
    WxLnServicioY_F = Val(sGetINI(sIniFile, "Variables", "ServicioY_F", "?"))
    WxLnObservacionesX_F = Val(sGetINI(sIniFile, "Variables", "ObservacionesX_F", "?"))
    WxLnObservacionesY_F = Val(sGetINI(sIniFile, "Variables", "ObservacionesY_F", "?"))
    WxLnHistoriaX_F = Val(sGetINI(sIniFile, "Variables", "HistoriaX_F", "?"))
    WxLnHistoriaY_F = Val(sGetINI(sIniFile, "Variables", "HistoriaY_F", "?"))
    '
    WxLnCodigoY_F = Val(sGetINI(sIniFile, "Variables", "CodigoY_F", "?"))
    WxLnProductoY_F = Val(sGetINI(sIniFile, "Variables", "ProductoY_F", "?"))
    WxLnProductoWidhtY_F = Val(sGetINI(sIniFile, "Variables", "ProductoWidhtY_F", "?"))
    WxLnCantidadY_F = Val(sGetINI(sIniFile, "Variables", "CantidadY_F", "?"))
    WxLnPrecioY_F = Val(sGetINI(sIniFile, "Variables", "PrecioY_F", "?"))
    WxLnImporteY_F = Val(sGetINI(sIniFile, "Variables", "ImporteY_F", "?"))
    '
    WxLnCajeroX_F = Val(sGetINI(sIniFile, "Variables", "CajeroX_F", "?"))
    WxLnCajeroY_F = Val(sGetINI(sIniFile, "Variables", "CajeroY_F", "?"))
    WxLnCajaX_F = Val(sGetINI(sIniFile, "Variables", "CajaX_F", "?"))
    WxLnCajaY_F = Val(sGetINI(sIniFile, "Variables", "CajaY_F", "?"))
    WxLnAdelantosX_F = Val(sGetINI(sIniFile, "Variables", "AdelantosX_F", "?"))
    WxLnAdelantosY_F = Val(sGetINI(sIniFile, "Variables", "AdelantosY_F", "?"))
    WxLnTotalPagarX_F = Val(sGetINI(sIniFile, "Variables", "TotalPagarX_F", "?"))
    WxLnTotalPagarY_F = Val(sGetINI(sIniFile, "Variables", "TotalPagarY_F", "?"))
    WxLnCuentaX_F = Val(sGetINI(sIniFile, "Variables", "CuentaX_F", "?"))
    WxLnCuentaY_F = Val(sGetINI(sIniFile, "Variables", "CuentaY_F", "?"))
    WxLnExoneracionesX_F = Val(sGetINI(sIniFile, "Variables", "ExoneracionesX_F", "?"))
    WxLnExoneracionesY_F = Val(sGetINI(sIniFile, "Variables", "ExoneracionesY_F", "?"))
    WxLnTotalEnLetrasX_F = Val(sGetINI(sIniFile, "Variables", "TotalEnLetrasX_F", "?"))
    WxLnTotalEnLetrasY_F = Val(sGetINI(sIniFile, "Variables", "TotalEnLetrasY_F", "?"))
    WxLnTotalLetrasWidhtY_F = Val(sGetINI(sIniFile, "Variables", "TotalLetrasWidhtY_F", "?"))
    WxLnTotalX_F = Val(sGetINI(sIniFile, "Variables", "TotalX_F", "?"))
    WxLnTotalY_F = Val(sGetINI(sIniFile, "Variables", "TotalY_F", "?"))
    WxLnSubTotalX_F = Val(sGetINI(sIniFile, "Variables", "SubTotalX_F", "?"))
    WxLnSubTotalY_F = Val(sGetINI(sIniFile, "Variables", "SubTotalY_F", "?"))
    WxLnIGVX_F = Val(sGetINI(sIniFile, "Variables", "IGVX_F", "?"))
    WxLnIGVY_F = Val(sGetINI(sIniFile, "Variables", "IGVY_F", "?"))
    WxLnTerminalX_F = Val(sGetINI(sIniFile, "Variables", "TerminalX_F", "?"))
    WxLnTerminalY_F = Val(sGetINI(sIniFile, "Variables", "TerminalY_F", "?"))
    WxLnSerieImpresoraX_F = Val(sGetINI(sIniFile, "Variables", "SerieImpresoraX_F", "?"))
    WxLnSerieImpresoraY_F = Val(sGetINI(sIniFile, "Variables", "SerieImpresoraY_F", "?"))
    '
    WxLnCabeceraAlto_F = Val(sGetINI(sIniFile, "Variables", "CabeceraAlto_F", "?"))
    WxLnPieAlto_F = Val(sGetINI(sIniFile, "Variables", "PieAlto_F", "?"))
    
    WxLnNombreHoja_F = sGetINI(sIniFile, "Variables", "nombreHoja_F", "")
    WxLnTipoReporteador_F = Val(sGetINI(sIniFile, "Variables", "tipoReporteador_F", "0"))
    WxLnMargenIzquierdoX_F = Val(sGetINI(sIniFile, "Variables", "margenIzquierdoX_F", "100"))
    WxLnMargenDerechoX_F = Val(sGetINI(sIniFile, "Variables", "margenDerechoX_F", "100"))
    WxLnMargenSuperiorY_F = Val(sGetINI(sIniFile, "Variables", "margenSuperiorY_F", "100"))
    WxLnMargenInferiorY_F = Val(sGetINI(sIniFile, "Variables", "margenInferiorY_F", "100"))
    
    If lnIdTipoComprobanteDefault = 2 Then   'factura
       WxLnCabRucX = Val(sGetINI(sIniFile, "Variables", "CabRucX", "?"))
       WxLnCabRucY = Val(sGetINI(sIniFile, "Variables", "CabRucY", "?"))
       If WxLnCabRucX = 0 And WxLnCabRucY = 0 Then
            WxLnCabRucX = WxLnRzSocialPacienteX + 200
            WxLnCabRucY = WxLnRzSocialPacienteY
            Call sSetINI(sIniFile, "Variables", "CabRucX", CStr(WxLnCabRucX))
            Call sSetINI(sIniFile, "Variables", "CabRucY", CStr(WxLnCabRucY))
       End If
       
       WxLnCabRucX_F = Val(sGetINI(sIniFile, "Variables", "CabRucX_F", "?"))
       WxLnCabRucY_F = Val(sGetINI(sIniFile, "Variables", "CabRucY_F", "?"))
       If WxLnCabRucX_F = 0 And WxLnCabRucY_F = 0 Then
            WxLnCabRucX_F = WxLnRzSocialX_F + 200
            WxLnCabRucY_F = WxLnRzSocialY_F
            Call sSetINI(sIniFile, "Variables", "CabRucX_F", CStr(WxLnCabRucX_F))
            Call sSetINI(sIniFile, "Variables", "CabRucY_F", CStr(WxLnCabRucY_F))
       End If
    
    
       WxLnCabDireccionX = Val(sGetINI(sIniFile, "Variables", "CabDireccionX", "?"))
       WxLnCabDireccionY = Val(sGetINI(sIniFile, "Variables", "CabDireccionY", "?"))
       If WxLnCabDireccionX = 0 And WxLnCabDireccionY = 0 Then
            WxLnCabDireccionX = WxLnRzSocialPacienteX + 100
            WxLnCabDireccionY = WxLnRzSocialPacienteY
            Call sSetINI(sIniFile, "Variables", "CabDireccionX", CStr(WxLnCabDireccionX))
            Call sSetINI(sIniFile, "Variables", "CabDireccionY", CStr(WxLnCabDireccionY))
       End If
       
       
       WxLnCabDireccionX_F = Val(sGetINI(sIniFile, "Variables", "CabDireccionX_F", "?"))
       WxLnCabDireccionY_F = Val(sGetINI(sIniFile, "Variables", "CabDireccionY_F", "?"))
       If WxLnCabDireccionX_F = 0 And WxLnCabDireccionY_F = 0 Then
            WxLnCabDireccionX_F = WxLnRzSocialX_F + 100
            WxLnCabDireccionY_F = WxLnRzSocialY_F
            Call sSetINI(sIniFile, "Variables", "CabDireccionX_F", CStr(WxLnCabDireccionX_F))
            Call sSetINI(sIniFile, "Variables", "CabDireccionY_F", CStr(WxLnCabDireccionY_F))
       End If
    End If
    
    
 Else
    MsgBox "Se ha borrado el archivo de Configuracion", vbInformation, ""
 End If
 Exit Sub
ErrINI:
  If WxLnMargenInferiorY_F > 0 And lnIdTipoComprobanteDefault = 2 Then
     MsgBox "Es una FACTURA, debe ingresar las siguientes variables al archivo: C:\ARCHV..\DIG..\GALENHOS\ARCHIVOS\SETUP_CAJA_FACTURA.INI:" & Chr(13) & _
     "CabRucX=100" & Chr(13) & _
     "CabRucY=0" & Chr(13) & _
     "CabRucX_F=100" & Chr(13) & _
     "CabRucY_F=0" & Chr(13) & _
     "CabDireccionX=200" & Chr(13) & _
     "CabDireccionY=0" & Chr(13) & _
     "CabDireccionX_F=200" & Chr(13) & _
     "CabDireccionY_F=0" & Chr(13) & _
     "Configurar Formato de FACTURA para el RUC y DIRECCION en: migracion.exe ->procesos ->variosProcesos->tab Boleta"
  Else
     MsgBox Err.Description
  End If
  End
End Sub




'sunat
Public Function getPathIniFile(lcRutaINI As String, lnIdTipoComprobanteDefault As Long, lnBoletaFacturaFormTicket As Boolean) As String
    Dim sIniFile As String
    Select Case lnIdTipoComprobanteDefault
       Case 3      '"Boleta"
            If lnBoletaFacturaFormTicket = True Then
                sIniFile = lcRutaINI & "\setup_caja_boletaTicket.ini"
            Else
                sIniFile = lcRutaINI & "\setup_caja_boleta.ini"
            End If
       Case 2      '"Factura"
            If lnBoletaFacturaFormTicket = True Then
                sIniFile = lcRutaINI & "\setup_caja_facturaTicket.ini"
            Else
                sIniFile = lcRutaINI & "\setup_caja_factura.ini"
            End If
       Case 1      ' "Recibo"
           sIniFile = lcRutaINI & "\setup_caja_recibo.ini"
       Case 4       '"Ticket"
           sIniFile = lcRutaINI & "\setup_caja_ticket.ini"
    End Select
    getPathIniFile = sIniFile
End Function


Sub grabarSetup_Caja(ByVal lcRutaINI As String, ByVal lcTipoServicioFarmacia As String, _
                    ByVal lnIdTipoComprobanteDefault As Long)
 On Error GoTo ErrINI
    Dim sIniFile As String
 
    sIniFile = getPathIniFile(lcRutaINI, lnIdTipoComprobanteDefault, False)
 
    If lcTipoServicioFarmacia = "SERVICIOS" Then
        Call sSetINI(sIniFile, "Variables", "NumeroSerieX", CStr(WxLnNumeroSerieX))
        Call sSetINI(sIniFile, "Variables", "NumeroSerieY", CStr(WxLnNumeroSerieY))
        Call sSetINI(sIniFile, "Variables", "EstadoX", CStr(WxLnEstadoX))
        Call sSetINI(sIniFile, "Variables", "EstadoY", CStr(WxLnEstadoY))
        Call sSetINI(sIniFile, "Variables", "TipoX", CStr(WxLnTipoX))
        Call sSetINI(sIniFile, "Variables", "TipoY", CStr(WxLnTipoY))
        Call sSetINI(sIniFile, "Variables", "RzSocialX", CStr(WxLnRzSocialX))
        Call sSetINI(sIniFile, "Variables", "RzSocialY", CStr(WxLnRzSocialY))
        Call sSetINI(sIniFile, "Variables", "FechaX", CStr(WxLnFechaX))
        Call sSetINI(sIniFile, "Variables", "FechaY", CStr(WxLnFechaY))
        Call sSetINI(sIniFile, "Variables", "ServicioX", CStr(WxLnServicioX))
        Call sSetINI(sIniFile, "Variables", "ServicioY", CStr(WxLnServicioY))
        Call sSetINI(sIniFile, "Variables", "ObservacionesX", CStr(WxLnObservacionesX))
        Call sSetINI(sIniFile, "Variables", "ObservacionesY", CStr(WxLnObservacionesY))
        Call sSetINI(sIniFile, "Variables", "HistoriaX", CStr(WxLnHistoriaX))
        Call sSetINI(sIniFile, "Variables", "HistoriaY", CStr(WxLnHistoriaY))

        'JR 1005 (4L)
        Call sSetINI(sIniFile, "Variables", "NombrePaqueteX", CStr(WxLnNombrePaqueteX))
        Call sSetINI(sIniFile, "Variables", "NombrePaqueteY", CStr(WxLnNombrePaqueteY))
        Call sSetINI(sIniFile, "Variables", "DniPacienteX", CStr(WxLnDniPacienteX))
        Call sSetINI(sIniFile, "Variables", "DniPacienteY", CStr(WxLnDniPacienteY))
        
        Call sSetINI(sIniFile, "Variables", "CodigoY", CStr(WxLnCodigoY))
        Call sSetINI(sIniFile, "Variables", "ProductoY", CStr(WxLnProductoY))
        Call sSetINI(sIniFile, "Variables", "ProductoWidhtY", CStr(WxLnProductoWidhtY))
        Call sSetINI(sIniFile, "Variables", "CantidadY", CStr(WxLnCantidadY))
        Call sSetINI(sIniFile, "Variables", "PrecioY", CStr(WxLnPrecioY))
        Call sSetINI(sIniFile, "Variables", "ImporteY", CStr(WxLnImporteY))
        

        Call sSetINI(sIniFile, "Variables", "CajeroX", CStr(WxLnCajeroX))
        Call sSetINI(sIniFile, "Variables", "CajeroY", CStr(WxLnCajeroY))
        Call sSetINI(sIniFile, "Variables", "CajaX", CStr(WxLnCajaX))
        Call sSetINI(sIniFile, "Variables", "CajaY", CStr(WxLnCajaY))
        Call sSetINI(sIniFile, "Variables", "AdelantosX", CStr(WxLnAdelantosX))
        Call sSetINI(sIniFile, "Variables", "AdelantosY", CStr(WxLnAdelantosY))
        Call sSetINI(sIniFile, "Variables", "TotalPagarX", CStr(WxLnTotalPagarX))
        Call sSetINI(sIniFile, "Variables", "TotalPagarY", CStr(WxLnTotalPagarY))
        Call sSetINI(sIniFile, "Variables", "CuentaX", CStr(WxLnCuentaX))
        Call sSetINI(sIniFile, "Variables", "CuentaY", CStr(WxLnCuentaY))
        
        Call sSetINI(sIniFile, "Variables", "ExoneracionesX", CStr(WxLnExoneracionesX))
        Call sSetINI(sIniFile, "Variables", "ExoneracionesY", CStr(WxLnExoneracionesY))
        Call sSetINI(sIniFile, "Variables", "TotalEnLetrasX", CStr(WxLnTotalEnLetrasX))
        Call sSetINI(sIniFile, "Variables", "TotalEnLetrasY", CStr(WxLnTotalEnLetrasY))
        Call sSetINI(sIniFile, "Variables", "TotalLetrasWidhtY", CStr(WxLnTotalLetrasWidhtY))
        Call sSetINI(sIniFile, "Variables", "TotalX", CStr(WxLnTotalX))
        Call sSetINI(sIniFile, "Variables", "TotalY", CStr(WxLnTotalY))



        Call sSetINI(sIniFile, "Variables", "SubTotalX", CStr(WxLnSubTotalX))
        Call sSetINI(sIniFile, "Variables", "SubTotalY", CStr(WxLnSubTotalY))
        Call sSetINI(sIniFile, "Variables", "IGVX", CStr(WxLnIGVX))
        Call sSetINI(sIniFile, "Variables", "IGVY", CStr(WxLnIGVY))
        'JR 1005 (2L)
        Call sSetINI(sIniFile, "Variables", "UsuarioDespachoX", CStr(WxLnUsuarioDespachoX))
        Call sSetINI(sIniFile, "Variables", "UsuarioDespachoY", CStr(WxLnUsuarioDespachoY))
        
        Call sSetINI(sIniFile, "Variables", "CabeceraAlto", CStr(WxLnCabeceraAlto))
        Call sSetINI(sIniFile, "Variables", "PieAlto", CStr(WxLnPieAlto))

        'mgaray
        Call sSetINI(sIniFile, "Variables", "nombreHoja", CStr(WxLnNombreHoja))
        Call sSetINI(sIniFile, "Variables", "tipoReporteador", CStr(WxLnTipoReporteador))
        Call sSetINI(sIniFile, "Variables", "margenIzquierdoX", CStr(WxLnMargenIzquierdoX))
        Call sSetINI(sIniFile, "Variables", "margenDerechoX", CStr(WxLnMargenDerechoX))
        Call sSetINI(sIniFile, "Variables", "margenSuperiorY", CStr(WxLnMargenSuperiorY))
        Call sSetINI(sIniFile, "Variables", "margenInferiorY", CStr(WxLnMargenInferiorY))
        
        
        Call sSetINI(sIniFile, "Variables", "CabRucX", CStr(WxLnCabRucX))
        Call sSetINI(sIniFile, "Variables", "CabRucY", CStr(WxLnCabRucY))
        Call sSetINI(sIniFile, "Variables", "CabDireccionX", CStr(WxLnCabDireccionX))
        Call sSetINI(sIniFile, "Variables", "CabDireccionY", CStr(WxLnCabDireccionY))
    
    Else
    'Farmacia
        Call sSetINI(sIniFile, "Variables", "NumeroSerieX_F", CStr(WxLnNumeroSerieX_F))
        Call sSetINI(sIniFile, "Variables", "NumeroSerieY_F", CStr(WxLnNumeroSerieY_F))
        Call sSetINI(sIniFile, "Variables", "EstadoX_F", CStr(WxLnEstadoX_F))
        Call sSetINI(sIniFile, "Variables", "EstadoY_F", CStr(WxLnEstadoY_F))
        Call sSetINI(sIniFile, "Variables", "TipoX_F", CStr(WxLnTipoX_F))
        Call sSetINI(sIniFile, "Variables", "TipoY_F", CStr(WxLnTipoY_F))
        Call sSetINI(sIniFile, "Variables", "RzSocialX_F", CStr(WxLnRzSocialX_F))
        Call sSetINI(sIniFile, "Variables", "RzSocialY_F", CStr(WxLnRzSocialY_F))
        
        Call sSetINI(sIniFile, "Variables", "FechaX_F", CStr(WxLnFechaX_F))
        Call sSetINI(sIniFile, "Variables", "FechaY_F", CStr(WxLnFechaY_F))
        Call sSetINI(sIniFile, "Variables", "ServicioX_F", CStr(WxLnServicioX_F))
        Call sSetINI(sIniFile, "Variables", "ServicioY_F", CStr(WxLnServicioY_F))
        Call sSetINI(sIniFile, "Variables", "ObservacionesX_F", CStr(WxLnObservacionesX_F))
        Call sSetINI(sIniFile, "Variables", "ObservacionesY_F", CStr(WxLnObservacionesY_F))
        Call sSetINI(sIniFile, "Variables", "HistoriaX_F", CStr(WxLnHistoriaX_F))
        Call sSetINI(sIniFile, "Variables", "HistoriaY_F", CStr(WxLnHistoriaY_F))
        'JR 1005 (4L)
        Call sSetINI(sIniFile, "Variables", "NombrePaqueteX_F", CStr(WxLnNombrePaqueteX_F))
        Call sSetINI(sIniFile, "Variables", "NombrePaqueteY_F", CStr(WxLnNombrePaqueteY_F))
        Call sSetINI(sIniFile, "Variables", "DniPacienteX_F", CStr(WxLnDniPacienteX_F))
        Call sSetINI(sIniFile, "Variables", "DniPacienteY_F", CStr(WxLnDniPacienteY_F))
        
        Call sSetINI(sIniFile, "Variables", "CodigoY_F", CStr(WxLnCodigoY_F))
        Call sSetINI(sIniFile, "Variables", "ProductoY_F", CStr(WxLnProductoY_F))
        Call sSetINI(sIniFile, "Variables", "ProductoWidhtY_F", CStr(WxLnProductoWidhtY_F))
        Call sSetINI(sIniFile, "Variables", "CantidadY_F", CStr(WxLnCantidadY_F))
        Call sSetINI(sIniFile, "Variables", "PrecioY_F", CStr(WxLnPrecioY_F))
        Call sSetINI(sIniFile, "Variables", "ImporteY_F", CStr(WxLnImporteY_F))
        

        Call sSetINI(sIniFile, "Variables", "CajeroX_F", CStr(WxLnCajeroX_F))
        Call sSetINI(sIniFile, "Variables", "CajeroY_F", CStr(WxLnCajeroY_F))
        Call sSetINI(sIniFile, "Variables", "CajaX_F", CStr(WxLnCajaX_F))
        Call sSetINI(sIniFile, "Variables", "CajaY_F", CStr(WxLnCajaY_F))
        Call sSetINI(sIniFile, "Variables", "AdelantosX_F", CStr(WxLnAdelantosX_F))
        Call sSetINI(sIniFile, "Variables", "AdelantosY_F", CStr(WxLnAdelantosY_F))
        Call sSetINI(sIniFile, "Variables", "TotalPagarX_F", CStr(WxLnTotalPagarX_F))
        Call sSetINI(sIniFile, "Variables", "TotalPagarY_F", CStr(WxLnTotalPagarY_F))
        Call sSetINI(sIniFile, "Variables", "CuentaX_F", CStr(WxLnCuentaX_F))
        Call sSetINI(sIniFile, "Variables", "CuentaY_F", CStr(WxLnCuentaY_F))
        Call sSetINI(sIniFile, "Variables", "ExoneracionesX_F", CStr(WxLnExoneracionesX_F))
        Call sSetINI(sIniFile, "Variables", "ExoneracionesY_F", CStr(WxLnExoneracionesY_F))

        Call sSetINI(sIniFile, "Variables", "TotalEnLetrasX_F", CStr(WxLnTotalEnLetrasX_F))
        Call sSetINI(sIniFile, "Variables", "TotalEnLetrasY_F", CStr(WxLnTotalEnLetrasY_F))
        Call sSetINI(sIniFile, "Variables", "TotalLetrasWidhtY_F", CStr(WxLnTotalLetrasWidhtY_F))
        Call sSetINI(sIniFile, "Variables", "TotalX_F", CStr(WxLnTotalX_F))
        Call sSetINI(sIniFile, "Variables", "TotalY_F", CStr(WxLnTotalY_F))
        Call sSetINI(sIniFile, "Variables", "SubTotalX_F", CStr(WxLnSubTotalX_F))
        Call sSetINI(sIniFile, "Variables", "SubTotalY_F", CStr(WxLnSubTotalY_F))
        Call sSetINI(sIniFile, "Variables", "IGVX_F", CStr(WxLnIGVX_F))
        Call sSetINI(sIniFile, "Variables", "IGVY_F", CStr(WxLnIGVY_F))
        'JR 1005 (2L)
        Call sSetINI(sIniFile, "Variables", "UsuarioDespachoX_F", CStr(WxLnUsuarioDespachoX_F))
        Call sSetINI(sIniFile, "Variables", "UsuarioDespachoY_F", CStr(WxLnUsuarioDespachoY_F))

        Call sSetINI(sIniFile, "Variables", "CabeceraAlto_F", CStr(WxLnCabeceraAlto_F))
        Call sSetINI(sIniFile, "Variables", "PieAlto_F", CStr(WxLnPieAlto_F))
        
        'mgaray
        Call sSetINI(sIniFile, "Variables", "nombreHoja_F", CStr(WxLnNombreHoja_F))
        Call sSetINI(sIniFile, "Variables", "tipoReporteador_F", CStr(WxLnTipoReporteador_F))
        Call sSetINI(sIniFile, "Variables", "margenIzquierdoX_F", CStr(WxLnMargenIzquierdoX_F))
        Call sSetINI(sIniFile, "Variables", "margenDerechoX_F", CStr(WxLnMargenDerechoX_F))
        Call sSetINI(sIniFile, "Variables", "margenSuperiorY_F", CStr(WxLnMargenSuperiorY_F))
        Call sSetINI(sIniFile, "Variables", "margenInferiorY_F", CStr(WxLnMargenInferiorY_F))
        
        Call sSetINI(sIniFile, "Variables", "CabRucX_F", CStr(WxLnCabRucX_F))
        Call sSetINI(sIniFile, "Variables", "CabRucY_F", CStr(WxLnCabRucY_F))
        Call sSetINI(sIniFile, "Variables", "CabDireccionX_F", CStr(WxLnCabDireccionX_F))
        Call sSetINI(sIniFile, "Variables", "CabDireccionY_F", CStr(WxLnCabDireccionY_F))
        
        
    End If
    MsgBox "Configuración Grabada con Exito en : " & Chr(13) & sIniFile, vbInformation
    Exit Sub
ErrINI:
  End
End Sub
'=================================================

Function sSetINI(sIniFile As String, sSection As String, sKey As String, sDefault As String) As String
    Call WritePrivateProfileString(sSection, sKey, sDefault, sIniFile)
End Function

Sub CargaSetup_X_PC()
 On Error GoTo ErrINI
 WxDEFAULT_BUSQ_PACIENTE = sghDefaultVentana.sighDNI
 WxDEFAULT_BUSQ_CE = sghDefaultVentana.sighApellidoPaterno
 WxDEFAULT_BUSQ_EMERGENCIA = sghDefaultVentana.sighApellidoPaterno
 WxDEFAULT_BUSQ_HOSPITALIZ = sghDefaultVentana.sighApellidoPaterno
 Dim sIniFile As String
 sIniFile = App.Path & "\archivos\setup_x_pc.ini"
 If Dir$(sIniFile) <> "" Then
    WxDEFAULT_BUSQ_PACIENTE = Val(sGetINI(sIniFile, "Variables", "DEFAULT_BUSQ_PACIENTE", "?"))
    WxDEFAULT_BUSQ_CE = Val(sGetINI(sIniFile, "Variables", "DEFAULT_BUSQ_CE", "?"))
    WxDEFAULT_BUSQ_EMERGENCIA = Val(sGetINI(sIniFile, "Variables", "DEFAULT_BUSQ_EMERGENCIA", "?"))
    WxDEFAULT_BUSQ_HOSPITALIZ = Val(sGetINI(sIniFile, "Variables", "DEFAULT_BUSQ_HOSPITALIZ", "?"))
  End If
ErrINI:
End Sub

'Frank Configuracion Boleta Farmacia - Ventas
Sub CargaSetup_FarmVentas(lcRutaINI As String, lnIdTipoComprobanteDefault As Long)
 On Error GoTo ErrINI
 Dim sIniFile As String
 Select Case lnIdTipoComprobanteDefault
 Case 1      '"Boleta"
    sIniFile = lcRutaINI & "\setup_farmventas_boleta.ini"
 Case 2      '"Ticket"
     sIniFile = lcRutaINI & "\setup_farmventas_ticket.ini"
 End Select
 If Dir$(sIniFile) <> "" Then
    WxLnCabeceraAlto = Val(sGetINI(sIniFile, "Variables", "CabeceraAlto", "?"))
    WxLnPieAlto = Val(sGetINI(sIniFile, "Variables", "PieAlto", "?"))
    WxLnMargenIzquierdoX = Val(sGetINI(sIniFile, "Variables", "margenIzquierdoX", "100"))
    WxLnMargenDerechoX = Val(sGetINI(sIniFile, "Variables", "margenDerechoX", "100"))
    WxLnMargenSuperiorY = Val(sGetINI(sIniFile, "Variables", "margenSuperiorY", "100"))
    WxLnMargenInferiorY = Val(sGetINI(sIniFile, "Variables", "margenInferiorY", "100"))
    WxLnNombreEESSX = Val(sGetINI(sIniFile, "Variables", "NombreEESSX", "?"))
    WxLnNombreEESSY = Val(sGetINI(sIniFile, "Variables", "NombreEESSY", "?"))
    WxLnDireccionEESSX = Val(sGetINI(sIniFile, "Variables", "DireccionEESSX", "?"))
    WxLnDireccionEESSY = Val(sGetINI(sIniFile, "Variables", "DireccionEESSY", "?"))
    WxLnTelefonoEESSX = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSX", "?"))
    WxLnTelefonoEESSY = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSY", "?"))
    WxLnTipoFormatoX = Val(sGetINI(sIniFile, "Variables", "TipoFormatoX", "?"))
    WxLnTipoFormatoY = Val(sGetINI(sIniFile, "Variables", "TipoFormatoY", "?"))
    WxLnFarmaciaX = Val(sGetINI(sIniFile, "Variables", "FarmaciaX", "?"))
    WxLnFarmaciaY = Val(sGetINI(sIniFile, "Variables", "FarmaciaY", "?"))
    WxLnPacienteX = Val(sGetINI(sIniFile, "Variables", "PacienteX", "?"))
    WxLnPacienteY = Val(sGetINI(sIniFile, "Variables", "PacienteY", "?"))
    WxLnFUAX = Val(sGetINI(sIniFile, "Variables", "FUAX", "?"))
    WxLnFUAY = Val(sGetINI(sIniFile, "Variables", "FUAY", "?"))
    WxLnDiagPrincipalX = Val(sGetINI(sIniFile, "Variables", "DiagPrincipalX", "?"))
    WxLnDiagPrincipalY = Val(sGetINI(sIniFile, "Variables", "DiagPrincipalY", "?"))
    WxLnNroCuentaX = Val(sGetINI(sIniFile, "Variables", "NroCuentaX", "?"))
    WxLnNroCuentaY = Val(sGetINI(sIniFile, "Variables", "NroCuentaY", "?"))
    WxLnServicioHospX = Val(sGetINI(sIniFile, "Variables", "ServicioHospX", "?"))
    WxLnServicioHospY = Val(sGetINI(sIniFile, "Variables", "ServicioHospY", "?"))
    WxLnNroMovmientoX = Val(sGetINI(sIniFile, "Variables", "NroMovmientoX", "?"))
    WxLnNroMovmientoY = Val(sGetINI(sIniFile, "Variables", "NroMovmientoY", "?"))
    WxLnFechaMovimientoX = Val(sGetINI(sIniFile, "Variables", "FechaMovimientoX", "?"))
    WxLnFechaMovimientoY = Val(sGetINI(sIniFile, "Variables", "FechaMovimientoY", "?"))
    WxLnNItemY = Val(sGetINI(sIniFile, "Variables", "NItemY", "?"))
    WxLnCodigoY = Val(sGetINI(sIniFile, "Variables", "CodigoY", "?"))
    WxLnProductoY = Val(sGetINI(sIniFile, "Variables", "ProductoY", "?"))
    WxLnProductoWidhtY = Val(sGetINI(sIniFile, "Variables", "ProductoWidhtY", "?"))
    WxLnCantidadY = Val(sGetINI(sIniFile, "Variables", "CantidadY", "?"))
    WxLnPrecioY = Val(sGetINI(sIniFile, "Variables", "PrecioY", "?"))
    WxLnImporteY = Val(sGetINI(sIniFile, "Variables", "ImporteY", "?"))
    WxLnTotalPagarX = Val(sGetINI(sIniFile, "Variables", "TotalPagarX", "?"))
    WxLnTotalPagarY = Val(sGetINI(sIniFile, "Variables", "TotalPagarY", "?"))
    WxLnNombrePaqueteX = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteX", "?"))
    WxLnNombrePaqueteY = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteY", "?"))
    WxLnDniPacienteX = Val(sGetINI(sIniFile, "Variables", "DniPacienteX", "?"))
    WxLnDniPacienteY = Val(sGetINI(sIniFile, "Variables", "DniPacienteY", "?"))
    WxLnUsuarioDespachoX = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoX", "?"))
    WxLnUsuarioDespachoY = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoY", "?"))
    WxLnTotalItemsX = Val(sGetINI(sIniFile, "Variables", "TotalItemsX", "?"))
    WxLnTotalItemsY = Val(sGetINI(sIniFile, "Variables", "TotalItemsY", "?"))
 Else
    MsgBox "Se ha borrado el archivo de Configuracion", vbInformation, ""
 End If
 Exit Sub
ErrINI:
  End
End Sub


Sub grabarSetup_FarmVenta(ByVal lcRutaINI As String, ByVal lnIdTipoComprobanteDefault As Long)
 On Error GoTo ErrINI
    Dim sIniFile As String
 
'    sIniFile = getPathIniFile(lcRutaINI, lnIdTipoComprobanteDefault)
    Select Case lnIdTipoComprobanteDefault
    Case 1      '"Boleta"
       sIniFile = lcRutaINI & "\setup_farmventas_boleta.ini"
    Case 2      '"Ticket"
        sIniFile = lcRutaINI & "\setup_farmventas_ticket.ini"
    End Select
    
    Call sSetINI(sIniFile, "Variables", "NombreEESSX", CStr(WxLnNombreEESSX))
    Call sSetINI(sIniFile, "Variables", "NombreEESSY", CStr(WxLnNombreEESSY))
    Call sSetINI(sIniFile, "Variables", "DireccionEESSX", CStr(WxLnDireccionEESSX))
    Call sSetINI(sIniFile, "Variables", "DireccionEESSY", CStr(WxLnDireccionEESSY))
    Call sSetINI(sIniFile, "Variables", "TelefonoEESSX", CStr(WxLnTelefonoEESSX))
    Call sSetINI(sIniFile, "Variables", "TelefonoEESSY", CStr(WxLnTelefonoEESSY))
    Call sSetINI(sIniFile, "Variables", "TipoFormatoX", CStr(WxLnTipoFormatoX))
    Call sSetINI(sIniFile, "Variables", "TipoFormatoY", CStr(WxLnTipoFormatoY))
    Call sSetINI(sIniFile, "Variables", "FarmaciaX", CStr(WxLnFarmaciaX))
    Call sSetINI(sIniFile, "Variables", "FarmaciaY", CStr(WxLnFarmaciaY))
    Call sSetINI(sIniFile, "Variables", "PacienteX", CStr(WxLnPacienteX))
    Call sSetINI(sIniFile, "Variables", "PacienteY", CStr(WxLnPacienteY))
    Call sSetINI(sIniFile, "Variables", "FUAX", CStr(WxLnFUAX))
    Call sSetINI(sIniFile, "Variables", "FUAY", CStr(WxLnFUAY))
    Call sSetINI(sIniFile, "Variables", "DiagPrincipalX", CStr(WxLnDiagPrincipalX))
    Call sSetINI(sIniFile, "Variables", "DiagPrincipalY", CStr(WxLnDiagPrincipalY))
    Call sSetINI(sIniFile, "Variables", "NroCuentaX", CStr(WxLnNroCuentaX))
    Call sSetINI(sIniFile, "Variables", "NroCuentaY", CStr(WxLnNroCuentaY))
    Call sSetINI(sIniFile, "Variables", "ServicioHospX", CStr(WxLnServicioHospX))
    Call sSetINI(sIniFile, "Variables", "ServicioHospY", CStr(WxLnServicioHospY))
    Call sSetINI(sIniFile, "Variables", "NroMovmientoX", CStr(WxLnNroMovmientoX))
    Call sSetINI(sIniFile, "Variables", "NroMovmientoY", CStr(WxLnNroMovmientoY))
    Call sSetINI(sIniFile, "Variables", "FechaMovimientoX", CStr(WxLnFechaMovimientoX))
    Call sSetINI(sIniFile, "Variables", "FechaMovimientoY", CStr(WxLnFechaMovimientoY))
    
    Call sSetINI(sIniFile, "Variables", "CodigoY", CStr(WxLnCodigoY))
    Call sSetINI(sIniFile, "Variables", "ProductoY", CStr(WxLnProductoY))
    Call sSetINI(sIniFile, "Variables", "ProductoWidhtY", CStr(WxLnProductoWidhtY))
    Call sSetINI(sIniFile, "Variables", "CantidadY", CStr(WxLnCantidadY))
    Call sSetINI(sIniFile, "Variables", "PrecioY", CStr(WxLnPrecioY))
    Call sSetINI(sIniFile, "Variables", "ImporteY", CStr(WxLnImporteY))
    Call sSetINI(sIniFile, "Variables", "NItemY", CStr(WxLnNItemY))
    
    Call sSetINI(sIniFile, "Variables", "TotalPagarX", CStr(WxLnTotalPagarX))
    Call sSetINI(sIniFile, "Variables", "TotalPagarY", CStr(WxLnTotalPagarY))
    
    Call sSetINI(sIniFile, "Variables", "TotalItemsX", CStr(WxLnTotalItemsX))
    Call sSetINI(sIniFile, "Variables", "TotalItemsY", CStr(WxLnTotalItemsY))
        
    Call sSetINI(sIniFile, "Variables", "CabeceraAlto", CStr(WxLnCabeceraAlto))
    Call sSetINI(sIniFile, "Variables", "PieAlto", CStr(WxLnPieAlto))
    
    Call sSetINI(sIniFile, "Variables", "margenIzquierdoX", CStr(WxLnMargenIzquierdoX))
    Call sSetINI(sIniFile, "Variables", "margenDerechoX", CStr(WxLnMargenDerechoX))
    Call sSetINI(sIniFile, "Variables", "margenSuperiorY", CStr(WxLnMargenSuperiorY))
    Call sSetINI(sIniFile, "Variables", "margenInferiorY", CStr(WxLnMargenInferiorY))
    
    'JR 04042016
    Call sSetINI(sIniFile, "Variables", "NombrePaqueteX", CStr(WxLnNombrePaqueteX))
    Call sSetINI(sIniFile, "Variables", "NombrePaqueteY", CStr(WxLnNombrePaqueteY))
    Call sSetINI(sIniFile, "Variables", "DniPacienteX", CStr(WxLnDniPacienteX))
    Call sSetINI(sIniFile, "Variables", "DniPacienteY", CStr(WxLnDniPacienteY))
    Call sSetINI(sIniFile, "Variables", "UsuarioDespachoX", CStr(WxLnUsuarioDespachoX))
    Call sSetINI(sIniFile, "Variables", "UsuarioDespachoY", CStr(WxLnUsuarioDespachoY))
        
    
    MsgBox "Configuración Grabada con Exito en : " & Chr(13) & sIniFile, vbInformation
    Exit Sub
ErrINI:
  End
End Sub

'JR 0628
Sub CargaSetup_LabResult(lcRutaINI As String, lnIdTipoComprobanteDefault As Long)
 On Error GoTo ErrINI
 Dim sIniFile As String
 Select Case lnIdTipoComprobanteDefault
 Case 1      '"Boleta"
    sIniFile = lcRutaINI & "\setup_LabResult_boleta.ini"
 Case 2      '"Ticket"
     sIniFile = lcRutaINI & "\setup_LabResult_ticket.ini"
 End Select
 If Dir$(sIniFile) <> "" Then
    WxLnCabeceraAlto = Val(sGetINI(sIniFile, "Variables", "CabeceraAlto", "?"))
    WxLnPieAlto = Val(sGetINI(sIniFile, "Variables", "PieAlto", "?"))
    WxLnMargenIzquierdoX = Val(sGetINI(sIniFile, "Variables", "margenIzquierdoX", "100"))
    WxLnMargenDerechoX = Val(sGetINI(sIniFile, "Variables", "margenDerechoX", "100"))
    WxLnMargenSuperiorY = Val(sGetINI(sIniFile, "Variables", "margenSuperiorY", "100"))
    WxLnMargenInferiorY = Val(sGetINI(sIniFile, "Variables", "margenInferiorY", "100"))
    WxLnNombreEESSX = Val(sGetINI(sIniFile, "Variables", "NombreEESSX", "?"))
    WxLnNombreEESSY = Val(sGetINI(sIniFile, "Variables", "NombreEESSY", "?"))
    WxLnDireccionEESSX = Val(sGetINI(sIniFile, "Variables", "DireccionEESSX", "?"))
    WxLnDireccionEESSY = Val(sGetINI(sIniFile, "Variables", "DireccionEESSY", "?"))
    WxLnTelefonoEESSX = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSX", "?"))
    WxLnTelefonoEESSY = Val(sGetINI(sIniFile, "Variables", "TelefonoEESSY", "?"))
    WxLnTipoFormatoX = Val(sGetINI(sIniFile, "Variables", "TipoFormatoX", "?"))
    WxLnTipoFormatoY = Val(sGetINI(sIniFile, "Variables", "TipoFormatoY", "?"))
    WxLnFarmaciaX = Val(sGetINI(sIniFile, "Variables", "FarmaciaX", "?"))
    WxLnFarmaciaY = Val(sGetINI(sIniFile, "Variables", "FarmaciaY", "?"))
    WxLnPacienteX = Val(sGetINI(sIniFile, "Variables", "PacienteX", "?"))
    WxLnPacienteY = Val(sGetINI(sIniFile, "Variables", "PacienteY", "?"))
    WxLnDiagPrincipalX = Val(sGetINI(sIniFile, "Variables", "DiagPrincipalX", "?"))
    WxLnDiagPrincipalY = Val(sGetINI(sIniFile, "Variables", "DiagPrincipalY", "?"))
    WxLnNroCuentaX = Val(sGetINI(sIniFile, "Variables", "NroCuentaX", "?"))
    WxLnNroCuentaY = Val(sGetINI(sIniFile, "Variables", "NroCuentaY", "?"))
    WxLnServicioHospX = Val(sGetINI(sIniFile, "Variables", "ServicioHospX", "?"))
    WxLnServicioHospY = Val(sGetINI(sIniFile, "Variables", "ServicioHospY", "?"))
    WxLnNroMovmientoX = Val(sGetINI(sIniFile, "Variables", "NroMovmientoX", "?"))
    WxLnNroMovmientoY = Val(sGetINI(sIniFile, "Variables", "NroMovmientoY", "?"))
    WxLnFechaMovimientoX = Val(sGetINI(sIniFile, "Variables", "FechaMovimientoX", "?"))
    WxLnFechaMovimientoY = Val(sGetINI(sIniFile, "Variables", "FechaMovimientoY", "?"))
    WxLnCodigoY = Val(sGetINI(sIniFile, "Variables", "CodigoY", "?"))
    WxLnProductoY = Val(sGetINI(sIniFile, "Variables", "ProductoY", "?"))
    WxLnProductoWidhtY = Val(sGetINI(sIniFile, "Variables", "ProductoWidhtY", "?"))
    WxLnCantidadY = Val(sGetINI(sIniFile, "Variables", "CantidadY", "?"))
    WxLnPrecioY = Val(sGetINI(sIniFile, "Variables", "PrecioY", "?"))
    WxLnImporteY = Val(sGetINI(sIniFile, "Variables", "ImporteY", "?"))
    WxLnTotalPagarX = Val(sGetINI(sIniFile, "Variables", "TotalPagarX", "?"))
    WxLnTotalPagarY = Val(sGetINI(sIniFile, "Variables", "TotalPagarY", "?"))
    'JR 04042016
    WxLnNombrePaqueteX = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteX", "?"))
    WxLnNombrePaqueteY = Val(sGetINI(sIniFile, "Variables", "NombrePaqueteY", "?"))
    WxLnDniPacienteX = Val(sGetINI(sIniFile, "Variables", "DniPacienteX", "?"))
    WxLnDniPacienteY = Val(sGetINI(sIniFile, "Variables", "DniPacienteY", "?"))
    WxLnUsuarioDespachoX = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoX", "?"))
    WxLnUsuarioDespachoY = Val(sGetINI(sIniFile, "Variables", "UsuarioDespachoY", "?"))
 Else
    MsgBox "Se ha borrado el archivo de Configuracion", vbInformation, ""
 End If
 Exit Sub
ErrINI:
  End
End Sub


