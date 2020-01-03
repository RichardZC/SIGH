VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Teclado As New sighentidades.Teclado

'mgaray
Dim lcBoletaTipo As String, lcBoletaEstado As String, lcRazonSocialHC As String
Dim lcTotalEnLetras As String, lnIdCaja As Long, lcObservaciones As String
Dim lcTipoPago As String, lcTipoComprobante As String

Enum reporteadorAUsarEnum
    DataReport = 0
    CrystalReport = 1
End Enum

Sub ImprimeBoletaEnFormatoUsuario(rsReporte As Recordset, lcCajeroIniciales As String, lcNombreServicio As String, _
                                  lnTotal As Double, lnImporteEXO As Double, lcCuenta As String, lcAdelantos As String, _
                                  lcSerieImpresoraDefault As String, lcSerieImpresora2 As String, lbMuestraVistaPrevia As Boolean, lcBoletaNumero As String, _
                                  lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsTicket As Boolean, _
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection, Optional lbEsCopiaDeBoleta As Boolean)
                                  
'    Dim lcBoletaTipo As String, lcBoletaEstado As String, lcRazonSocialHC As String
'    Dim lcTotalEnLetras As String, lnIdCaja As Long, lcObservaciones As String
'    Dim lcTipoPago As String, lcTipoComprobante As String
    '
    rsReporte.MoveFirst
    lcTipoPago = IIf(rsReporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsReporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))
    If lbEsCopiaDeBoleta = True Then
       lcTipoPago = "...COPIA..."
    End If
    lcBoletaEstado = IIf(rsReporte.Fields!IdEstadoComprobante = 9, "A N U L A D O ", "..")
    'lcRazonSocialHC = Left(Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial)) + IIf(IsNull(rsReporte.Fields!NroHistoriaClinica), " (Particular)", " (HC " & Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & ")") + String(45, " "), 40)
    lcRazonSocialHC = Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial))
    If IsNull(rsReporte.Fields!NroHistoriaClinica) Then
       lcRazonSocialHC = " (Particular) " & lcRazonSocialHC
    Else
       lcRazonSocialHC = lcRazonSocialHC & " (HC " & Trim(Str(rsReporte.Fields!NroHistoriaClinica))
    End If
    lcRazonSocialHC = Left(lcRazonSocialHC + String(45, " "), 40)
    
    '
    lcTotalEnLetras = "Son: " & sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighentidades.DevuelveSoloDecimales(lnTotal) & "/100   Nuevos Soles"
    '
    lnIdCaja = rsReporte.Fields!IdCaja
    If lbEsBoletaDeFarmacia = True Then
        SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, False, oConexion)
    Else
        SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True, oConexion)
    End If
    If lbElDocumentoEsTicket = True Then
        '************************** ticket ***********************************************
        Set FormBoleTicket.DataSource = rsReporte
        FormBoleTicket.Sections("cabecera").Controls("lblHospital").Caption = UCase(wxParametro205)
        FormBoleTicket.Sections("cabecera").Controls("DireccionEESS").Caption = mo_Teclado.CapitalizarNombres(wxParametro206)
        FormBoleTicket.Sections("cabecera").Controls("RucEESS").Caption = UCase("RUC: ") & (wxParametro339)
        FormBoleTicket.Sections("cabecera").Controls("TelefonoEESS").Caption = UCase("TELF: ") & (wxParametro207)
        FormBoleTicket.Caption = lcTipoComprobante
        FormBoleTicket.Sections("cabecera").Controls("NombreCabecera").Caption = "TICKET DE VENTA"
        FormBoleTicket.Sections("cabecera").Controls("BoletaNumeroSerie").Caption = "Nº " & lcBoletaNumero
        FormBoleTicket.Sections("cabecera").Controls("BoletaEstado").Caption = lcBoletaEstado
        FormBoleTicket.Sections("cabecera").Controls("BoletaTipo").Caption = lcTipoPago
        FormBoleTicket.Sections("cabecera").Controls("NombreCabeceraPaciente").Caption = "PACIENTE"
        'Modificado por Frank
'        FormBoleTicket.Sections("cabecera").Controls("RazonSocial").Caption = lcRazonSocialHC
        FormBoleTicket.Sections("cabecera").Controls("Historia").Caption = "Nº HISTORIA : " & IIf(IsNull(rsReporte.Fields!NroHistoriaClinica), "(Particular)", rsReporte.Fields!NroHistoriaClinica) & " ,   Nº CUENTA : " & IIf(IsNull(rsReporte.Fields!idCuentaAtencion), "...", rsReporte.Fields!idCuentaAtencion)
        FormBoleTicket.Sections("cabecera").Controls("RazonSocial").Caption = UCase(Left(Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial)) + String(33, " "), 33))
        'Modificado por Frank
        FormBoleTicket.Sections("cabecera").Controls("FechaCobranza").Caption = UCase("FECHA:  ") & Format(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY) & "                       HORA: " & Format$(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveHoraSoloFormato_HM)
        FormBoleTicket.Sections("cabecera").Controls("Servicio").Caption = UCase("UND. SERVICIO : " & lcNombreServicio)
        FormBoleTicket.Sections("cabecera").Controls("Observaciones").Caption = UCase("Observaciones: ") & (IIf(IsNull(rsReporte.Fields!Observaciones), "", rsReporte.Fields!Observaciones))
        
        FormBoleTicket.Sections("PieInforme").Controls("nombreCaja").Caption = UCase("Caja: " & IIf(IsNull(rsReporte.Fields!nombreCaja), "", rsReporte.Fields!nombreCaja))
        FormBoleTicket.Sections("PieInforme").Controls("Cajero").Caption = UCase("Cajero(A) : " & lcCajeroIniciales)
        FormBoleTicket.Sections("PieInforme").Controls("TotalPorPagar").Caption = "SubTotal:  " & Format(lnTotal + lnImporteEXO, "######.#0")
        FormBoleTicket.Sections("PieInforme").Controls("idCuentaAtencion").Caption = lcCuenta
        FormBoleTicket.Sections("PieInforme").Controls("Adelantos").Caption = lcAdelantos
        FormBoleTicket.Sections("PieInforme").Controls("Exoneraciones").Caption = "Exoneraciones  :" & Format(lnImporteEXO, "######.#0")
        FormBoleTicket.Sections("PieInforme").Controls("TotalEnLetras").Caption = lcTotalEnLetras
        FormBoleTicket.Sections("PieInforme").Controls("TotalEnLetras").Width = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalLetrasWidhtY_F, WxLnTotalLetrasWidhtY)
        FormBoleTicket.Sections("PieInforme").Controls("TotalBoleta").Caption = "Total: S/. " & Format(lnTotal, "######.#0")
        FormBoleTicket.Sections("PieInforme").Controls("Terminal").Caption = "Terminal:  " & lcBuscaParametro.RetornaNombreDeServidor
        FormBoleTicket.Sections("PieInforme").Controls("SerieImpresora").Caption = "Nº Serie Impresora:  " & lcSerieImpresoraDefault
        FormBoleTicket.Sections("PieInforme").Controls("TextoVacio").Caption = wxParametro346
        
        FormBoleTicket.RightMargin = getReportRightMargin
        FormBoleTicket.TopMargin = getReportTopMargin
        FormBoleTicket.LeftMargin = getReportLeftMargin
        FormBoleTicket.BottomMargin = getReportBottomMargin
        
        FormBoleTicket.Sections("cabecera").Controls("DireccionEESS").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnDireccionEESSX_F, WxLnDireccionEESSX)
        FormBoleTicket.Sections("cabecera").Controls("DireccionEESS").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnDireccionEESSY_F, WxLnDireccionEESSY)
        FormBoleTicket.Sections("cabecera").Controls("RucEESS").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnRucEESSX_F, WxLnRucEESSX)
        FormBoleTicket.Sections("cabecera").Controls("RucEESS").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnRucEESSY_F, WxLnRucEESSY)
        FormBoleTicket.Sections("cabecera").Controls("TelefonoEESS").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTelefonoEESSX_F, WxLnTelefonoEESSX)
        FormBoleTicket.Sections("cabecera").Controls("TelefonoEESS").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTelefonoEESSY_F, WxLnTelefonoEESSY)

        FormBoleTicket.Sections("cabecera").Controls("NombreCabecera").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieX_F, WxLnNumeroSerieX)
        FormBoleTicket.Sections("cabecera").Controls("NombreCabecera").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieY_F, WxLnNumeroSerieY)
        FormBoleTicket.Sections("cabecera").Controls("BoletaNumeroSerie").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieX_F + 250, WxLnNumeroSerieX + 250)
        FormBoleTicket.Sections("cabecera").Controls("BoletaNumeroSerie").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieY_F, WxLnNumeroSerieY)
        FormBoleTicket.Sections("cabecera").Controls("BoletaEstado").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoX_F, WxLnEstadoX)
        FormBoleTicket.Sections("cabecera").Controls("BoletaEstado").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoY_F, WxLnEstadoY)
        FormBoleTicket.Sections("cabecera").Controls("BoletaTipo").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoX_F, WxLnTipoX)
        FormBoleTicket.Sections("cabecera").Controls("BoletaTipo").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoY_F, WxLnTipoY)
        FormBoleTicket.Sections("cabecera").Controls("NombreCabeceraPaciente").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialX_F - 160, WxLnRzSocialX - 160)
        FormBoleTicket.Sections("cabecera").Controls("NombreCabeceraPaciente").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialY_F, WxLnRzSocialY)
        FormBoleTicket.Sections("cabecera").Controls("RazonSocial").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialX_F, WxLnRzSocialX)
        FormBoleTicket.Sections("cabecera").Controls("RazonSocial").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialY_F, WxLnRzSocialY)
        
        'Modificado Frank
        FormBoleTicket.Sections("cabecera").Controls("Historia").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnHistoriaX_F, WxLnHistoriaX)
        FormBoleTicket.Sections("cabecera").Controls("Historia").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnHistoriaY_F, WxLnHistoriaY)
        FormBoleTicket.Sections("cabecera").Controls("FechaCobranza").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaX_F, WxLnFechaX)
        FormBoleTicket.Sections("cabecera").Controls("FechaCobranza").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaY_F, WxLnFechaY)
        FormBoleTicket.Sections("cabecera").Controls("Servicio").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioX_F, WxLnServicioX)
        FormBoleTicket.Sections("cabecera").Controls("Servicio").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioY_F, WxLnServicioY)
        FormBoleTicket.Sections("cabecera").Controls("Observaciones").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesX_F, WxLnObservacionesX)
        FormBoleTicket.Sections("cabecera").Controls("Observaciones").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesY_F, WxLnObservacionesY)
       
        'FormBoleTicket.Sections("Detalle").Controls("txtCodigo").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCodigoY_F, WxLnCodigoY)
        FormBoleTicket.Sections("Detalle").Controls("txtNombreProducto").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoY_F, WxLnProductoY)
        FormBoleTicket.Sections("Detalle").Controls("txtNombreProducto").Width = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoWidhtY_F, WxLnProductoWidhtY)
        FormBoleTicket.Sections("Detalle").Controls("txtCantidad").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCantidadY_F, WxLnCantidadY)
        'FormBoleTicket.Sections("Detalle").Controls("txtPrecioUnitario").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnPrecioY_F, WxLnPrecioY)
        FormBoleTicket.Sections("Detalle").Controls("txtTotalPorPagar").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnImporteY_F, WxLnImporteY)
        
        FormBoleTicket.Sections("PieInforme").Controls("nombreCaja").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaX_F, WxLnCajaX)
        FormBoleTicket.Sections("PieInforme").Controls("nombreCaja").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaY_F, WxLnCajaY)
        FormBoleTicket.Sections("PieInforme").Controls("Cajero").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroX_F, WxLnCajeroX)
        FormBoleTicket.Sections("PieInforme").Controls("Cajero").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroY_F, WxLnCajeroY)
        FormBoleTicket.Sections("PieInforme").Controls("Adelantos").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosX_F, WxLnAdelantosX)
        FormBoleTicket.Sections("PieInforme").Controls("Adelantos").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosY_F, WxLnAdelantosY)
        FormBoleTicket.Sections("PieInforme").Controls("TotalPorPagar").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarX_F, WxLnTotalPagarX)
        FormBoleTicket.Sections("PieInforme").Controls("TotalPorPagar").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarY_F, WxLnTotalPagarY)
        FormBoleTicket.Sections("PieInforme").Controls("idCuentaAtencion").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaX_F, WxLnCuentaX)
        FormBoleTicket.Sections("PieInforme").Controls("idCuentaAtencion").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaY_F, WxLnCuentaY)
        FormBoleTicket.Sections("PieInforme").Controls("Exoneraciones").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesX_F, WxLnExoneracionesX)
        FormBoleTicket.Sections("PieInforme").Controls("Exoneraciones").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesY_F, WxLnExoneracionesY)
        FormBoleTicket.Sections("PieInforme").Controls("TotalEnLetras").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasX_F, WxLnTotalEnLetrasX)
        FormBoleTicket.Sections("PieInforme").Controls("TotalEnLetras").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasY_F, WxLnTotalEnLetrasY)
        FormBoleTicket.Sections("PieInforme").Controls("TotalBoleta").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalX_F, WxLnTotalX)
        FormBoleTicket.Sections("PieInforme").Controls("TotalBoleta").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalY_F, WxLnTotalY)
        FormBoleTicket.Sections("PieInforme").Controls("SerieImpresora").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnSerieImpresoraX_F, WxLnSerieImpresoraX)
        FormBoleTicket.Sections("PieInforme").Controls("SerieImpresora").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnSerieImpresoraY_F, WxLnSerieImpresoraY)
        FormBoleTicket.Sections("PieInforme").Controls("Terminal").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTerminalX_F, WxLnTerminalX)
        FormBoleTicket.Sections("PieInforme").Controls("Terminal").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTerminalY_F, WxLnTerminalY)
        FormBoleTicket.Sections("cabecera").Height = IIf(lbEsBoletaDeFarmacia = True, WxLnCabeceraAlto_F, WxLnCabeceraAlto)
        FormBoleTicket.Sections("PieInforme").Height = IIf(lbEsBoletaDeFarmacia = True, WxLnPieAlto_F, WxLnPieAlto)
        '
        FormBoleTicket.Orientation = rptOrientDefault
        
        '
        If lbMuestraVistaPrevia = True Then
           FormBoleTicket.Show 1
        Else
           On Error Resume Next
           FormBoleTicket.PrintReport
        End If
    
    Else
        'mgaray
        Dim reporteadorAUsar  As Integer
        reporteadorAUsar = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoReporteador_F, WxLnTipoReporteador)
        If reporteadorAUsar = reporteadorAUsarEnum.DataReport Then
            Call imprimirBoletaEnDataReport(rsReporte, lcCajeroIniciales, lcNombreServicio, _
                                  lnTotal, lnImporteEXO, lcCuenta, lcAdelantos, _
                                  lbMuestraVistaPrevia, lcBoletaNumero, _
                                  lbEsBoletaDeFarmacia, lbElDocumentoEsTicket, _
                                  lbElDocumentoEsFactura, oConexion)

        Else
            Call imprimirBoletaEnCrytalReport(rsReporte, lcCajeroIniciales, lcNombreServicio, _
                                  lnTotal, lnImporteEXO, lcCuenta, lcAdelantos, _
                                  lbMuestraVistaPrevia, lcBoletaNumero, _
                                  lbEsBoletaDeFarmacia, lbElDocumentoEsTicket, _
                                  lbElDocumentoEsFactura, oConexion)
        End If
    End If
    If lbEsBoletaDeFarmacia = True Then
        'SeteaOtraImpresoraDefault "Epson FX-800"
        SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True, oConexion)
    End If
End Sub

'====================================================================================
'mgaray


'Posicion inicial de los controles de la cabecera y pie de pagina
'Por defecto los establecemos a la posicion 0,0 porque
'si en algun comprobante no se muestra un valor y no se le da los valor x/y
'este control se queda en la posición que se le dio en el diseño inicial
'Esto ocasionaria que el alto de la sección en la que se encuentre estara limitada a su posicion
'Y + su alto
Private Sub posicionInicialCabeceraPiePaginaBoleta(ByRef FormBoleta As DataReport)
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Top = 0
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Left = 0
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Top = 0
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Left = 0
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Top = 0
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Left = 0
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Top = 0
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Left = 0
    FormBoleta.Sections("cabecera").Controls("FechaCobranza").Top = 0
    FormBoleta.Sections("cabecera").Controls("FechaCobranza").Left = 0
    FormBoleta.Sections("cabecera").Controls("Servicio").Top = 0
    FormBoleta.Sections("cabecera").Controls("Servicio").Left = 0
    FormBoleta.Sections("cabecera").Controls("Observaciones").Top = 0
    FormBoleta.Sections("cabecera").Controls("Observaciones").Left = 0
        
        
    FormBoleta.Sections("Pie").Controls("Cajero").Top = 0
    FormBoleta.Sections("Pie").Controls("Cajero").Left = 0
    FormBoleta.Sections("Pie").Controls("nombreCaja").Top = 0
    FormBoleta.Sections("Pie").Controls("nombreCaja").Left = 0
    FormBoleta.Sections("Pie").Controls("Adelantos").Top = 0
    FormBoleta.Sections("Pie").Controls("Adelantos").Left = 0
    FormBoleta.Sections("Pie").Controls("TotalPorPagar").Top = 0
    FormBoleta.Sections("Pie").Controls("TotalPorPagar").Left = 0
    FormBoleta.Sections("Pie").Controls("idCuentaAtencion").Top = 0
    FormBoleta.Sections("Pie").Controls("idCuentaAtencion").Left = 0
    FormBoleta.Sections("Pie").Controls("Exoneraciones").Top = 0
    FormBoleta.Sections("Pie").Controls("Exoneraciones").Left = 0
    FormBoleta.Sections("Pie").Controls("TotalEnLetras").Top = 0
    FormBoleta.Sections("Pie").Controls("TotalEnLetras").Left = 0
    FormBoleta.Sections("Pie").Controls("TotalBoleta").Top = 0
    FormBoleta.Sections("Pie").Controls("TotalBoleta").Left = 0
    
    
    FormBoleta.Sections("Pie").Controls("SubTotal").Top = 0
    FormBoleta.Sections("Pie").Controls("SubTotal").Left = 0
    FormBoleta.Sections("Pie").Controls("IGV").Top = 0
    FormBoleta.Sections("Pie").Controls("IGV").Left = 0
    
End Sub

Private Sub posicionInicialCabeceraPiePaginaBoletaCrystalReport(ByRef crReport As CRAXDRT.Report)
    Dim oControls As CRAXDRT.ReportObjects
    Dim oControl As CRAXDRT.FieldObject
    Dim oSections As CRAXDRT.Sections
    Dim oSection As CRAXDRT.Section
    
    Set oSections = crReport.Sections
    For Each oSection In oSections
        Set oControls = oSection.ReportObjects
       
        For Each oControl In oControls
            oControl.Left = 0
            oControl.Top = 0
        Next
    Next
End Sub

Private Sub imprimirBoletaEnDataReport(rsReporte As Recordset, lcCajeroIniciales As String, lcNombreServicio As String, _
                                  lnTotal As Double, lnImporteEXO As Double, lcCuenta As String, lcAdelantos As String, _
                                  lbMuestraVistaPrevia As Boolean, lcBoletaNumero As String, _
                                  lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsTicket As Boolean, _
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection)
    '************************** Boleta/Recibo/Factura ***********************************************
    Set FormBoleta.DataSource = rsReporte

    FormBoleta.Caption = lcTipoComprobante
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Caption = lcBoletaNumero
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Caption = lcBoletaEstado
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Caption = lcTipoPago
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Caption = lcRazonSocialHC
    FormBoleta.Sections("cabecera").Controls("FechaCobranza").Caption = Format(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
    FormBoleta.Sections("cabecera").Controls("Servicio").Caption = lcNombreServicio
    FormBoleta.Sections("cabecera").Controls("Observaciones").Caption = IIf(IsNull(rsReporte.Fields!Observaciones), "", rsReporte.Fields!Observaciones)
    
    FormBoleta.Sections("pie").Controls("Cajero").Caption = lcCajeroIniciales
    FormBoleta.Sections("pie").Controls("nombreCaja").Caption = IIf(IsNull(rsReporte.Fields!nombreCaja), "", rsReporte.Fields!nombreCaja)
    FormBoleta.Sections("pie").Controls("TotalPorPagar").Caption = Format(lnTotal + lnImporteEXO, "######.#0")
    FormBoleta.Sections("pie").Controls("idCuentaAtencion").Caption = lcCuenta
    FormBoleta.Sections("pie").Controls("Adelantos").Caption = lcAdelantos
    FormBoleta.Sections("pie").Controls("Exoneraciones").Caption = Format(lnImporteEXO, "######.#0")
    FormBoleta.Sections("pie").Controls("TotalEnLetras").Caption = lcTotalEnLetras
    FormBoleta.Sections("pie").Controls("TotalEnLetras").Width = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalLetrasWidhtY_F, WxLnTotalLetrasWidhtY)
    FormBoleta.Sections("pie").Controls("TotalBoleta").Caption = Format(lnTotal, "######.#0")
    
    If lbElDocumentoEsFactura Then
        FormBoleta.Sections("pie").Controls("SubTotal").Caption = Format(100 * lnTotal / (100 + Val(wxParametro221)), "######.#0")
        FormBoleta.Sections("pie").Controls("IGV").Caption = Format(wxParametro221 * lnTotal / (100 + Val(wxParametro221)), "######.#0")
    Else
        FormBoleta.Sections("pie").Controls("SubTotal").Caption = ""
        FormBoleta.Sections("pie").Controls("IGV").Caption = ""
    End If
    
    FormBoleta.RightMargin = getReportRightMargin
    FormBoleta.TopMargin = getReportTopMargin
    FormBoleta.LeftMargin = getReportLeftMargin
    FormBoleta.BottomMargin = getReportBottomMargin
    
    'mgaray
    posicionInicialCabeceraPiePaginaBoleta FormBoleta
    
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieX_F, WxLnNumeroSerieX)
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieY_F, WxLnNumeroSerieY)
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoX_F, WxLnEstadoX)
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoY_F, WxLnEstadoY)
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoX_F, WxLnTipoX)
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoY_F, WxLnTipoY)
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialX_F, WxLnRzSocialX)
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialY_F, WxLnRzSocialY)
    FormBoleta.Sections("cabecera").Controls("FechaCobranza").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaX_F, WxLnFechaX)
    FormBoleta.Sections("cabecera").Controls("FechaCobranza").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaY_F, WxLnFechaY)
    FormBoleta.Sections("cabecera").Controls("Servicio").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioX_F, WxLnServicioX)
    FormBoleta.Sections("cabecera").Controls("Servicio").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioY_F, WxLnServicioY)
    FormBoleta.Sections("cabecera").Controls("Observaciones").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesX_F, WxLnObservacionesX)
    FormBoleta.Sections("cabecera").Controls("Observaciones").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesY_F, WxLnObservacionesY)
    
    FormBoleta.Sections("Detalle").Controls("txtCodigo").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCodigoY_F, WxLnCodigoY)
    FormBoleta.Sections("Detalle").Controls("txtNombreProducto").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoY_F, WxLnProductoY)
    FormBoleta.Sections("Detalle").Controls("txtNombreProducto").Width = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoWidhtY_F, WxLnProductoWidhtY)
    FormBoleta.Sections("Detalle").Controls("txtCantidad").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCantidadY_F, WxLnCantidadY)
    FormBoleta.Sections("Detalle").Controls("txtPrecioUnitario").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnPrecioY_F, WxLnPrecioY)
    FormBoleta.Sections("Detalle").Controls("txtTotalPorPagar").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnImporteY_F, WxLnImporteY)
    
    FormBoleta.Sections("Pie").Controls("Cajero").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroX_F, WxLnCajeroX)
    FormBoleta.Sections("Pie").Controls("Cajero").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroY_F, WxLnCajeroY)
    FormBoleta.Sections("Pie").Controls("nombreCaja").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaX_F, WxLnCajaX)
    FormBoleta.Sections("Pie").Controls("nombreCaja").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaY_F, WxLnCajaY)
    FormBoleta.Sections("Pie").Controls("Adelantos").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosX_F, WxLnAdelantosX)
    FormBoleta.Sections("Pie").Controls("Adelantos").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosY_F, WxLnAdelantosY)
    FormBoleta.Sections("Pie").Controls("TotalPorPagar").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarX_F, WxLnTotalPagarX)
    FormBoleta.Sections("Pie").Controls("TotalPorPagar").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarY_F, WxLnTotalPagarY)
    FormBoleta.Sections("Pie").Controls("idCuentaAtencion").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaX_F, WxLnCuentaX)
    FormBoleta.Sections("Pie").Controls("idCuentaAtencion").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaY_F, WxLnCuentaY)
    FormBoleta.Sections("Pie").Controls("Exoneraciones").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesX_F, WxLnExoneracionesX)
    FormBoleta.Sections("Pie").Controls("Exoneraciones").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesY_F, WxLnExoneracionesY)
    FormBoleta.Sections("Pie").Controls("TotalEnLetras").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasX_F, WxLnTotalEnLetrasX)
    FormBoleta.Sections("Pie").Controls("TotalEnLetras").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasY_F, WxLnTotalEnLetrasY)
    FormBoleta.Sections("Pie").Controls("TotalBoleta").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalX_F, WxLnTotalX)
    FormBoleta.Sections("Pie").Controls("TotalBoleta").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalY_F, WxLnTotalY)
    
    If lbElDocumentoEsFactura Then
        FormBoleta.Sections("Pie").Controls("SubTotal").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnSubTotalX_F, WxLnSubTotalX)
        FormBoleta.Sections("Pie").Controls("SubTotal").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnSubTotalY_F, WxLnSubTotalY)
        FormBoleta.Sections("Pie").Controls("IGV").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnIGVX_F, WxLnIGVX)
        FormBoleta.Sections("Pie").Controls("IGV").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnIGVY_F, WxLnIGVY)
    End If
    
    FormBoleta.Sections("cabecera").Height = IIf(lbEsBoletaDeFarmacia = True, WxLnCabeceraAlto_F, WxLnCabeceraAlto)
    FormBoleta.Sections("Pie").Height = IIf(lbEsBoletaDeFarmacia = True, WxLnPieAlto_F, WxLnPieAlto)
    '
    FormBoleta.Orientation = rptOrientDefault
    '
    If lbMuestraVistaPrevia = True Then
       FormBoleta.Show 1
    Else
       On Error Resume Next
       FormBoleta.PrintReport
    End If
End Sub


Private Sub imprimirBoletaEnCrytalReport(rsReporte As Recordset, lcCajeroIniciales As String, lcNombreServicio As String, _
                                  lnTotal As Double, lnImporteEXO As Double, lcCuenta As String, lcAdelantos As String, _
                                  lbMuestraVistaPrevia As Boolean, lcBoletaNumero As String, _
                                  lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsTicket As Boolean, _
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection)
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    'Dim rs As New Recordset
    
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition
    Dim rutaRaiz As String
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "smigracion_datos") > 0 Then
        rutaRaiz = Mid(rutaRaiz, 1, InStr(rutaRaiz, "smigracion_datos") - 2) & "/SIGH/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/rptFormBoleta.rpt", 1)
    
'    crReport.DiscardSavedData
    crReport.Database.SetDataSource crearRecorsetTemporal(rsReporte)
    rsReporte.MoveFirst
'    FormBoleta.Caption = lcTipoComprobante
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "BoletaNumeroSerie":
                crParameter.AddCurrentValue (lcBoletaNumero)
            Case "BoletaEstado":
                crParameter.AddCurrentValue (lcBoletaEstado)
            Case "BoletaTipo":
                crParameter.AddCurrentValue (lcTipoPago)
            Case "RazonSocial":
                crParameter.AddCurrentValue (lcRazonSocialHC)
            Case "FechaCobranza":
                crParameter.AddCurrentValue (CStr(Format(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY_HM)))
            Case "Servicio":
                crParameter.AddCurrentValue (lcNombreServicio)
            Case "Observaciones":
                crParameter.AddCurrentValue (CStr(IIf(IsNull(rsReporte.Fields!Observaciones), "", rsReporte.Fields!Observaciones)))
            Case "Cajero":
                crParameter.AddCurrentValue (lcCajeroIniciales)
            Case "nombreCaja":
                crParameter.AddCurrentValue (CStr(IIf(IsNull(rsReporte.Fields!nombreCaja), "", rsReporte.Fields!nombreCaja)))
            Case "Adelantos":
                If lcAdelantos <> "" Then
                    crParameter.AddCurrentValue (Format(lcAdelantos, "######.#0"))
                Else
                    crParameter.AddCurrentValue ("")
                End If
            Case "TotalBoletaPorPagar":
                crParameter.AddCurrentValue (Format(lnTotal + lnImporteEXO, "######.#0"))
            Case "idCuentaAtencion":
                crParameter.AddCurrentValue (lcCuenta)
            Case "Exoneraciones":
                crParameter.AddCurrentValue (Format(lnImporteEXO, "######.#0"))
            Case "TotalEnLetras":
                crParameter.AddCurrentValue (lcTotalEnLetras)
            Case "TotalBoleta":
                crParameter.AddCurrentValue (Format(lnTotal, "######.#0"))
            Case "SubTotal":
                If lbElDocumentoEsFactura Then
                    crParameter.AddCurrentValue (Format(100 * lnTotal / (100 + Val(wxParametro221)), "######.#0"))
                Else
                    crParameter.AddCurrentValue ("")
                End If
            Case "IGV":
                If lbElDocumentoEsFactura Then
                    crParameter.AddCurrentValue (Format(wxParametro221 * lnTotal / (100 + Val(wxParametro221)), "######.#0"))
                Else
                    crParameter.AddCurrentValue ("")
                End If
        End Select
    Next
    
    crReport.RightMargin = getReportRightMargin
    crReport.TopMargin = getReportTopMargin
    crReport.LeftMargin = getReportLeftMargin
    crReport.BottomMargin = getReportBottomMargin
    
    Call posicionarControlesBoletaCrystalReport(crReport, lbEsBoletaDeFarmacia, lbElDocumentoEsFactura)
    'cabecera
    crReport.Sections(2).Height = IIf(lbEsBoletaDeFarmacia = True, WxLnCabeceraAlto_F, WxLnCabeceraAlto)
    'pie
    crReport.Sections(5).Height = IIf(lbEsBoletaDeFarmacia = True, WxLnPieAlto_F, WxLnPieAlto)
    
    crReport.PaperOrientation = crDefaultPaperOrientation
        
    
    If lbMuestraVistaPrevia = True Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = lcTipoComprobante
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
    Else
       On Error Resume Next
       crReport.PrintOut False
    End If
End Sub

Private Sub posicionarControlesBoletaCrystalReport(ByRef crReport As CRAXDRT.Report, _
            lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsFactura As Boolean)
    Dim arrayNombresCampos(21) As String
    Dim nombreCampo As String
    
    obtenerNombresCampos arrayNombresCampos
    
    posicionInicialCabeceraPiePaginaBoletaCrystalReport crReport
    
    Dim oControls As CRAXDRT.ReportObjects
    Dim oControl As CRAXDRT.FieldObject
    
    Dim oSections As CRAXDRT.Sections
    Dim oSection As CRAXDRT.Section
    
    Set oSections = crReport.Sections
    
        
    For Each oSection In oSections
        Set oControls = oSection.ReportObjects
        For Each oControl In oControls
            nombreCampo = UCase(leerNombreCampo(oControl.Field.Name, arrayNombresCampos))
            Select Case nombreCampo
                Case UCase("BoletaNumeroSerie"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieX_F, WxLnNumeroSerieX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnNumeroSerieY_F, WxLnNumeroSerieY)
                Case UCase("BoletaEstado"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoX_F, WxLnEstadoX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnEstadoY_F, WxLnEstadoY)
                Case UCase("BoletaTipo"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoX_F, WxLnTipoX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTipoY_F, WxLnTipoY)
                Case UCase("RazonSocial"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialX_F, WxLnRzSocialX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnRzSocialY_F, WxLnRzSocialY)
                Case UCase("FechaCobranza"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaX_F, WxLnFechaX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnFechaY_F, WxLnFechaY)
                Case UCase("Servicio"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioX_F, WxLnServicioX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnServicioY_F, WxLnServicioY)
                Case UCase("Observaciones"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesX_F, WxLnObservacionesX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnObservacionesY_F, WxLnObservacionesY)
                Case UCase("Cajero"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroX_F, WxLnCajeroX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajeroY_F, WxLnCajeroY)
                Case UCase("nombreCaja"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaX_F, WxLnCajaX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCajaY_F, WxLnCajaY)
                Case UCase("Adelantos"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosX_F, WxLnAdelantosX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnAdelantosY_F, WxLnAdelantosY)
                Case UCase("TotalBoletaPorPagar"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarX_F, WxLnTotalPagarX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalPagarY_F, WxLnTotalPagarY)
                Case UCase("idCuentaAtencion"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaX_F, WxLnCuentaX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCuentaY_F, WxLnCuentaY)
                Case UCase("Exoneraciones"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesX_F, WxLnExoneracionesX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnExoneracionesY_F, WxLnExoneracionesY)
                Case UCase("TotalEnLetras"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasX_F, WxLnTotalEnLetrasX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalEnLetrasY_F, WxLnTotalEnLetrasY)
                    oControl.Width = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalLetrasWidhtY_F, WxLnTotalLetrasWidhtY)
                Case UCase("TotalBoleta"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalX_F, WxLnTotalX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnTotalY_F, WxLnTotalY)
                Case UCase("SubTotal"):
                    If lbElDocumentoEsFactura Then
                        oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnSubTotalX_F, WxLnSubTotalX)
                        oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnSubTotalY_F, WxLnSubTotalY)
                    End If
                Case UCase("IGV"):
                    If lbElDocumentoEsFactura Then
                        oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnIGVX_F, WxLnIGVX)
                        oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnIGVY_F, WxLnIGVY)
                    End If
                Case UCase("Codigo"):
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCodigoY_F, WxLnCodigoY)
                Case UCase("NombreProducto"):
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoY_F, WxLnProductoY)
                    oControl.Width = IIf(lbEsBoletaDeFarmacia = True, WxLnProductoWidhtY_F, WxLnProductoWidhtY)
                Case UCase("Cantidad"):
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCantidadY_F, WxLnCantidadY)
                Case UCase("PrecioUnitario"):
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnPrecioY_F, WxLnPrecioY)
                Case UCase("TotalPorPagar"):
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnImporteY_F, WxLnImporteY)
            End Select
        Next
    Next
End Sub

Public Function leerNombreCampo(cadenaBusqueda As String, arrayNombreCampo() As String)
    Dim cadenaABuscar As Variant
    Dim nombreCampo As String
    Dim posicionEncontro As Integer
    
    For Each cadenaABuscar In arrayNombreCampo
        If cadenaABuscar <> "" Then
            posicionEncontro = InStr(UCase(cadenaBusqueda), UCase(cadenaABuscar))
            If posicionEncontro > 0 Then
                nombreCampo = cadenaABuscar
                Exit For
            End If
        End If
    Next
    leerNombreCampo = nombreCampo
End Function

Public Function obtenerNombresCampos(nombresCampos() As String)
    nombresCampos(0) = "BoletaNumeroSerie"
    nombresCampos(1) = "BoletaEstado"
    nombresCampos(2) = "BoletaTipo"
    nombresCampos(3) = "RazonSocial"
    nombresCampos(4) = "FechaCobranza"
    nombresCampos(5) = "Servicio"
    nombresCampos(6) = "Observaciones"
    nombresCampos(7) = "Cajero"
    nombresCampos(8) = "nombreCaja"
    nombresCampos(9) = "Adelantos"
    nombresCampos(10) = "TotalBoletaPorPagar"
    nombresCampos(11) = "idCuentaAtencion"
    nombresCampos(12) = "Exoneraciones"
    nombresCampos(13) = "TotalEnLetras"
    nombresCampos(14) = "TotalBoleta"
    nombresCampos(15) = "SubTotal"
    nombresCampos(16) = "IGV"
    
    nombresCampos(17) = "Codigo"
    nombresCampos(18) = "NombreProducto"
    nombresCampos(19) = "Cantidad"
    nombresCampos(20) = "PrecioUnitario"
    nombresCampos(21) = "TotalPorPagar"
    
    obtenerNombresCampos = nombresCampos
End Function

Public Function getDetailMinimunHeight(tipoReporteador As reporteadorAUsarEnum) As Integer
    Select Case tipoReporteador
        Case reporteadorAUsarEnum.DataReport
            getDetailMinimunHeight = 2883 '3016 '3383 '2883
        Case reporteadorAUsarEnum.CrystalReport
            getDetailMinimunHeight = 315
        Case Else
            getDetailMinimunHeight = 315
    End Select
End Function

Public Function getReportTopMargin() As Integer
    getReportTopMargin = WxLnMargenSuperiorY
End Function


Public Function getReportBottomMargin() As Integer
    getReportBottomMargin = WxLnMargenInferiorY
End Function

Public Function getReportRightMargin() As Integer
    getReportRightMargin = WxLnMargenDerechoX
End Function


Public Function getReportLeftMargin() As Integer
    getReportLeftMargin = WxLnMargenIzquierdoX
End Function

'Alto minimo que se le le puede dar a la seccion de pie de pagina
'Esta Determinada por control que tiene el mayor alto configurado(Para el caso de boleta TotalEnLetras tiene un alto de 510 twip)
Public Function getFooterBoletaHeightMinimun(tipoReporteador As reporteadorAUsarEnum) As Integer
    Select Case tipoReporteador
        Case reporteadorAUsarEnum.DataReport
            getFooterBoletaHeightMinimun = 510
        Case reporteadorAUsarEnum.CrystalReport
            getFooterBoletaHeightMinimun = 315
    End Select
End Function

'Alto minimo que se le le puede dar a la seccion de cabecera de pagina
'Esta Determinada por control que tiene el mayor alto configurado(Para el caso de boleta todos los controles tiene el mismo alto: 315)
Public Function getHeaderBoletaHeightMinimun(tipoReporteador As reporteadorAUsarEnum) As Integer
    Select Case tipoReporteador
        Case reporteadorAUsarEnum.DataReport
            getHeaderBoletaHeightMinimun = 315
        Case reporteadorAUsarEnum.CrystalReport
            getHeaderBoletaHeightMinimun = 315
    End Select
End Function

'Metodo solo llamado al usar el reporte crystal report
Private Function crearRecorsetTemporal(rsReporte As Recordset) As Recordset
    Dim rs As New Recordset
    
    With rs
        .Fields.Append "codigo", adVarChar, 20
        .Fields.Append "NombreProducto", adVarChar, 250
        .Fields.Append "cantidad", adInteger
        .Fields.Append "precioUnitario", adCurrency
        .Fields.Append "totalPorPagar", adCurrency
        .LockType = adLockOptimistic
        .Open
        If Not (rsReporte.BOF And rsReporte.EOF) Then
            rsReporte.MoveFirst
            While rsReporte.EOF = False
                .AddNew
                .Fields!Codigo = rsReporte!Codigo
                .Fields!NombreProducto = rsReporte!NombreProducto
                .Fields!Cantidad = rsReporte!Cantidad
                .Fields!PrecioUnitario = rsReporte!PrecioUnitario
                .Fields!TotalPorPagar = rsReporte!TotalPorPagar
                .Update
                rsReporte.MoveNext
            Wend
        End If
   End With
   Set crearRecorsetTemporal = rs
End Function


