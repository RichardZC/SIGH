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
Dim mo_Procesos As New SIGHProxies.Procesos
Dim mo_ReporteUtil As New ReporteUtil
'mgaray
Dim lcBoletaTipo As String, lcBoletaEstado As String, lcRazonSocialHC As String
Dim lcTotalEnLetras As String, lnIdCaja As Long, lcObservaciones As String
Dim lcTipoPago As String, lcTipoComprobante As String

Enum reporteadorAUsarEnum
    DataReport = 0
    CrystalReport = 1
End Enum


Sub ImprimeBoletaEnFormatoUsuario(lbEsTicketFACTURA As Boolean, lcRucDireccionProveedor As String, lcVuelto As String, _
                                  rsReporte As Recordset, lcCajeroIniciales As String, lcNombreServicio As String, _
                                  lnTotal As Double, lnImporteEXO As Double, lcCuenta As String, lcAdelantos As String, _
                                  lcSerieImpresoraDefault As String, lcSerieImpresora2 As String, lbMuestraVistaPrevia As Boolean, lcBoletaNumero As String, _
                                  lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsTicket As Boolean, _
                                  lnSubTotal As Double, lnIGV As Double, _
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection, lbEsTicketSullana As Boolean, _
                                  Optional lbEsCopiaDeBoleta As Boolean)
                                  
                                  
'    Dim lcBoletaTipo As String, lcBoletaEstado As String, lcRazonSocialHC As String
'    Dim lcTotalEnLetras As String, lnIdCaja As Long, lcObservaciones As String
'    Dim lcTipoPago As String, lcTipoComprobante As String
    Dim lnOperGrav As Double
    Dim lcQr As String, lcServicioNombre As String, lcObservacionesHRC As String
    Dim lcServicioInafecto As String, lbIncluyeIGV As Boolean, lcTexto123 As String
    '
    lcServicioInafecto = lcBuscaParametro.SeleccionaFilaParametro(537)
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
       lcRazonSocialHC = lcRazonSocialHC & " (HC " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte.Fields!NroHistoriaClinica)), False)
    End If
    On Error Resume Next
    If rsReporte!IdTipoComprobante <> 2 Then   'solo boleta lo recorta
       lcRazonSocialHC = Left(lcRazonSocialHC + String(45, " "), 40)
    End If
    '
    lcTotalEnLetras = "Son: " & sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighentidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles"
    '
    lnIdCaja = rsReporte.Fields!IdCaja
    If lbElDocumentoEsTicket = True Then
        SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True, oConexion)
    Else
        If lbEsBoletaDeFarmacia = True Then
                SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, False, oConexion)
        Else
                SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True, oConexion)
        End If
    End If
    If lbElDocumentoEsTicket = True Then
            '************************** ticket BOLETA***********************************************
            lcServicioNombre = "Servicio: " & lcNombreServicio
            If Not IsNull(rsReporte!DNI) Then
               If rsReporte!DNI <> "" Then
                  lcServicioNombre = "Dni: " & rsReporte!DNI & "  " & lcServicioNombre
               End If
            End If
            
            lcQr = mo_Procesos.DevuelveTextoParaQRsullana(rsReporte)
            Set FormBoleSull.DataSource = rsReporte
            If Len(wxParametro205) > 50 Then
               FormBoleSull.Sections("cabecera").Controls("lblHospital").Font.Size = 8
            End If
            FormBoleSull.Sections("cabecera").Controls("lblHospital").Caption = UCase(wxParametro205)
            FormBoleSull.Sections("cabecera").Controls("lblRUC").Caption = UCase("RUC: ") & (wxParametro339)
            FormBoleSull.Sections("cabecera").Controls("lblDireccion").Caption = mo_Teclado.CapitalizarNombres(wxParametro206)
            If rsReporte!IdTipoComprobante <> 2 Then
                FormBoleSull.Sections("cabecera").Controls("lblBoletaT").Caption = "BOLETA DE VENTA ELECTRONICA"
                FormBoleSull.Sections("cabecera").Controls("lbPaciente").Caption = "PACIENTE: " & UCase(Left(Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial)) + String(33, " "), 33))
                If wxHuboCambioAfactura = True Then
                    FormBoleSull.Sections("cabecera").Controls("etServicio").Top = FormBoleSull.Sections("cabecera").Controls("lbPaciente").Top + 255
                    FormBoleSull.Sections("cabecera").Controls("lbRucDireccion").Top = FormBoleSull.Sections("cabecera").Controls("etServicio").Top
                    FormBoleSull.Sections("cabecera").Controls("lbRucDireccion").Caption = ""
                    FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Top = FormBoleSull.Sections("cabecera").Controls("etServicio").Top + 255
                End If
            Else
                FormBoleSull.Sections("cabecera").Controls("lblBoletaT").Caption = "FACTURA ELECTRONICA"
                FormBoleSull.Sections("cabecera").Height = FormBoleSull.Sections("cabecera").Height + (225 * 4) + (25 * 4)
                FormBoleSull.Sections("cabecera").Controls("lbRucDireccion").Top = FormBoleSull.Sections("cabecera").Controls("etServicio").Top
                FormBoleSull.Sections("cabecera").Controls("etServicio").Top = FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Top + 225
                FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Top = FormBoleSull.Sections("cabecera").Controls("Etiqueta2").Top + 225
                FormBoleSull.Sections("cabecera").Controls("Línea1").Top = FormBoleSull.Sections("cabecera").Controls("Línea1").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("Etiqueta2").Top = FormBoleSull.Sections("cabecera").Controls("Etiqueta2").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("Etiqueta3").Top = FormBoleSull.Sections("cabecera").Controls("Etiqueta3").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("Etiqueta5").Top = FormBoleSull.Sections("cabecera").Controls("Etiqueta5").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("Etiqueta6").Top = FormBoleSull.Sections("cabecera").Controls("Etiqueta6").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("Línea2").Top = FormBoleSull.Sections("cabecera").Controls("Línea2").Top + 225 + 225
                FormBoleSull.Sections("cabecera").Controls("lbRucDireccion").Caption = lcRucDireccionProveedor
                FormBoleSull.Sections("cabecera").Controls("lbPaciente").Caption = "Raz.Soc: " & UCase(Left(Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial)) + String(33, " "), 33))
                FormBoleSull.Sections("cabecera").Height = FormBoleSull.Sections("cabecera").Height - (225 * 4)
            End If
            FormBoleSull.Sections("cabecera").Controls("lblBoletaN").Caption = lcBoletaNumero
            FormBoleSull.Sections("cabecera").Controls("LbFecha").Caption = "FECHA: " & Format(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY)
            FormBoleSull.Sections("cabecera").Controls("lbHora").Caption = "HORA: " & Format$(rsReporte.Fields!FechaCobranza, sighentidades.DevuelveHoraSoloFormato_HMS)
            
            If IsNull(rsReporte!NroHistoriaClinica) Then
               FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Caption = lcCuenta
            Else
               FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Caption = lcCuenta & "  HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(rsReporte!NroHistoriaClinica, False)
            End If
            FormBoleSull.Sections("cabecera").Controls("etServicio").Caption = lcServicioNombre
            If lbEsBoletaDeFarmacia = False And (Val(sighentidades.Parametro568) = rsReporte!idProducto Or _
                                                            Val(sighentidades.Parametro569) = rsReporte!idProducto) Then
               lcObservacionesHRC = Trim(rsReporte!Observaciones)
               FormBoleSull.Sections("cabecera").Controls("etServicio").Caption = Left(lcObservacionesHRC, 53)
               FormBoleSull.Sections("cabecera").Controls("etCuentaOrdPago").Caption = IIf(Len(lcObservacionesHRC) > 53, Mid(lcObservacionesHRC, 54, 100), " ")
            End If
            'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
            If lbEsBoletaDeFarmacia = False Then
               '******************************************************servicios*********************************************
               If rsReporte!IdTipoComprobante = 2 Then
                   'Servicios-Factura
                   If rsReporte!idProducto = Val(wxParametro549) Then
                        '********Facturas con DESCRIPCIONES LARGAS - HRA************
                        lcServicioInafecto = ""
                        If rsReporte!IGV = 0 Then
                            lnTotal = rsReporte!Subtotal + rsReporte!IGV
                            FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + rsReporte!exoneraciones, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(0, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(0, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lblOpInafecta").Caption = Format(lnTotal, "#####0.#0")
                        Else
                            lnTotal = rsReporte!Total
                            lnIGV = Round(Val(lcBuscaParametro.SeleccionaFilaParametro(221)) * lnTotal / 100, 2)
                            FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(lnTotal - lnIGV, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(lnIGV, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lblOpInafecta").Caption = Format(0, "#####0.#0")
                        End If
                   Else
                        lbIncluyeIGV = False
                        If sighentidades.Parametro562 = "S" Then
                           If mo_ReglasCaja.LaFacturaEsUnReemboloso(rsReporte!IdComprobantePago) = True And rsReporte!IGV > 0 Then
                              lbIncluyeIGV = True
                           End If
                        End If
                        If lbIncluyeIGV = True Then
                            'Servicios-Factura-con IGV
                            lnTotal = rsReporte!Total
                            lnIGV = Round(Val(lcBuscaParametro.SeleccionaFilaParametro(221)) * lnTotal / 100, 2)
                            FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(lnTotal - lnIGV, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(lnIGV, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lblOpInafecta").Caption = Format(0, "#####0.#0")
                        Else
                            'Servicios-Factura-sin IGV
                            lnTotal = rsReporte!Subtotal + rsReporte!IGV
                            FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + rsReporte!exoneraciones, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(0, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(0, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                            FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                        End If
                   End If
               Else
                   'Servicios-Boleta
                   FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + lnImporteEXO, "#####0.#0")
                   FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(0, "#####0.#0")
                   FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(0, "#####0.#0")
                   FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                   FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
               End If
               If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") Then
                   FormBoleSull.Sections("PieInforme").Controls("lblOpInafecta").Caption = Format(lnTotal, "#####0.#0")
               End If
            Else
                '****************************************************farmacia**********************************************
                If rsReporte!IdTipoComprobante = 2 Then
                   'Farmacia-Factura
                   lbIncluyeIGV = False
                   If sighentidades.Parametro562 = "S" Then
                      If mo_ReglasCaja.LaFacturaEsUnReemboloso(rsReporte!IdComprobantePago) = True And rsReporte!IGV > 0 Then
                         lbIncluyeIGV = True
                      End If
                   Else
                      lbIncluyeIGV = True
                   End If
                   If lbIncluyeIGV = False Then
                      'Farmacia-Factura sin IGV
                      lnIGV = Round(Val(lcBuscaParametro.SeleccionaFilaParametro(221)) * lnTotal / 100, 2)
                      lnTotal = rsReporte!Subtotal
                      FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(0, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(0, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("etExoneracion").Caption = "Op.Gratuita S/"
                      FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal - lnImporteEXO, "#####0.#0")
                   Else
                      'Farmacia-Factura con IGV
                      lnTotal = rsReporte!Subtotal + rsReporte!IGV
                      lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
                      lnOperGrav = Round(lnTotal / ((100 + lnIGV) / 100), 2)
                      FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + rsReporte!exoneraciones, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(lnOperGrav, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(lnTotal - lnOperGrav, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                      FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                   End If
                Else
                   'Farmacia/boleta
                   If lnImporteEXO = 0 Then
                        lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
                        lnOperGrav = Round(lnTotal / ((100 + lnIGV) / 100), 2)
                        FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal + lnImporteEXO, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(lnOperGrav, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(lnTotal - lnOperGrav, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal, "#####0.#0")
                   Else
                        lnIGV = Round(Val(lcBuscaParametro.SeleccionaFilaParametro(221)) * lnTotal / 100, 2)
                        lnTotal = rsReporte!Subtotal
                        FormBoleSull.Sections("PieInforme").Controls("TotalPorPagar").Caption = Format(lnTotal, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("LbOpGravada").Caption = Format(0, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("lbIGV").Caption = Format(0, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("etExoneracion").Caption = "Op.Gratuita S/"
                        FormBoleSull.Sections("PieInforme").Controls("LbExoneracion").Caption = Format(lnImporteEXO, "#####0.#0")
                        FormBoleSull.Sections("PieInforme").Controls("lbTotal").Caption = Format(lnTotal - lnImporteEXO, "#####0.#0")
                        
                   End If
                End If
            End If
            If wxParametro548 <> "S" Then
               FormBoleSull.Sections("PieInforme").Controls("lblVuelto").Caption = lcVuelto
            End If
            
            If sighentidades.Parametro322valorInt <> "9" Then
               Set FormBoleSull.Sections("PieInforme").Controls("image1").Picture = sighentidades.GetPictureQrCode(lcQr, 100, 100)
            End If
            FormBoleSull.Sections("PieInforme").Controls("TotalEnLetras").Caption = lcTotalEnLetras
            'lcTexto123 = mo_ReglasCaja.DevuelveMovimientoEnLabImaFar(rsreporte!IdOrdenPago, rsreporte!IdTipoOrden, rsreporte!idCuentaAtencion)
            lcTexto123 = mo_ReporteUtil.NullToVacio(rsReporte!Observaciones)
            FormBoleSull.Sections("PieInforme").Controls("lbCaja").Caption = UCase("CAJA: " & _
                              IIf(IsNull(rsReporte.Fields!nombreCaja), "", rsReporte.Fields!nombreCaja)) & lcTexto123
            FormBoleSull.Sections("PieInforme").Controls("lbCajero").Caption = UCase("CAJERO : " & lcCajeroIniciales)
            FormBoleSull.Sections("PieInforme").Controls("lblSitioWeb").Caption = LCase(wxParametro538)
            FormBoleSull.RightMargin = 0  'getReportRightMargin
            FormBoleSull.TopMargin = 0 'getReportTopMargin
            FormBoleSull.LeftMargin = 0 'getReportLeftMargin
            FormBoleSull.BottomMargin = 0 ' getReportBottomMargin
            FormBoleSull.Orientation = rptOrientDefault

            If lbMuestraVistaPrevia = True Then
              FormBoleSull.Show 1
            Else
               On Error Resume Next
               FormBoleSull.PrintReport
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
                                  lbElDocumentoEsFactura, oConexion, lnSubTotal, lnIGV, lcRucDireccionProveedor)

        Else
            Call imprimirBoletaEnCrytalReport(rsReporte, lcCajeroIniciales, lcNombreServicio, _
                                  lnTotal, lnImporteEXO, lcCuenta, lcAdelantos, _
                                  lbMuestraVistaPrevia, lcBoletaNumero, _
                                  lbEsBoletaDeFarmacia, lbElDocumentoEsTicket, _
                                  lbElDocumentoEsFactura, oConexion, lnSubTotal, lnIGV, lcRucDireccionProveedor)
        End If
    End If
    If lbElDocumentoEsTicket = True Then
    Else
        If lbEsBoletaDeFarmacia = True Then
                SeteaOtraImpresoraDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True, oConexion)
        End If
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
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection, _
                                  lnSubTotal As Double, lnIGV As Double, _
                                  lcRucDireccionProveedor As String)
    '************************** Boleta/Recibo/Factura ***********************************************
    Dim lcFactRuc As String, lcFactDireccion As String
    SeparaRucYDireccion lcRucDireccionProveedor, lcFactRuc, lcFactDireccion
    
    
    Set FormBoleta.DataSource = rsReporte

    FormBoleta.Caption = lcTipoComprobante
    If lbElDocumentoEsFactura Then
        FormBoleta.Sections("cabecera").Controls("FactRUC").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCabRucX_F, WxLnCabRucX)
        FormBoleta.Sections("cabecera").Controls("FactRUC").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCabRucY_F, WxLnCabRucY)
        FormBoleta.Sections("cabecera").Controls("FactRUC").Caption = lcFactRuc
        FormBoleta.Sections("cabecera").Controls("etRucDireccion").Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCabDireccionX_F, WxLnCabDireccionX)
        FormBoleta.Sections("cabecera").Controls("etRucDireccion").Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCabDireccionY_F, WxLnCabDireccionY)
        FormBoleta.Sections("cabecera").Controls("etRucDireccion").Caption = lcFactDireccion
    Else
        FormBoleta.Sections("cabecera").Controls("FactRUC").Caption = ""
        FormBoleta.Sections("cabecera").Controls("etRucDireccion").Caption = ""
    End If
    
    FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Caption = lcBoletaNumero
    FormBoleta.Sections("cabecera").Controls("BoletaEstado").Caption = lcBoletaEstado
    FormBoleta.Sections("cabecera").Controls("BoletaTipo").Caption = lcTipoPago
    FormBoleta.Sections("cabecera").Controls("RazonSocial").Caption = lcRazonSocialHC
    If Val(wxParametro208) = 7637 Then   'huaral
        FormBoleta.Sections("cabecera").Controls("RazonSocial").Font.Size = 11
        FormBoleta.Sections("cabecera").Controls("FechaCobranza").Font.Size = 11
        FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Font.Size = 11
        FormBoleta.Sections("cabecera").Controls("lblHospital").Caption = "HOSPITAL HUARAL Y SBS <> RUC=" & lcBuscaParametro.SeleccionaFilaParametro(339)
        FormBoleta.Sections("cabecera").Controls("BoletaNumeroSerie").Caption = IIf(lbElDocumentoEsFactura = True, "FACTURA: ", _
                                                                                IIf(lbEsBoletaDeFarmacia = True, "BOLETA DE VENTA ELECTRONICA: ", _
                                                                                "RECIBO DE CAJA: ")) & lcBoletaNumero
    End If
    
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
        FormBoleta.Sections("pie").Controls("SubTotal").Caption = Format(lnSubTotal, "######.#0")
        FormBoleta.Sections("pie").Controls("IGV").Caption = Format(lnIGV, "######.#0")
    Else
        FormBoleta.Sections("pie").Controls("SubTotal").Caption = ""
        FormBoleta.Sections("pie").Controls("IGV").Caption = ""
    End If
    
    FormBoleta.RightMargin = 0  'getReportRightMargin
    FormBoleta.TopMargin = 0  'getReportTopMargin
    FormBoleta.LeftMargin = 0  'getReportLeftMargin
    FormBoleta.BottomMargin = 0  'getReportBottomMargin
    
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
    'FormBoleta.Sections("cabecera").Controls("etRucDireccion").Top = FormBoleta.Sections("cabecera").Controls("RazonSocial").Top + 200
    'FormBoleta.Sections("cabecera").Controls("etRucDireccion").Left = FormBoleta.Sections("cabecera").Controls("RazonSocial").Left
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

Sub SeparaRucYDireccion(lcRucDireccion As String, ByRef lcRuc As String, ByRef lcDireccion As String)
    Dim lnFor As Integer, lcRucN As String, lcDireccionN As String
    If lcRucDireccion = "" Then
        lcDireccion = ""
        lcRuc = ""
    Else
        lcRucN = "RUC: "
        lcRuc = ""
        For lnFor = InStr(lcRucDireccion, lcRucN) + Len(lcRucN) To InStr(lcRucDireccion, lcRucN) + Len(lcRucN) + 11
           lcRuc = lcRuc & Mid(lcRucDireccion, lnFor, 1)
        Next
        lcDireccionN = "DIRECCION: "
        lcDireccion = ""
        For lnFor = InStr(lcRucDireccion, lcDireccionN) + Len(lcDireccionN) To Len(lcRucDireccion)
           lcDireccion = lcDireccion & Mid(lcRucDireccion, lnFor, 1)
        Next
    End If
End Sub

Private Sub imprimirBoletaEnCrytalReport(rsReporte As Recordset, lcCajeroIniciales As String, lcNombreServicio As String, _
                                  lnTotal As Double, lnImporteEXO As Double, lcCuenta As String, lcAdelantos As String, _
                                  lbMuestraVistaPrevia As Boolean, lcBoletaNumero As String, _
                                  lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsTicket As Boolean, _
                                  lbElDocumentoEsFactura As Boolean, oConexion As Connection, _
                                  lnSubTotal As Double, lnIGV As Double, lcRucDireccionProveedor As String)
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    'Dim rs As New Recordset
    
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition
    Dim rutaRaiz As String, ldFechaCobranza9 As Date, lcObservaciones As String, lcNombreCaja As String
    Dim lcFactRuc As String, lcFactDireccion As String, lcXX As String
    SeparaRucYDireccion lcRucDireccionProveedor, lcFactRuc, lcFactDireccion
    On Error GoTo errCry
    
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
    
   ' crReport.DiscardSavedData
    rsReporte.MoveFirst
    ldFechaCobranza9 = rsReporte!FechaCobranza
    lcObservaciones = IIf(IsNull(rsReporte!Observaciones), "", rsReporte!Observaciones)
    lcNombreCaja = IIf(IsNull(rsReporte!nombreCaja), "", rsReporte!nombreCaja)
    crReport.Database.SetDataSource crearRecorsetTemporal(rsReporte)
    
'    FormBoleta.Caption = lcTipoComprobante
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "pHospital"
               If Val(wxParametro280) = 7637 Then
                  crParameter.AddCurrentValue ("HOSPITAL HUARAL Y SBS")
               Else
                  crParameter.AddCurrentValue (" ")
               End If
            Case "BoletaNumeroSerie":
                If Val(wxParametro280) = 7637 Then
                   crParameter.AddCurrentValue (IIf(lbElDocumentoEsFactura = True, "FACTURA: ", _
                                                IIf(lbEsBoletaDeFarmacia = True, "BOLETA DE VENTA ELECTRONICA: ", _
                                                "RECIBO DE CAJA: ")) & lcBoletaNumero)
                Else
                   crParameter.AddCurrentValue (lcBoletaNumero)
                End If
            Case "BoletaEstado":
                crParameter.AddCurrentValue (lcBoletaEstado)
            Case "BoletaTipo":
                crParameter.AddCurrentValue (lcTipoPago)
            Case "RazonSocial":
                crParameter.AddCurrentValue (lcRazonSocialHC)
            Case "ruc"
                crParameter.AddCurrentValue (lcFactRuc)
            Case "direccion"
                crParameter.AddCurrentValue (lcFactDireccion)
            Case "FechaCobranza":
                crParameter.AddCurrentValue (Format(ldFechaCobranza9, sighentidades.DevuelveFechaSoloFormato_DMY_HM))
            Case "Servicio":
                crParameter.AddCurrentValue (lcNombreServicio)
            Case "Observaciones":
                crParameter.AddCurrentValue (lcObservaciones)
            Case "Cajero":
                crParameter.AddCurrentValue (lcCajeroIniciales)
            Case "nombreCaja":
                crParameter.AddCurrentValue (lcNombreCaja)
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
                    crParameter.AddCurrentValue (Format(lnSubTotal, "######.#0"))
                Else
                    crParameter.AddCurrentValue ("")
                End If
            Case "IGV":
                If lbElDocumentoEsFactura Then
                    crParameter.AddCurrentValue (Format(lnIGV, "######.#0"))
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
lcXX = ""
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = lcTipoComprobante
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
    Else
       On Error Resume Next
       crReport.PrintOut False
    End If
    Exit Sub
errCry:
    MsgBox Err.Description
    Exit Sub
   Resume
End Sub

Private Sub posicionarControlesBoletaCrystalReport(ByRef crReport As CRAXDRT.Report, _
            lbEsBoletaDeFarmacia As Boolean, lbElDocumentoEsFactura As Boolean)
    Dim arrayNombresCampos(23) As String
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
                Case UCase("DIRECCION"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCabDireccionX_F, WxLnCabDireccionX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCabDireccionY_F, WxLnCabDireccionY)
                Case UCase("RUC"):
                    oControl.Top = IIf(lbEsBoletaDeFarmacia = True, WxLnCabRucX_F, WxLnCabRucX)
                    oControl.Left = IIf(lbEsBoletaDeFarmacia = True, WxLnCabRucY_F, WxLnCabRucY)
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
    
    nombresCampos(22) = "Ruc"
    nombresCampos(23) = "Direccion"
    
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


