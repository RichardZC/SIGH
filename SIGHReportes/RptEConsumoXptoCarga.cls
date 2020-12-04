VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEConsumoXptoCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para consumo por punto de carga
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim rsTmpSOAT As New Recordset
Dim mrs_Tmp As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lbEsAdmisionCE As Boolean
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long)

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New EConsumoXPtoCarga
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
End Sub


Property Get Devuelve_mrs_Tmp() As Recordset
    Set Devuelve_mrs_Tmp = mrs_Tmp.Clone()
End Property
Property Get Devuelve_rsTmpSOAT() As Recordset
    Set Devuelve_rsTmpSOAT = rsTmpSOAT.Clone()
End Property


Sub ProcesaDatosYllenaTmp(lnIdResponsable As Long, mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPlan As Long, mb_EnResumen As Boolean, lbAgregarCuentasSinAlta As Boolean)
    On Error GoTo ErrProcesar
    Dim rsReporte As New ADODB.Recordset
    Dim rsTmp As New Recordset
    Dim rsTmp10 As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim oRsPuntosDeCarga As New Recordset
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
    Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
    Dim ldFechaPrincipio As Date
    Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim oConexion As New ADODB.Connection
    Dim lbPrimeraVez As Boolean
    Dim lcTexto1 As String: Dim lcTexto2 As String: Dim lcTexto3 As String
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim lnDebioPagar As Double, lnTotalPorCuenta As Double
    Dim lnIdTipoServicio As Long, lnIdDepartamento As Long, lnIdEspacialidad As Long, lnIdServicioPaciente As Long
    Dim lcDtipoServicio As String, lcDpto As String, lcDespecialidad As String, lcDServicio As String
    Dim lnSalieron As Integer, lnRetornaron As Integer, lnSalieronYnoRetornaron As Integer
    Dim lnTSalieron As Integer, lnTRetornaron As Integer, lnTSalieronYnoRetornaron As Integer
    Dim lcSql As String, lbEsCuentaPagante As Boolean
    Dim lbProcesaEnServidor As Boolean, lcHOraInicio As String, lcHoraFinal As String
    Dim ml_IdCuenta As Long, lnFor As Integer, f As Long, lRecordCount As Long
    Dim lcParametro245 As String

        lbProcesaEnServidor = True
        lcParametro245 = lcBuscaParametro.SeleccionaFilaParametro(245)
        
        lcHOraInicio = lcBuscaParametro.RetornaHoraServidorSQL1
        
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        
        '
        ReporteHRACreaTmp mb_EnResumen
        GenerarRecordsetTemporal
        If lnIdResponsable = 9 Then
           lnIdResponsable = 1
        End If
        Set oRsPuntosDeCarga = mo_reglasComunes.FactPuntosCargaDevuelveTodos
        '**********************************atenciones Hospitalizacion y Emergencia****************************
        For lnFor = 1 To 2
                If lnFor = 2 Then
                   'agrega cuentas con fecha de Ingreso y sin Fecha de Alta Medica
                   If lbAgregarCuentasSinAlta = False Then
                      Exit For
                   End If
                   If mb_EnResumen = True Then
                      Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaIngresoMedicoSinAlta(mda_FechaInicio, mda_FechaFin, False, oConexion, 0)
                   Else
                      Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaIngresoMedicoYfuenteFinanciamientoSinAlta(mda_FechaInicio, mda_FechaFin, ml_IdPlan, oConexion)
                      mrs_Tmp1.Filter = "idEstadoAtencion<>0 and idTipoServicio>1"
                   End If
                Else
                    'solo cuentas con Fecha de Alta Medica
                    If mb_EnResumen = True Then
                       Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaEgresoMedico(mda_FechaInicio, mda_FechaFin, False, oConexion, 0)
                    Else
                       Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaEgresoMedicoYfuenteFinanciamiento(mda_FechaInicio, mda_FechaFin, ml_IdPlan)
                       mrs_Tmp1.Filter = "idEstadoAtencion<>0 and idTipoServicio>1"
                    End If
                End If
                ml_IdCuenta = mrs_Tmp1.RecordCount
                
                If ml_IdCuenta > 0 Then
                   f = 0
                   lRecordCount = ml_IdCuenta
                   mrs_Tmp1.MoveFirst
                   Do While Not mrs_Tmp1.EOF
                        f = f + 1
                        RaiseEvent ProgressActualizaValor(f, lRecordCount)
                        '
                        lbEsAdmisionCE = False
                        
           
                        lnSaldoInicial = IIf(IsNull(mrs_Tmp1.Fields!NroHistoriaClinica), 0, mrs_Tmp1.Fields!NroHistoriaClinica)
                        ml_IdCuenta = mrs_Tmp1.Fields!idCuentaAtencion
                        lcTexto2 = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                        If Len(lcTexto2) > 30 Then
                           lcTexto2 = Left(lcTexto2, 30)
                        Else
                           lcTexto2 = Left(lcTexto2, 30) & String(30 - Len(lcTexto2), " ")
                        End If
                        lcTexto2 = lcTexto2 & " (Hc: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False) & ") (Cta: " & Format(ml_IdCuenta, "########") & ")"
                        lcTexto2 = lcTexto2 & IIf(mrs_Tmp1.Fields!idTipoServicio = 3, " (Ho)", " (Em)")
                        lcTexto2 = lcTexto2 & IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso), " (sin AM)", "")
                        lnTotalPorCuenta = 0
                        'servicios
                        lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!generaPago = 1, True, False)
                        If lbEsCuentaPagante = True Then
                            Set rsTmp = mo_ReglasFacturacion.FacturacionServicioPagosPorCuentaTodosConnection(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
                        Else
                            Set rsTmp = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuenta(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
                        End If
                        
                        If rsTmp.RecordCount > 0 Then
                           rsTmp.MoveFirst
                           Do While Not rsTmp.EOF
                              If rsTmp.Fields!TotalPorPagar > 0 And rsTmp.Fields!idProducto <> Val(lcParametro245) Then
                                lbPrimeraVez = True
                                If mrs_Tmp.RecordCount > 0 Then
                                   mrs_Tmp.MoveFirst
                                   Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga Then
                                         lbPrimeraVez = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                   Loop
                                End If
                                If lbPrimeraVez = True Then
                                      '
                                      oRsPuntosDeCarga.MoveFirst
                                      oRsPuntosDeCarga.Find "idPuntoCarga=" & rsTmp.Fields!IdPuntoCarga
                                      lcTexto1 = IIf(oRsPuntosDeCarga.EOF, "", Trim(oRsPuntosDeCarga.Fields!descripcion))
                                      '
                                      mrs_Tmp.AddNew
                                      mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                      mrs_Tmp.Fields!Paciente = lcTexto2
                                      mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                      mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga
                                      mrs_Tmp.Fields!dPuntoCarga = lcTexto1
                                      mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalPorPagar
                                      mrs_Tmp.Fields!fechaEgreso = IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso), mrs_Tmp1.Fields!FEchaIngreso, mrs_Tmp1.Fields!fechaEgreso)
                                      mrs_Tmp.Fields!horaEgreso = IIf(IsNull(mrs_Tmp1.Fields!horaEgreso), mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!horaEgreso)
                                      mrs_Tmp.Fields!dFuenteFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                                      mrs_Tmp.Fields!idTipoServicio = mrs_Tmp1.Fields!idTipoServicio
                                Else
                                      mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalPorPagar
                                End If
                                mrs_Tmp.Update
                                ReporteHRAagregaCtas mrs_Tmp, mb_EnResumen, rsTmp.Fields!TotalPorPagar, oConexion
                                lnTotalPorCuenta = lnTotalPorCuenta + rsTmp.Fields!TotalPorPagar
                              End If
                              rsTmp.MoveNext
                           Loop
                        End If
                        rsTmp.Close
                        'farmacia
                        lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!generaPago = 1, True, False)
                        If lbEsCuentaPagante = True Then
                            Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "idEstadoFacturacion=1 or idEstadoFacturacion=4"
                        Else
                            Set rsTmp = mo_ReglasFarmacia.farmMovimientoVentasXcuenta(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "totalFinanciado>0 and idTipoFinanciamiento=" & mrs_Tmp1.Fields!IdFormaPago
                        End If
                        If rsTmp.RecordCount > 0 Then
                           rsTmp.MoveFirst
                           Do While Not rsTmp.EOF
                              If rsTmp.Fields!idEstadoMovimiento = 1 Then
                                lbPrimeraVez = True
                                If mrs_Tmp.RecordCount > 0 Then
                                   mrs_Tmp.MoveFirst
                                   Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 5 Then
                                         lbPrimeraVez = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                   Loop
                                End If
                                If lbPrimeraVez = True Then
                                      mrs_Tmp.AddNew
                                      mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                      mrs_Tmp.Fields!Paciente = lcTexto2
                                      mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                      mrs_Tmp.Fields!IdPuntoCarga = 5   'farmacia
                                      mrs_Tmp.Fields!dPuntoCarga = "Farmacia"
                                      mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalFinanciado
                                      mrs_Tmp.Fields!fechaEgreso = IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso), mrs_Tmp1.Fields!FEchaIngreso, mrs_Tmp1.Fields!fechaEgreso)
                                      mrs_Tmp.Fields!horaEgreso = IIf(IsNull(mrs_Tmp1.Fields!horaEgreso), mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!horaEgreso)
                                      mrs_Tmp.Fields!dFuenteFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                                      mrs_Tmp.Fields!idTipoServicio = mrs_Tmp1.Fields!idTipoServicio
                                Else
                                      mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalFinanciado
                                End If
                                ReporteHRAagregaCtas mrs_Tmp, mb_EnResumen, rsTmp.Fields!TotalFinanciado, oConexion
                                lnTotalPorCuenta = lnTotalPorCuenta + rsTmp.Fields!TotalFinanciado
                              End If
                              rsTmp.MoveNext
                           Loop
                        End If
                        rsTmp.Close
                        'Añade IMPORTE TOTAL DE ESA CUENTA(200)
                        If lnTotalPorCuenta > 0 Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                            mrs_Tmp.Fields!Paciente = lcTexto2
                            mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                            mrs_Tmp.Fields!IdPuntoCarga = 200
                            mrs_Tmp.Fields!dPuntoCarga = "Total (1) Consumos"
                            mrs_Tmp.Fields!consumo = lnTotalPorCuenta
                            mrs_Tmp.Update
                        End If
                        'Pagos realizados de la Cuenta
                        Set rsTmp10 = mo_ReglasCaja.CajaComprobantesPagoXcuenta(ml_IdCuenta, oConexion)
                        rsTmp10.Filter = "idEstadoComprobante=4"
                        If rsTmp10.RecordCount > 0 Then
                           ReporteHRApagos rsTmp10, ml_IdCuenta, oConexion, lbEsCuentaPagante, mb_EnResumen
                           rsTmp10.MoveFirst
                           Do While Not rsTmp10.EOF
                              '***************** GalenHos v.3.0 (inicio)*****************
                              'Pagos Adelantados(201)
                              If rsTmp10.Fields!Adelantos > 0 Then
                                lbPrimeraVez = True
                                If mrs_Tmp.RecordCount > 0 Then
                                   mrs_Tmp.MoveFirst
                                   Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 201 Then
                                         lbPrimeraVez = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                   Loop
                                End If
                                If lbPrimeraVez = True Then
                                      mrs_Tmp.AddNew
                                      mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                      mrs_Tmp.Fields!Paciente = lcTexto2
                                      mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                      mrs_Tmp.Fields!IdPuntoCarga = 201
                                      mrs_Tmp.Fields!dPuntoCarga = "Total (2) Pagos Adelantados"
                                      mrs_Tmp.Fields!consumo = rsTmp10.Fields!Adelantos
                                Else
                                      mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Dctos
                                End If
                              End If
                              '***************** GalenHos v.3.0 (fin)*****************
                              'Exoneraciones(202)
                              If rsTmp10.Fields!Exoneraciones > 0 Then
                                lbPrimeraVez = True
                                If mrs_Tmp.RecordCount > 0 Then
                                   mrs_Tmp.MoveFirst
                                   Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 202 Then
                                         lbPrimeraVez = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                   Loop
                                End If
                                If lbPrimeraVez = True Then
                                      mrs_Tmp.AddNew
                                      mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                      mrs_Tmp.Fields!Paciente = lcTexto2
                                      mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                      mrs_Tmp.Fields!IdPuntoCarga = 202
                                      mrs_Tmp.Fields!dPuntoCarga = "Total (3) Exoneraciones"
                                      mrs_Tmp.Fields!consumo = rsTmp10.Fields!Exoneraciones
                                Else
                                      mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Exoneraciones
                                End If
                              End If
                              'Total Boleta(203)
                              If rsTmp10.Fields!Total > 0 Then
                                lbPrimeraVez = True
                                If mrs_Tmp.RecordCount > 0 Then
                                   mrs_Tmp.MoveFirst
                                   Do While Not mrs_Tmp.EOF
                                      If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 203 Then
                                         lbPrimeraVez = False
                                         Exit Do
                                      End If
                                      mrs_Tmp.MoveNext
                                   Loop
                                End If
                                If lbPrimeraVez = True Then
                                      mrs_Tmp.AddNew
                                      mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                      mrs_Tmp.Fields!Paciente = lcTexto2
                                      mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                      mrs_Tmp.Fields!IdPuntoCarga = 203
                                      mrs_Tmp.Fields!dPuntoCarga = "Total (4) Pagado"
                                      mrs_Tmp.Fields!consumo = rsTmp10.Fields!Total
                                Else
                                      mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Total
                                End If
                              End If
                              '
                              rsTmp10.MoveNext
                           Loop
                           rsTmp10.Close
                        Else
                            rsTmp10.Close
                            'Busca Exoneraciones Servicios registradas sin Boleta (202)
                            Set rsTmp = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuenta(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "idFuenteFinanciamiento=0"
                            If rsTmp.RecordCount > 0 Then
                               rsTmp.MoveFirst
                               Do While Not rsTmp.EOF
                                    If rsTmp.Fields!TotalFinanciado > 0 Then
                                      lbPrimeraVez = True
                                      If mrs_Tmp.RecordCount > 0 Then
                                         mrs_Tmp.MoveFirst
                                         Do While Not mrs_Tmp.EOF
                                            If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 202 Then
                                               lbPrimeraVez = False
                                               Exit Do
                                            End If
                                            mrs_Tmp.MoveNext
                                         Loop
                                      End If
                                      If lbPrimeraVez = True Then
                                            mrs_Tmp.AddNew
                                            mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                            mrs_Tmp.Fields!Paciente = lcTexto2
                                            mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                            mrs_Tmp.Fields!IdPuntoCarga = 202
                                            mrs_Tmp.Fields!dPuntoCarga = "Total (3) Exoneraciones"
                                            mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalFinanciado
                                      Else
                                            mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalFinanciado
                                      End If
                                    End If
                                    rsTmp.MoveNext
                                Loop
                            End If
                            rsTmp.Close
                            'Busca Exoneraciones Farmacia registradas sin Boleta (202)
                            Set rsTmp = mo_ReglasFacturacion.FacturacionBienesFinanciamientosXcuenta(ml_IdCuenta, oConexion)
                            rsTmp.Filter = "idFuenteFinanciamiento=0"
                            If rsTmp.RecordCount > 0 Then
                               rsTmp.MoveFirst
                               Do While Not rsTmp.EOF
                                    If rsTmp.Fields!TotalFinanciado > 0 Then
                                      lbPrimeraVez = True
                                      If mrs_Tmp.RecordCount > 0 Then
                                         mrs_Tmp.MoveFirst
                                         Do While Not mrs_Tmp.EOF
                                            If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 202 Then
                                               lbPrimeraVez = False
                                               Exit Do
                                            End If
                                            mrs_Tmp.MoveNext
                                         Loop
                                      End If
                                      If lbPrimeraVez = True Then
                                            mrs_Tmp.AddNew
                                            mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                            mrs_Tmp.Fields!Paciente = lcTexto2
                                            mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                            mrs_Tmp.Fields!IdPuntoCarga = 202
                                            mrs_Tmp.Fields!dPuntoCarga = "Total (3) Exoneraciones"
                                            mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalFinanciado
                                      Else
                                            mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalFinanciado
                                      End If
                                    End If
                                    rsTmp.MoveNext
                                 Loop
                            End If
                            rsTmp.Close
                        End If
                        'solo para Reporte HRA- Si es un Reembolso
                        If mb_EnResumen = True And lbEsCuentaPagante = False Then
                           Set rsTmp10 = mo_ReglasFacturacion.FacturacionReembolsosXcuenta(ml_IdCuenta, oConexion)
                            
                            If rsTmp10.RecordCount > 0 Then
                               ReporteHRApagos rsTmp10, ml_IdCuenta, oConexion, lbEsCuentaPagante, mb_EnResumen
                            End If
                            rsTmp10.Close
                        End If
                        '
                        mrs_Tmp1.MoveNext
                   Loop
                   mrs_Tmp1.Close
                End If
        Next
        '***********************atenciones Consultorios Externos*****************************************
        If mrs_Tmp1.State = 1 Then
           Set mrs_Tmp1 = Nothing
        End If
        If rsTmp.State = 1 Then
           Set rsTmp = Nothing
        End If
        If mb_EnResumen = True Then
           Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaEgresoMedico(mda_FechaInicio, mda_FechaFin, True, oConexion, 0)
        Else
           Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaIngresoYfuenteFinanciamiento(mda_FechaInicio, mda_FechaFin, ml_IdPlan)
           mrs_Tmp1.Filter = "idEstadoAtencion<>0 and idTipoServicio=1"
        End If
        ml_IdCuenta = mrs_Tmp1.RecordCount
        If ml_IdCuenta > 0 Then
           f = 0
           lRecordCount = ml_IdCuenta
           mrs_Tmp1.MoveFirst
           Do While Not mrs_Tmp1.EOF
                f = f + 1
                RaiseEvent ProgressActualizaValor(f, lRecordCount)
                '
                lbEsAdmisionCE = False
                lnSaldoInicial = mrs_Tmp1.Fields!NroHistoriaClinica
                ml_IdCuenta = mrs_Tmp1.Fields!idCuentaAtencion
                lcTexto2 = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                lcTexto2 = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                lcTexto2 = lcTexto2 & " (Hc: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False) & ") (Cta: " & Format(ml_IdCuenta, "########") & ")"
                lcTexto2 = lcTexto2 & " (Ce)"
                lnTotalPorCuenta = 0
                'servicios
                lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!generaPago = 1, True, False)
                If rsTmp.State = 1 Then rsTmp.Close
                If lbEsCuentaPagante = True Then
                    Set rsTmp = mo_ReglasFacturacion.FacturacionServiciosPagosSeleccionarPorCuenta(ml_IdCuenta, oConexion)
                    rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
                    
                Else
                    Set rsTmp = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXcuenta(ml_IdCuenta, oConexion)
                    rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
                End If
                
                If rsTmp.RecordCount > 0 Then
                   rsTmp.MoveFirst
                   Do While Not rsTmp.EOF
                      If rsTmp.Fields!TotalPorPagar > 0 And rsTmp.Fields!idProducto <> Val(lcParametro245) Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga Then
                                 lbPrimeraVez = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbPrimeraVez = True Then
                              
                              If rsTmp.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionCE Then
                                 lbEsAdmisionCE = True
                              End If
                              '
                              oRsPuntosDeCarga.MoveFirst
                              oRsPuntosDeCarga.Find "idPuntoCarga=" & rsTmp.Fields!IdPuntoCarga
                              lcTexto1 = IIf(oRsPuntosDeCarga.EOF, "", Trim(oRsPuntosDeCarga.Fields!descripcion))
                              '
                              mrs_Tmp.AddNew
                              mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                              mrs_Tmp.Fields!Paciente = lcTexto2
                              mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                              mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga
                              mrs_Tmp.Fields!dPuntoCarga = lcTexto1
                              mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalPorPagar
                              mrs_Tmp.Fields!fechaEgreso = mrs_Tmp1.Fields!FEchaIngreso
                              mrs_Tmp.Fields!horaEgreso = mrs_Tmp1.Fields!HoraIngreso
                              mrs_Tmp.Fields!dFuenteFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                              mrs_Tmp.Fields!idTipoServicio = mrs_Tmp1.Fields!idTipoServicio
                        Else
                              mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalPorPagar
                        End If
                        mrs_Tmp.Update
                        ReporteHRAagregaCtas mrs_Tmp, mb_EnResumen, rsTmp.Fields!TotalPorPagar, oConexion
                        lnTotalPorCuenta = lnTotalPorCuenta + rsTmp.Fields!TotalPorPagar
                      End If
                      rsTmp.MoveNext
                   Loop
                End If
                rsTmp.Close
                'farmacia
                lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!generaPago = 1, True, False)
                If rsTmp.State = 1 Then rsTmp.Close
                If lbEsCuentaPagante = True Then
                    Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion(ml_IdCuenta, oConexion)
                    rsTmp.Filter = "idEstadoFacturacion=1 or idEstadoFacturacion=4"
                Else
                    Set rsTmp = mo_ReglasFacturacion.FacturacionBienesFinanciamientosXcuentaYconexion(ml_IdCuenta, oConexion)
                    rsTmp.Filter = "totalFinanciado>0 and idTipoFinanciamiento=" & mrs_Tmp1.Fields!IdFormaPago
                End If
                If rsTmp.RecordCount > 0 Then
                   rsTmp.MoveFirst
                   Do While Not rsTmp.EOF
                      If rsTmp.Fields!idEstadoMovimiento = 1 Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 5 Then
                                 lbPrimeraVez = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbPrimeraVez = True Then
                              mrs_Tmp.AddNew
                              mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                              mrs_Tmp.Fields!Paciente = lcTexto2
                              mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                              mrs_Tmp.Fields!IdPuntoCarga = 5   'farmacia
                              mrs_Tmp.Fields!dPuntoCarga = "Farmacia"
                              mrs_Tmp.Fields!consumo = rsTmp.Fields!TotalFinanciado
                              mrs_Tmp.Fields!fechaEgreso = mrs_Tmp1.Fields!FEchaIngreso
                              mrs_Tmp.Fields!horaEgreso = mrs_Tmp1.Fields!HoraIngreso
                              mrs_Tmp.Fields!dFuenteFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                              mrs_Tmp.Fields!idTipoServicio = mrs_Tmp1.Fields!idTipoServicio
                        Else
                              mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp.Fields!TotalFinanciado
                        End If
                        ReporteHRAagregaCtas mrs_Tmp, mb_EnResumen, rsTmp.Fields!TotalFinanciado, oConexion
                        lnTotalPorCuenta = lnTotalPorCuenta + rsTmp.Fields!TotalFinanciado
                      End If
                      'rsReporte.Close
                      rsTmp.MoveNext
                   Loop
                End If
                rsTmp.Close
                'Añade IMPORTE TOTAL DE ESA CUENTA(200)
                If lnTotalPorCuenta > 0 Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                    mrs_Tmp.Fields!Paciente = lcTexto2
                    mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                    mrs_Tmp.Fields!IdPuntoCarga = 200
                    mrs_Tmp.Fields!dPuntoCarga = "Total (1) Consumos"
                    mrs_Tmp.Fields!consumo = lnTotalPorCuenta
                    mrs_Tmp.Update
                End If
                'Pagos realizados de la Cuenta
                Set rsTmp10 = mo_ReglasCaja.CajaComprobantesPagoXcuenta(ml_IdCuenta, oConexion)
                rsTmp10.Filter = "idEstadoComprobante=4"
                If rsTmp10.RecordCount > 0 Then
                   ReporteHRApagos rsTmp10, ml_IdCuenta, oConexion, lbEsCuentaPagante, mb_EnResumen
                   rsTmp10.MoveFirst
                   Do While Not rsTmp10.EOF
                      '***************** GalenHos v.3.0 (inicio)*****************
                      'Pagos Adelantados(201)
                      If rsTmp10.Fields!Adelantos > 0 Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 201 Then
                                 lbPrimeraVez = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbPrimeraVez = True Then
                              mrs_Tmp.AddNew
                              mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                              mrs_Tmp.Fields!Paciente = lcTexto2
                              mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                              mrs_Tmp.Fields!IdPuntoCarga = 201
                              mrs_Tmp.Fields!dPuntoCarga = "Total (2) Pagos Adelantados"
                              mrs_Tmp.Fields!consumo = rsTmp10.Fields!Adelantos
                        Else
                              mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Dctos
                        End If
                      End If
                      '***************** GalenHos v.3.0 (fin)*****************
                      'Exoneraciones(202)
                      If rsTmp10.Fields!Exoneraciones > 0 Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 202 Then
                                 lbPrimeraVez = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbPrimeraVez = True Then
                              mrs_Tmp.AddNew
                              mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                              mrs_Tmp.Fields!Paciente = lcTexto2
                              mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                              mrs_Tmp.Fields!IdPuntoCarga = 202
                              mrs_Tmp.Fields!dPuntoCarga = "Total (3) Exoneraciones"
                              mrs_Tmp.Fields!consumo = rsTmp10.Fields!Exoneraciones
                        Else
                              mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Exoneraciones
                        End If
                      End If
                      'Total Boleta(203)
                      If rsTmp10.Fields!Total > 0 Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta And mrs_Tmp.Fields!IdPuntoCarga = 203 Then
                                 lbPrimeraVez = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbPrimeraVez = True Then
                              mrs_Tmp.AddNew
                              mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                              mrs_Tmp.Fields!Paciente = lcTexto2
                              mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                              mrs_Tmp.Fields!IdPuntoCarga = 203
                              mrs_Tmp.Fields!dPuntoCarga = "Total (4) Pagado"
                              mrs_Tmp.Fields!consumo = rsTmp10.Fields!Total
                        Else
                              mrs_Tmp.Fields!consumo = mrs_Tmp.Fields!consumo + rsTmp10.Fields!Total
                        End If
                      End If
                      '
                      rsTmp10.MoveNext
                   Loop
                End If
                rsTmp10.Close
                'solo para Reporte HRA- Si es un Reembolso
                If mb_EnResumen = True And lbEsCuentaPagante = False Then
                   Set rsTmp10 = mo_ReglasFacturacion.FacturacionReembolsosXcuenta(ml_IdCuenta, oConexion)
                    If rsTmp10.RecordCount > 0 Then
                       ReporteHRApagos rsTmp10, ml_IdCuenta, oConexion, lbEsCuentaPagante, mb_EnResumen
                    End If
                    rsTmp10.Close
                End If
                '
                mrs_Tmp1.MoveNext
           Loop
        End If
        oConexion.Close
        '
        lcHoraFinal = lcBuscaParametro.RetornaHoraServidorSQL1
        lcSql = DateDiff("s", CDate(lcHOraInicio), CDate(lcHoraFinal))
        'MsgBox "Termino el Proceso OK" & Chr(13) & Chr(13) & "Segundos de Diferencia: " & lcSql, vbInformation, "Mensaje"
        
        Exit Sub
ErrProcesar:
    MsgBox Err.Description
    Resume
End Sub


Sub ReporteHRAagregaCtas(mrs_Tmp As Recordset, mb_EnResumen As Boolean, lnConsumo As Double, oConexion As Connection)
    On Error GoTo ErrHRAAdd
    If mb_EnResumen = True Then
        Dim lbEncontroRegistro As Boolean
        
        lbEncontroRegistro = False
        If rsTmpSOAT.RecordCount > 0 Then
           rsTmpSOAT.MoveFirst
           rsTmpSOAT.Find "idCuentaAtencion=" & mrs_Tmp.Fields!idCuentaAtencion
           If Not rsTmpSOAT.EOF Then
              lbEncontroRegistro = True
           End If
        End If
        If lbEncontroRegistro = False Then
           rsTmpSOAT.AddNew
           rsTmpSOAT.Fields!idTipoFinanciamiento = mrs_Tmp!IdFormaPago
           rsTmpSOAT.Fields!idCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
           rsTmpSOAT.Fields!Paciente = mrs_Tmp.Fields!Paciente
           rsTmpSOAT.Fields!nroHistoria = mrs_Tmp.Fields!nroHistoria
           rsTmpSOAT.Fields!Falta = mrs_Tmp.Fields!fechaEgreso
           rsTmpSOAT.Fields!Halta = mrs_Tmp.Fields!horaEgreso
           rsTmpSOAT.Fields!dFinanciamiento = mrs_Tmp.Fields!dFuenteFinanciamiento
           rsTmpSOAT.Fields!origen = IIf(mrs_Tmp.Fields!idTipoServicio = 1, "CE", IIf(mrs_Tmp.Fields!idTipoServicio = 3, "Hosp", "Emerg"))
           If mrs_Tmp.Fields!idCuentaAtencion > 0 Then
              If mrs_Tmp.Fields!idTipoServicio = 1 Then
                 rsTmpSOAT.Fields!Servicio = mo_ReglasAdmision.DevuelveServicioIngreso(mrs_Tmp!idCuentaAtencion, oConexion)
              Else
                rsTmpSOAT.Fields!Servicio = mo_ReglasAdmision.DevuelveServicioAlta(mrs_Tmp!idCuentaAtencion, oConexion)
              End If
           End If
           If InStr(UCase(mrs_Tmp!dFuenteFinanciamiento), "SIS") > 0 Then
              rsTmpSOAT.Fields!fua = mo_ReglasAdmision.SisFuaAtencionDevuelveNumeroFUA(mrs_Tmp!idCuentaAtencion)
           End If
        End If
        Select Case mrs_Tmp.Fields!IdPuntoCarga
        Case sghPuntosCargaBasicos.sghPtoCargaRayosX
             rsTmpSOAT.Fields!dRx = rsTmpSOAT.Fields!dRx + lnConsumo
        Case sghPuntosCargaBasicos.sghPtoCargaFarmacia
             rsTmpSOAT.Fields!dFarmacia = rsTmpSOAT.Fields!dFarmacia + lnConsumo
        Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
             rsTmpSOAT.Fields!dLab = rsTmpSOAT.Fields!dLab + lnConsumo
        Case Else
             rsTmpSOAT.Fields!dOtros = rsTmpSOAT.Fields!dOtros + lnConsumo
        End Select
        rsTmpSOAT.Fields!tFacturado = rsTmpSOAT.Fields!tFacturado + lnConsumo
        rsTmpSOAT.Fields!tDeuda = rsTmpSOAT.Fields!tDeuda + lnConsumo

        rsTmpSOAT.Update
    End If
    Exit Sub
ErrHRAAdd:
    MsgBox Err.Description
    Resume
End Sub













Sub ReporteHRApagos(rsTmp10 As Recordset, lnIdCuentaAtencion As Long, moConexion As Connection, lbEsCuentaPagante As Boolean, mb_EnResumen As Boolean)
    On Error GoTo ErrHRApagos
    If mb_EnResumen = True Then
        Dim rsTmpReemb2 As New Recordset, oRsTmpReemb1 As New Recordset
        Dim lnFarmacia As Double, lnServicio As Double, lnRx As Double, lnLaboratorio As Double, lnOtros As Double, lnPucho As Double, lnDctos As Double
        Dim lnTFarmacia As Double, lnTRx As Double, lnTLaboratorio As Double, lnTOtros As Double, lnTDctos As Double, lnDctosXitem As Double
        Dim lcSql As String, lnTtotalFacturadoCuenta As Double, LnTotalPagadoCuenta As Double, lnTotalDeudaCuenta As Double
        Dim lnTotalFacturadoCuenta As Double, lnTotalPagadoEnReembolso As Double
        lnTFarmacia = 0: lnTRx = 0: lnTLaboratorio = 0: lnTOtros = 0: lnTDctos = 0: lnTotalPagadoEnReembolso = 0
        rsTmp10.MoveFirst
        Do While Not rsTmp10.EOF
If lnIdCuentaAtencion = 33655 Then
lnServicio = 0
End If
           If rsTmp10.Fields!IdEstadoComprobante = 4 Then   'solo pagados
                If lbEsCuentaPagante = True Then
                        lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
                        If rsTmp10.Fields!IdTipoOrden <> 1 Then
                            '**** Cuenta de un Paciente PAGANTE - BOLETA DE FARMACIA
                            lnTFarmacia = rsTmp10.Fields!Total
                            lnServicio = 0
                            lnTDctos = lnTDctos + (rsTmp10.Fields!Total - rsTmp10.Fields!Subtotal)
                        Else
                            '**** Cuenta de un Paciente PAGANTE - BOLETA DE SERVICIOS
                            lnServicio = rsTmp10.Fields!Total
                            lnDctos = (rsTmp10.Fields!Total - rsTmp10.Fields!Subtotal)
                            lnTDctos = lnTDctos + lnDctos
                            Set oRsTmpReemb1 = mo_ReglasFacturacion.FacturacionServiciosPagosSeleccionarPorIdComprobantePago(rsTmp10.Fields!IdComprobantePago, moConexion)
                            If oRsTmpReemb1.RecordCount > 0 Then
                               oRsTmpReemb1.MoveFirst
                               Do While Not oRsTmpReemb1.EOF
                                     lnDctosXitem = 0
                                     If lnDctos <> 0 Then
                                        Set rsTmpReemb2 = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXidOrdenIdProducto(oRsTmpReemb1.Fields!idOrden, oRsTmpReemb1.Fields!idProducto, moConexion)
                                        If rsTmpReemb2.RecordCount > 0 Then
                                           rsTmpReemb2.MoveFirst
                                           Do While Not rsTmpReemb2.EOF
                                              lnDctosXitem = lnDctosXitem + rsTmpReemb2.Fields!TotalFinanciado
                                              rsTmpReemb2.MoveNext
                                           Loop
                                        End If
                                        rsTmpReemb2.Close
                                     End If
                                      Set rsTmpReemb2 = mo_reglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & oRsTmpReemb1.Fields!idProducto, moConexion)
                                     If rsTmpReemb2.RecordCount = 0 Then
                                        lnOtros = lnOtros + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                     Else
                                        If rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Then
                                           'Laboratorio
                                           lnLaboratorio = lnLaboratorio + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                        ElseIf rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX Then
                                           'Rx
                                           lnRx = lnRx + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                        Else
                                           lnOtros = lnOtros + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                        End If
                                     End If
                                     rsTmpReemb2.Close
                                     oRsTmpReemb1.MoveNext
                               Loop
                            End If
                            oRsTmpReemb1.Close
                        End If
                Else
                         '**** Cuenta de un Paciente con SEGUROS - BOLETA DE REEMBOLSOS
                         lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
                         Set rsTmpReemb2 = mo_ReglasFacturacion.ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta(rsTmp10.Fields!IdComprobantePago, moConexion, lnIdCuentaAtencion)
                         If rsTmpReemb2.RecordCount > 0 Then
                            lnTotalPagadoEnReembolso = lnTotalPagadoEnReembolso + rsTmp10.Fields!Total
                             'Reembolsos
                             rsTmpReemb2.MoveFirst
                             Do While Not rsTmpReemb2.EOF
                                lnFarmacia = lnFarmacia + rsTmpReemb2.Fields!ReembolsoPagadoFarmacia
                                lnServicio = lnServicio + rsTmpReemb2.Fields!ReembolsoPagadoServicio
                                rsTmpReemb2.MoveNext
                             Loop
                             rsTmpReemb2.Close
                             lnTFarmacia = lnTFarmacia + lnFarmacia
                             If lnServicio > 0 Then
                                Set oRsTmpReemb1 = mo_ReglasFacturacion.ReembolsoDetalleItemSeleccionarPorIdCuenta(lnIdCuentaAtencion, moConexion)
                                If oRsTmpReemb1.RecordCount > 0 Then
                                   oRsTmpReemb1.MoveFirst
                                   Do While Not oRsTmpReemb1.EOF
                                         If oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Then
                                            'Laboratorio
                                            lnLaboratorio = lnLaboratorio + oRsTmpReemb1.Fields!TotalFinanciado
                                         ElseIf oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX Then
                                            'Rx
                                            lnRx = lnRx + oRsTmpReemb1.Fields!TotalFinanciado
                                         Else
                                            lnOtros = lnOtros + oRsTmpReemb1.Fields!TotalFinanciado
                                         End If
                                         oRsTmpReemb1.MoveNext
                                   Loop
                                End If
                                oRsTmpReemb1.Close
                             End If
                        Else
                            rsTmpReemb2.Close
                            'Otros pagos que no cubre el SEGURO
                            lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
                            If rsTmp10.Fields!IdTipoOrden <> 1 Then
                                '**** Cuenta de un Paciente SEGUROS - BOLETA DE FARMACIA
                                lnTFarmacia = rsTmp10.Fields!Total
                                lnServicio = 0
                                lnTDctos = lnTDctos + (rsTmp10.Fields!Total - rsTmp10.Fields!Subtotal)
                            Else
                                '**** Cuenta de un Paciente SEGUROS - BOLETA DE SERVICIOS
                                lnServicio = rsTmp10.Fields!Total
                                lnDctos = (rsTmp10.Fields!Total - rsTmp10.Fields!Subtotal)
                                lnTDctos = lnTDctos + lnDctos
                                Set oRsTmpReemb1 = mo_ReglasFacturacion.FacturacionServiciosPagosSeleccionarPorIdComprobantePago(rsTmp10.Fields!IdComprobantePago, moConexion)
                                If oRsTmpReemb1.RecordCount > 0 Then
                                   oRsTmpReemb1.MoveFirst
                                   Do While Not oRsTmpReemb1.EOF
                                         lnDctosXitem = 0
                                         If lnDctos <> 0 Then
                                            Set rsTmpReemb2 = mo_ReglasFacturacion.FacturacionServicioFinanciamientosXidOrdenIdProducto(oRsTmpReemb1.Fields!idOrden, oRsTmpReemb1.Fields!idProducto, moConexion)
                                            If rsTmpReemb2.RecordCount > 0 Then
                                               rsTmpReemb2.MoveFirst
                                               Do While Not rsTmpReemb2.EOF
                                                  lnDctosXitem = lnDctosXitem + rsTmpReemb2.Fields!TotalFinanciado
                                                  rsTmpReemb2.MoveNext
                                               Loop
                                            End If
                                            rsTmpReemb2.Close
                                         End If
                                         Set rsTmpReemb2 = mo_reglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & oRsTmpReemb1.Fields!idProducto, moConexion)
                                         If rsTmpReemb2.RecordCount = 0 Then
                                            lnOtros = lnOtros + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                         Else
                                            If rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Then
                                               'Laboratorio
                                               lnLaboratorio = lnLaboratorio + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                            ElseIf rsTmpReemb2.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX Then
                                               'Rx
                                               lnRx = lnRx + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                            Else
                                               lnOtros = lnOtros + oRsTmpReemb1.Fields!Total - lnDctosXitem
                                            End If
                                         End If
                                         rsTmpReemb2.Close
                                         oRsTmpReemb1.MoveNext
                                   Loop
                                End If
                                oRsTmpReemb1.Close
                            End If
                        End If
               End If
               If lnServicio > 0 Then
                   If lnServicio <> (lnRx + lnLaboratorio + lnOtros) Then  'cuadre
                      If lnServicio < (lnRx + lnLaboratorio + lnOtros) Then
                         lnPucho = (lnRx + lnLaboratorio + lnOtros) - lnServicio
                         lnPucho = Round(lnPucho / 3, 2)
                         lnRx = lnRx - lnPucho
                         lnLaboratorio = lnLaboratorio - lnPucho
                         lnOtros = lnOtros - lnPucho
                      Else
                         lnPucho = lnServicio - (lnRx + lnLaboratorio + lnOtros)
                         lnOtros = lnOtros + lnPucho
                      End If
                   End If
                   'lnTFarmacia = lnTFarmacia + lnFarmacia
                   lnTRx = lnTRx + lnRx
                   lnTLaboratorio = lnTLaboratorio + lnLaboratorio
                   lnTOtros = lnTOtros + lnOtros
               End If
           End If
           rsTmp10.MoveNext
        Loop
        If lnRx < 0 Or lnLaboratorio < 0 Or lnOtros < 0 Then   'cantidades negativas
           If lnRx < 0 Then
              If (lnLaboratorio + lnRx) >= 0 Then
                 lnLaboratorio = lnLaboratorio + lnRx
              Else
                 lnOtros = lnOtros + lnRx
              End If
              lnRx = 0
           End If
           If lnLaboratorio < 0 Then
              If (lnRx + lnLaboratorio) >= 0 Then
                 lnRx = lnRx + lnLaboratorio
              Else
                 lnOtros = lnOtros + lnLaboratorio
              End If
              lnLaboratorio = 0
           End If
           If lnOtros < 0 Then
              If (lnRx + lnOtros) >= 0 Then
                 lnRx = lnRx + lnOtros
              Else
                 lnLaboratorio = lnLaboratorio + lnOtros
              End If
              lnOtros = 0
           End If
        End If
        rsTmpSOAT.MoveFirst
        rsTmpSOAT.Find "idCuentaAtencion=" & lnIdCuentaAtencion
        If Not rsTmpSOAT.EOF Then
            lnTotalFacturadoCuenta = Round(rsTmpSOAT.Fields!tFacturado, 2)
            LnTotalPagadoCuenta = (lnTFarmacia + lnTRx + lnTLaboratorio + lnTOtros)
            If lbEsCuentaPagante = True Then
               lnTotalDeudaCuenta = (lnTotalFacturadoCuenta - (LnTotalPagadoCuenta - (lnTDctos)))
            Else
               lnTotalDeudaCuenta = lnTotalFacturadoCuenta - lnTotalPagadoEnReembolso
            End If
            rsTmpSOAT.Fields!PFarmacia = lnTFarmacia
            rsTmpSOAT.Fields!PRx = lnTRx
            rsTmpSOAT.Fields!PLab = lnTLaboratorio
            rsTmpSOAT.Fields!POtros = lnTOtros
            
            rsTmpSOAT.Fields!tPdctos = lnTDctos
            rsTmpSOAT.Fields!tPagado = LnTotalPagadoCuenta
            rsTmpSOAT.Fields!tDeuda = lnTotalDeudaCuenta
            rsTmpSOAT.Update
        End If
    End If
    Set rsTmpReemb2 = Nothing: Set oRsTmpReemb1 = Nothing
    Exit Sub
ErrHRApagos:
    MsgBox Err.Description
    Resume
End Sub






Sub ProcesaDatosYllenaTmpRapidamente(lnIdResponsable As Long, mda_FechaInicio As Date, mda_FechaFin As Date, _
                                     ml_IdPlan As Long, mb_EnResumen As Boolean, lbAgregarCuentasSinAlta As Boolean)
    On Error GoTo ErrProcesar
    Dim rsReporte As New ADODB.Recordset
    Dim rsTmp As New Recordset
    Dim rsTmp10 As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim oRsPuntosDeCarga As New Recordset
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
    Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
    Dim ldFechaPrincipio As Date
    Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim oConexion As New ADODB.Connection
    Dim lbPrimeraVez As Boolean
    Dim lcTexto1 As String: Dim lcTexto2 As String: Dim lcTexto3 As String
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim lnDebioPagar As Double, lnTotalPorCuenta As Double
    Dim lnIdTipoServicio As Long, lnIdDepartamento As Long, lnIdEspacialidad As Long, lnIdServicioPaciente As Long
    Dim lcDtipoServicio As String, lcDpto As String, lcDespecialidad As String, lcDServicio As String
    Dim lnSalieron As Integer, lnRetornaron As Integer, lnSalieronYnoRetornaron As Integer
    Dim lnTSalieron As Integer, lnTRetornaron As Integer, lnTSalieronYnoRetornaron As Integer
    Dim lcSql As String, lbEsCuentaPagante As Boolean, lbSoloPasoPorCita As Boolean
    Dim lbProcesaEnServidor As Boolean, lcHOraInicio As String, lcHoraFinal As String
    Dim ml_IdCuenta As Long, lbEncontroRegistro As Boolean
    Dim lcParametro245 As String, lcParametro257 As String
    Dim lnTotalPagadoCta As Double, lnFor As Integer, lnConsumo As Double, lnPagos As Double, lnPagosReembolso As Double
    Dim rsTmpReemb2 As New Recordset, oRsTmpReemb1 As New Recordset
    Dim lnFarmacia As Double, lnServicio As Double, lnRx As Double, lnLaboratorio As Double, lnOtros As Double, lnPucho As Double, lnDctos As Double
    Dim lnTFarmacia As Double, lnTRx As Double, lnTLaboratorio As Double, lnTOtros As Double, lnTDctos As Double, lnDctosXitem As Double
    Dim lnTtotalFacturadoCuenta As Double, LnTotalPagadoCuenta As Double, lnTotalDeudaCuenta As Double
    Dim lnTotalFacturadoCuenta As Double, lnTotalPagadoEnReembolso As Double, lnNroCptConsumidos As Integer
    Dim lnTFarmaciaR As Double, lnTRxR As Double, lnTLaboratorioR As Double, lnTOtrosR As Double
    Dim lnTOtrosC As Double, lnTLaboratorioC As Double, lnTRxC As Double, lnTFarmaciaC As Double
    Dim lRecordCount As Long, f As Long
        lbProcesaEnServidor = True
        lcParametro245 = lcBuscaParametro.SeleccionaFilaParametro(245)
        lcParametro257 = lcBuscaParametro.SeleccionaFilaParametro(257)
        
        lcHOraInicio = lcBuscaParametro.RetornaHoraServidorSQL1
    
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        '
        ReporteHRACreaTmp mb_EnResumen
        GenerarRecordsetTemporal
        If lnIdResponsable = 9 Then
           lnIdResponsable = 1
        End If
        'atenciones Hospitalizacion/Emergencia/ce
        For lnFor = 1 To 3
             Select Case lnFor
             Case 1
                'Hosp/emerg con Alta Medica
                Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaEgresoMedico(mda_FechaInicio, mda_FechaFin, False, oConexion, ml_IdPlan)
             Case 2
                'CE
                Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaEgresoMedico(mda_FechaInicio, mda_FechaFin, True, oConexion, ml_IdPlan)
             Case 3
                'Hosp/emerg sin Alta, con Fecha Ingreso
                If lbAgregarCuentasSinAlta = False Then
                   Exit For
                End If
                Set mrs_Tmp1 = mo_ReglasFacturacion.AtencionesSeleccionarPorFechaIngresoMedicoSinAlta(mda_FechaInicio, mda_FechaFin, False, oConexion, ml_IdPlan)
             End Select
             ml_IdCuenta = mrs_Tmp1.RecordCount
             If ml_IdCuenta > 0 Then

                f = 0
                lRecordCount = ml_IdCuenta
                mrs_Tmp1.MoveFirst
                Do While Not mrs_Tmp1.EOF
                     f = f + 1
                     RaiseEvent ProgressActualizaValor(f, lRecordCount)
                     '
                     lnSaldoInicial = IIf(IsNull(mrs_Tmp1.Fields!NroHistoriaClinica), 0, mrs_Tmp1.Fields!NroHistoriaClinica)
                     ml_IdCuenta = mrs_Tmp1.Fields!idCuentaAtencion
                     'JSPC 04/12/2020 Cambio43 Inicio
                     'Antes: lcTexto2 = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                     lcTexto2 = "(F.Ing: " & mrs_Tmp1.Fields!FEchaIngreso & ") " & Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                      'JSPC 04/12/2020 Cambio43 Fin
                      
                     'JSPC 04/12/2020 Cambio43 Inicio
                     'Se comento todo el código
                     'If Len(lcTexto2) > 30 Then
                     '   lcTexto2 = Left(lcTexto2, 30)
                     'Else
                     '   lcTexto2 = Left(lcTexto2, 30) & String(30 - Len(lcTexto2), " ")
                     'End If
                      'JSPC 04/12/2020 Cambio43 Fin
                     lcTexto2 = lcTexto2 & " (Hc: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False) & ") (Cta: " & Format(ml_IdCuenta, "########") & ")"
                     lcTexto2 = lcTexto2 & IIf(mrs_Tmp1.Fields!idTipoServicio = 3, " (Ho)", IIf(mrs_Tmp1.Fields!idTipoServicio = 1, " (Ce)", " (Em)"))
                     lcTexto2 = lcTexto2 & IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso) And mrs_Tmp1.Fields!idTipoServicio <> 1, " (sin AM)", "")
                     
                     lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!generaPago = 1, True, False)
                     
                     lnNroCptConsumidos = 0: lbSoloPasoPorCita = False
                     lnTotalPorCuenta = 0: lnTotalPagadoCta = 0: LnTotalPagadoCuenta = 0
                     lnTFarmacia = 0: lnTRx = 0: lnTLaboratorio = 0: lnTOtros = 0: lnTDctos = 0: lnTotalPagadoEnReembolso = 0
                     lnTFarmaciaR = 0: lnTRxR = 0: lnTLaboratorioR = 0: lnTOtrosR = 0
                     lnTOtrosC = 0: lnTLaboratorioC = 0: lnTRxC = 0: lnTFarmaciaC = 0
                     'Consumos de Servicios y Farmacia
                     Set rsTmp = mo_ReglasFacturacion.FacturacionCuentasAtencionPtosXcuenta(ml_IdCuenta, oConexion)
                     If rsTmp.RecordCount > 0 Then
                        rsTmp.MoveFirst
                        lnTDctos = IIf(IsNull(rsTmp.Fields!TotalExonerado), 0, rsTmp.Fields!TotalExonerado)
                        Do While Not rsTmp.EOF
                             '
                             lcTexto1 = IIf(IsNull(rsTmp.Fields!descripcion), "", Trim(rsTmp.Fields!descripcion))
                             '
                             If rsTmp.Fields!totalConsumos > 0 Then
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                mrs_Tmp.Fields!Paciente = lcTexto2
                                mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                                mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga
                                mrs_Tmp.Fields!dPuntoCarga = lcTexto1
                                mrs_Tmp.Fields!consumo = rsTmp.Fields!totalConsumos
                                mrs_Tmp.Fields!fechaEgreso = IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso), mrs_Tmp1.Fields!FEchaIngreso, mrs_Tmp1.Fields!fechaEgreso)
                                mrs_Tmp.Fields!horaEgreso = IIf(IsNull(mrs_Tmp1.Fields!horaEgreso), mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!horaEgreso)
                                mrs_Tmp.Fields!dFuenteFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                                mrs_Tmp.Fields!idTipoServicio = mrs_Tmp1.Fields!idTipoServicio
                                mrs_Tmp.Update
                                lnTotalPorCuenta = lnTotalPorCuenta + rsTmp.Fields!totalConsumos
                             End If
                             LnTotalPagadoCuenta = LnTotalPagadoCuenta + rsTmp.Fields!totalPagos
                             If mb_EnResumen = True Then
                                 lnConsumo = rsTmp.Fields!totalConsumos
                                 lnPagos = rsTmp.Fields!totalPagos
                                 lnPagosReembolso = rsTmp.Fields!totalPagosReembolso
                                 Select Case rsTmp.Fields!IdPuntoCarga
                                 Case sghPuntosCargaBasicos.sghPtoCargaFarmacia
                                      lnTFarmaciaC = lnTFarmaciaC + lnConsumo
                                      lnTFarmacia = lnTFarmacia + lnPagos
                                      lnTFarmaciaR = lnTFarmaciaR + lnPagosReembolso
                                 Case sghPuntosCargaBasicos.sghPtoCargaRayosX
                                      lnTRxC = lnTRxC + lnConsumo
                                      lnTRx = lnTRx + lnPagos
                                      lnTRxR = lnTRxR + lnPagosReembolso
                                 Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
                                      lnTLaboratorioC = lnTLaboratorioC + lnConsumo
                                      lnTLaboratorio = lnTLaboratorio + lnPagos
                                      lnTLaboratorioR = lnTLaboratorioR + lnPagosReembolso
                                 Case Else
                                      lnTOtrosC = lnTOtrosC + lnConsumo
                                      lnTOtros = lnTOtros + lnPagos
                                      lnTOtrosR = lnTOtrosR + lnPagosReembolso
                                      lbSoloPasoPorCita = True
                                 End Select
                                 lnNroCptConsumidos = lnNroCptConsumidos + 1
                             End If
                             rsTmp.MoveNext
                        Loop
                        
                        If mb_EnResumen = True Then
                             lnTotalFacturadoCuenta = lnTFarmaciaC + lnTRxC + lnTLaboratorioC + lnTOtrosC
                             LnTotalPagadoCuenta = (lnTFarmacia + lnTRx + lnTLaboratorio + lnTOtros)
                             lnTotalPagadoEnReembolso = lnTFarmaciaR + lnTRxR + lnTLaboratorioR + lnTOtrosR
                             If lbEsCuentaPagante = True Then
                                lnTotalDeudaCuenta = (lnTotalFacturadoCuenta - (LnTotalPagadoCuenta + lnTDctos))
                             Else
                                lnTotalDeudaCuenta = lnTotalFacturadoCuenta - lnTotalPagadoEnReembolso
                             End If
                             If lnTotalFacturadoCuenta > 0 Or LnTotalPagadoCuenta > 0 Or lnTotalPagadoEnReembolso > 0 Then
                                rsTmpSOAT.AddNew
                                rsTmpSOAT.Fields!idTipoFinanciamiento = mrs_Tmp1!IdFormaPago
                                rsTmpSOAT.Fields!idCuentaAtencion = ml_IdCuenta
                                rsTmpSOAT.Fields!Paciente = lcTexto2
                                rsTmpSOAT.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                                rsTmpSOAT.Fields!origen = IIf(mrs_Tmp1.Fields!idTipoServicio = 1, "CE", IIf(mrs_Tmp1.Fields!idTipoServicio = 3, "Hosp", "Emerg"))
                                rsTmpSOAT.Fields!Falta = IIf(IsNull(mrs_Tmp1.Fields!fechaEgreso), mrs_Tmp1.Fields!FEchaIngreso, mrs_Tmp1.Fields!fechaEgreso)
                                rsTmpSOAT.Fields!Halta = IIf(IsNull(mrs_Tmp1.Fields!horaEgreso), mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!horaEgreso)
                                rsTmpSOAT.Fields!dFinanciamiento = mrs_Tmp1.Fields!dFuenteFinanciamiento
                                rsTmpSOAT.Fields!PFarmacia = lnTFarmacia
                                rsTmpSOAT.Fields!PRx = lnTRx
                                rsTmpSOAT.Fields!PLab = lnTLaboratorio
                                rsTmpSOAT.Fields!POtros = lnTOtros
                                rsTmpSOAT.Fields!dFarmacia = lnTFarmaciaC
                                rsTmpSOAT.Fields!dRx = lnTRxC
                                rsTmpSOAT.Fields!dLab = lnTLaboratorioC
                                rsTmpSOAT.Fields!dOtros = lnTOtrosC
                                rsTmpSOAT.Fields!tFacturado = lnTotalFacturadoCuenta
                                rsTmpSOAT.Fields!tPagado = LnTotalPagadoCuenta
                                rsTmpSOAT.Fields!tPdctos = lnTDctos
                                rsTmpSOAT.Fields!tDeuda = lnTotalDeudaCuenta
                                If lbEsCuentaPagante = True And lnNroCptConsumidos = 1 And lbSoloPasoPorCita = True And mrs_Tmp1.Fields!idTipoServicio = 1 And lnTotalDeudaCuenta > 0 Then
                                   rsTmpSOAT.Fields!NoPagoCita = "Pagante, solo sacó cita, no pagó"
                                End If
                                If mrs_Tmp1!IdFormaPago = sghTipoFinanciamiento.sghSis Then
                                   rsTmpSOAT.Fields!fua = mo_ReglasAdmision.SisFuaAtencionDevuelveNumeroFUA(ml_IdCuenta)
                                End If
                                If mrs_Tmp.Fields!idCuentaAtencion > 0 Then
                                   If mrs_Tmp1.Fields!idTipoServicio = 1 Then
                                      rsTmpSOAT.Fields!Servicio = mo_ReglasAdmision.DevuelveServicioIngreso(mrs_Tmp!idCuentaAtencion, oConexion)
                                   Else
                                      rsTmpSOAT.Fields!Servicio = mo_ReglasAdmision.DevuelveServicioAlta(mrs_Tmp!idCuentaAtencion, oConexion)
                                   End If
                                End If
                                rsTmpSOAT.Update
                             End If
                         End If
                     End If
                     rsTmp.Close
                     'Añade IMPORTE TOTAL DE ESA CUENTA(200)
                     If lnTotalPorCuenta > 0 Then
                         mrs_Tmp.AddNew
                         mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                         mrs_Tmp.Fields!Paciente = lcTexto2
                         mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                         mrs_Tmp.Fields!IdPuntoCarga = 200
                         mrs_Tmp.Fields!dPuntoCarga = "Total (1) Consumos"
                         mrs_Tmp.Fields!consumo = lnTotalPorCuenta
                         mrs_Tmp.Update
                     End If
                     'Añade TOTAL PAGADO DE ESA CUENTA(203)
                     If LnTotalPagadoCuenta > 0 Then
                         mrs_Tmp.AddNew
                         mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                         mrs_Tmp.Fields!Paciente = lcTexto2
                         mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                         mrs_Tmp.Fields!IdPuntoCarga = 203
                         mrs_Tmp.Fields!dPuntoCarga = "Total (4) Pagado"
                         mrs_Tmp.Fields!consumo = LnTotalPagadoCuenta
                                                  
                         mrs_Tmp.Update
                     End If
                     '
                     'Añade TOTAL EXONERADO DE ESA CUENTA(202)
                     If lnTDctos > 0 Then
                         mrs_Tmp.AddNew
                         mrs_Tmp.Fields!nroHistoria = Val(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnSaldoInicial)), False))
                         mrs_Tmp.Fields!Paciente = lcTexto2
                         mrs_Tmp.Fields!idCuentaAtencion = ml_IdCuenta
                         mrs_Tmp.Fields!IdPuntoCarga = 202
                         mrs_Tmp.Fields!dPuntoCarga = "Total (3) Exoneraciones"
                         mrs_Tmp.Fields!consumo = lnTDctos
                         mrs_Tmp.Update
                     End If
                     mrs_Tmp1.MoveNext
                Loop
                mrs_Tmp1.Close
             End If
        Next
        oConexion.Close
        
        Exit Sub
ErrProcesar:
    MsgBox Err.Description
    Resume
End Sub















Sub GenerarRecordsetTemporal()
    If mrs_Tmp.State = 1 Then
       Set mrs_Tmp = Nothing
    End If
    With mrs_Tmp
         .Fields.Append "NroHistoria", adInteger
         .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
         .Fields.Append "idCuentaAtencion", adInteger
         .Fields.Append "idPuntoCarga", adInteger
         .Fields.Append "dPuntoCarga", adVarChar, 40, adFldIsNullable
         .Fields.Append "Consumo", adDouble
         .Fields.Append "FechaEgreso", adDate
         .Fields.Append "horaEgreso", adVarChar, 5
         .Fields.Append "dFuenteFinanciamiento", adVarChar, 100
         .Fields.Append "idTipoServicio", adInteger
         .LockType = adLockOptimistic
         .Open
    End With
End Sub

Sub ReporteHRACreaTmp(mb_EnResumen As Boolean)
    On Error GoTo errHraCreaTmp
    If mb_EnResumen = True Then
        If rsTmpSOAT.State = 1 Then
           Set rsTmpSOAT = Nothing
        End If
        With rsTmpSOAT
                .Fields.Append "idCuentaAtencion", adInteger
                .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
                .Fields.Append "NroHistoria", adInteger
                .Fields.Append "Origen", adVarChar, 10, adFldIsNullable
                .Fields.Append "Falta", adDate, 10, adFldIsNullable
                .Fields.Append "Halta", adVarChar, 5, adFldIsNullable
                .Fields.Append "dFinanciamiento", adVarChar, 100, adFldIsNullable
                .Fields.Append "tFacturado", adDouble
                .Fields.Append "tPagado", adDouble
                .Fields.Append "tPdctos", adDouble
                .Fields.Append "tDeuda", adDouble
                .Fields.Append "PFarmacia", adDouble
                .Fields.Append "PRx", adDouble
                .Fields.Append "PLab", adDouble
                .Fields.Append "POtros", adDouble
                .Fields.Append "DFarmacia", adDouble
                .Fields.Append "DRx", adDouble
                .Fields.Append "DLab", adDouble
                .Fields.Append "DOtros", adDouble
                .Fields.Append "NoPagoCita", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "FUA", adVarChar, 25, adFldIsNullable
                .Fields.Append "idTipoFinanciamiento", adInteger
                .LockType = adLockOptimistic
                .Open
        End With
    End If
    Exit Sub
errHraCreaTmp:
    MsgBox Err.Description
    Resume
End Sub
