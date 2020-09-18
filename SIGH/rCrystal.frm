VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rCrystal 
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "rCrystal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CrvReportes 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      lastProp        =   500
      _cx             =   15690
      _cy             =   9869
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Procesa y Muestra varios Reportes
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

'aqui declara los objetos que contendra al rporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Dim lcParametro251 As String, lcParametro202 As String, lcParametro252 As String, lcParametro269 As String
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReporteUtil As New ReporteUtil
Dim lc_TextoDelFiltro As String
Dim lc_TipoReporte As String
Dim lnIdPuntoCarga As Long
Dim lnOrdenadoPor As Long
Dim mrs_Tmp As New Recordset
Dim mrs_Tmp1 As New Recordset
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_HoraInicio As String
Dim ml_HoraFin As String
Dim mb_ConsiderarSinMovimientos As Boolean
Dim mb_SeMuestraLotes As Boolean
Dim mb_StockMinimoMayorAcantidad As Boolean
Dim ml_idUsuario As Long
Dim ml_IdProducto  As Long
Dim lnIdPuntoCargaOrigen As Long
Dim lnIdPuntoCargaDestino As Long
Dim ml_IdConcepto As Long
Dim ml_MovTipo As String
Dim ml_IdEstado  As Long
Dim lc_AlmacenesParaICI As String
Dim ml_IdAnio As Long
Dim ml_IdCuenta As Long
Dim ml_Dias  As Long
Dim ml_Almacen As String
Dim ml_Documento As String
Dim ml_AlmacenO As String
Dim ml_Importe As Double
Dim mb_ConsiderarRecalculo As Boolean
Dim mb_EnArchivoExcel As Boolean
Dim lnIdResponsable As Long
Dim mb_EnResumen As Boolean
Dim ln_DestinoReporte As sghImpresion
Dim ml_lcTipoServicio As String
Dim ml_LcIdCuentaAtencion As String
Dim ml_lcFormaPago As String
Dim ml_lcNroCola As String
Dim ml_lcServicioDelTarifario As String
Dim ml_idAtencion As Long
Dim ml_IdCajero As Long
Dim ml_IdTurno As Long
Dim ml_IdCaja As Long
Dim ml_IdCentroCostos As Long
Dim moConexion As New ADODB.Connection
Dim mo_ProgressRpt As XP_ProgressBar
Dim mo_ProgressRpt1 As XP_ProgressBar
Dim mb_ConProrrateoColExoneracion As Boolean
Dim mb_TotalizarXconsultorio As Boolean
Dim lcServicioPaciente As String
Dim ml_IncluyeHistoriasQueSalieron As Boolean
Dim mda_FechaSolicitudHasta As Date
Dim mda_FechaSolicitudDesde As Date
Dim mb_ConOtrosSaludDesagregado As Boolean
Dim mb_DetallaProcAdmyOtrosServ As Boolean
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String
Dim mb_tieneCredito As Boolean

Property Let TieneCredito(lValue As Boolean)
    mb_tieneCredito = lValue
End Property
Property Let DetallaProcAdmyOtrosServ(lValue As Boolean)
    mb_DetallaProcAdmyOtrosServ = lValue
End Property


Property Let ConOtrosSaludDesagregado(lValue As Boolean)
    mb_ConOtrosSaludDesagregado = lValue
End Property

Property Let ConProrrateoColExoneracion(lValue As Boolean)
    mb_ConProrrateoColExoneracion = lValue
End Property

Property Set progressRpt1(oValue As XP_ProgressBar)
    Set mo_ProgressRpt1 = oValue
End Property
Property Set progressRpt(oValue As XP_ProgressBar)
    Set mo_ProgressRpt = oValue
End Property


Property Let idCentroCostos(lValue As Long)
    ml_IdCentroCostos = lValue
End Property
Property Let IdCaja(lValue As Long)
    ml_IdCaja = lValue
End Property
Property Let IdTurno(lValue As Long)
    ml_IdTurno = lValue
End Property
Property Let IdCajero(lValue As Long)
    ml_IdCajero = lValue
End Property
Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property
Property Let lcServicioDelTarifario(lValue As String)
    ml_lcServicioDelTarifario = lValue
End Property
Property Let lcNroCola(lValue As String)
    ml_lcNroCola = lValue
End Property

Property Let lcFormaPago(lValue As String)
    ml_lcFormaPago = lValue
End Property
Property Let LcIdCuentaAtencion(lValue As String)
    ml_LcIdCuentaAtencion = lValue
End Property

Property Let lcTipoServicio(lValue As String)
    ml_lcTipoServicio = lValue
End Property


Property Let DestinoReporte(lValue As sghImpresion)
    ln_DestinoReporte = lValue
End Property

Property Let EnResumen(lValue As Boolean)
    mb_EnResumen = lValue
End Property

Property Let EnArchivoExcel(lValue As Boolean)
    mb_EnArchivoExcel = lValue
End Property
Property Let ConsiderarRecalculo(lValue As Boolean)
    mb_ConsiderarRecalculo = lValue
End Property

Property Let Importe(lValue As Double)
    ml_Importe = lValue
End Property
Property Let AlmacenO(lValue As String)
    ml_AlmacenO = lValue
End Property
Property Let Documento(lValue As String)
    ml_Documento = lValue
End Property
Property Let Almacen(lValue As String)
    ml_Almacen = lValue
End Property

Property Let Dias(lValue As Long)
    ml_Dias = lValue
End Property

Property Let IdCuenta(lValue As Long)
    ml_IdCuenta = lValue
End Property
Property Let IdAnio(lValue As Long)
    ml_IdAnio = lValue
End Property

Property Let AlmacenesParaICI(lValue As String)
    lc_AlmacenesParaICI = lValue
End Property

Property Let estado(lValue As Long)
    ml_IdEstado = lValue
End Property

Property Let Concepto(lValue As Long)
    ml_IdConcepto = lValue
End Property
Property Let MovTipo(lValue As String)
    ml_MovTipo = lValue
End Property


Property Let IdPuntoCargaDestino(iValue As Long)
   lnIdPuntoCargaDestino = iValue
End Property
Property Let IdPuntoCargaOrigen(iValue As Long)
   lnIdPuntoCargaOrigen = iValue
End Property

Property Let idProducto(lValue As Long)
    ml_IdProducto = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let StockMinimoMayorAcantidad(lValue As Boolean)
    mb_StockMinimoMayorAcantidad = lValue
End Property
Property Let SeMuestraLotes(lValue As Boolean)
    mb_SeMuestraLotes = lValue
End Property
Property Let ConsiderarSinMovimientos(lValue As Boolean)
    mb_ConsiderarSinMovimientos = lValue
End Property

Property Let HoraInicio(lValue As String)
    ml_HoraInicio = lValue
End Property
Property Let HoraFin(lValue As String)
    ml_HoraFin = lValue
End Property
Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Property Let idPuntoCarga(iValue As Long)
   lnIdPuntoCarga = iValue
End Property
Property Let OrdenadoPor(iValue As Long)
   lnOrdenadoPor = iValue
End Property

Property Let TipoReporte(iValue As String)
   lc_TipoReporte = iValue
End Property
Property Let TextoDelFiltro(iValue As String)
   lc_TextoDelFiltro = iValue
End Property
Property Let idResponsable(iValue As Long)
   lnIdResponsable = iValue
End Property
Property Let TotalizarXconsultorio(lValue As Boolean)
    mb_TotalizarXconsultorio = lValue
End Property

Property Let IncluyeHistoriasQueSalieron(lValue As Boolean)
    ml_IncluyeHistoriasQueSalieron = lValue
End Property

Property Let FechaSolicitudDesde(lValue As Date)
    mda_FechaSolicitudDesde = lValue
End Property
Property Let FechaSolicitudHasta(lValue As Date)
    mda_FechaSolicitudHasta = lValue
End Property

Private Sub Form_Activate()
       If ln_DestinoReporte <> sghPantalla Then
            Me.Visible = False
       End If

End Sub

Private Sub Form_Load()
    If Len(lc_TextoDelFiltro) > 250 Then
       lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
    End If

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    Dim rsreporte As New ADODB.Recordset
    Dim rsTmp As New Recordset
    Dim rsTmp1 As New Recordset
    Dim oRsExoneradoBoleta As New Recordset
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
    Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
    Dim ldFechaPrincipio As Date
    Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim oConexion As New ADODB.Connection
    Dim lbPrimeraVez As Boolean: Dim lbEncontroDato As Boolean
    Dim LcTexto1 As String: Dim LcTexto2 As String: Dim lcTexto3 As String
    Dim lnId1 As Long: Dim lnId2 As Long: Dim lnId3 As Long
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim lnPagoCta As Double, lnAnulado As Double, lnImptotal As Double
    Dim lnCanTotal As Long, lnCanAnulado As Long, lnCanExoneracion As Long
    Dim lnExoneracion As Double
    Dim lnIdTipoServicio As Long, lnIdCentroCosto As Long, idProducto As Long
    Dim lnCama As Double, lnSoloSOP As Double, lnResto As Double
    Dim lnAnuladoCama As Double, lnExoneracionCama As Double, lnImptotalCama As Double
    Dim lnAnuladoSOP As Double, lnExoneracionSOP As Double, lnImptotalSOP As Double
    Dim lnAnuladoResto As Double, lnExoneracionResto As Double, lnImptotalResto As Double
    Dim lnConsulta As Double, lnAnuladoConsulta As Double, lnExoneracionConsulta As Double
    Dim lnImptotalConsulta As Double, lbEncontroSop As Boolean
    Dim lnQueda As Double
    Dim lnLineas As Long, lRecordCount As Long
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    mflgContinuar = False
    lcTexto3 = ""
    lcParametro202 = lcBuscaParametro.SeleccionaFilaParametro(202)
    lcParametro252 = lcBuscaParametro.SeleccionaFilaParametro(252)
    lcParametro251 = lcBuscaParametro.SeleccionaFilaParametro(251)
    lcParametro269 = lcBuscaParametro.SeleccionaFilaParametro(269)
    lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
    lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    
    Select Case lc_TipoReporte
    Case "ImpresionPreCuenta"
        Set rsreporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
        If rsreporte.RecordCount = 0 Then
            MsgBox "No existen datos", vbInformation, "Reporte"
        Else
            GenerarRecordsetTemporal lc_TipoReporte
            
            mrs_Tmp.AddNew
            mrs_Tmp.Fields!FechaIngreso = "Fecha: " & rsreporte.Fields!FechaIngreso & "     Hora: " & rsreporte.Fields!HoraIngreso
            mrs_Tmp.Fields!Paciente = "Paciente: " & Trim(rsreporte.Fields!ApellidoPaterno) & " " & Trim(rsreporte.Fields!ApellidoMaterno) & " " & Trim(rsreporte.Fields!PrimerNombre) & " " & mo_ReporteUtil.NullToVacio(rsreporte!SegundoNombre)
            mrs_Tmp.Fields!Usuario = "Usuario: " & mo_ReglasCaja.SeleccionaDatosCajero(ml_idUsuario, sghUsuario) & "     No Cuenta: " & ml_LcIdCuentaAtencion
            mrs_Tmp.Fields!NroHistoriaClinica = "No Historia: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False)
            mrs_Tmp.Fields!Servicio = "Servicio: " & mo_ReporteUtil.NullToVacio(rsreporte!Servicio)
            mrs_Tmp.Fields!Medico = "Médico: " & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsreporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsreporte!NombresMedico))
            If ml_lcNroCola = "" Then
                mrs_Tmp.Fields!Interconsulta = ""
                mrs_Tmp.Fields!ColaTipoS = ml_lcTipoServicio
            Else
                mrs_Tmp.Fields!Interconsulta = "Interconsulta:           Si (   )             No (  )"
                mrs_Tmp.Fields!ColaTipoS = "Cola: " & ml_lcNroCola
            End If
            mrs_Tmp.Update
            'Reporte
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\CePreCuenta.rpt", 1)
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                    Case "TipoServicio"
                        crParamDef.AddCurrentValue (ml_lcTipoServicio)
                    Case "LcIdCuentaAtencion"
                        crParamDef.AddCurrentValue (ml_LcIdCuentaAtencion)
                    Case "lcFormaPago"
                        crParamDef.AddCurrentValue (ml_lcFormaPago)
                    Case "lcNroCola"
                        crParamDef.AddCurrentValue (ml_lcNroCola)
                    Case "lcServicioDelTarifario"
                        crParamDef.AddCurrentValue (ml_lcServicioDelTarifario)
                End Select
            Next
            crReport.Database.SetDataSource mrs_Tmp
        End If
    Case "ResumenCCosto"
        moConexion.CursorLocation = adUseClient
        moConexion.CommandTimeout = 300
        moConexion.Open sighEntidades.CadenaConexion
        
        If mb_ConProrrateoColExoneracion = False Then
           ResumenXcentroCosto mrs_Tmp
           lc_TextoDelFiltro = lc_TextoDelFiltro & " (s.p.)"
           If mb_DetallaProcAdmyOtrosServ = True Then
              lc_TextoDelFiltro = lc_TextoDelFiltro & " (PA.OS.desag.)"
           ElseIf mb_ConOtrosSaludDesagregado = True Then
              lc_TextoDelFiltro = lc_TextoDelFiltro & " (O.S.desag.)"
           End If
        Else
            '
            
            lcTexto3 = "..comienza a generar Temporal..."
            GenerarRecordsetTemporal lc_TipoReporte
            lcTexto3 = "..llena Temporal..."
            'Llena Temporal con Centro de Costos
            Set rsTmp = mo_ReglasComunes.CentrosCostoSeleccionarTodos(moConexion)
            If rsTmp.RecordCount = 0 Then
               MsgBox "Llene tabla CENTRO DE COSTOS"
               Exit Sub
            End If
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!IdCentroCosto = rsTmp.Fields!IdCentroCosto
               mrs_Tmp.Fields!CentroCosto = rsTmp.Fields!descripcion
               mrs_Tmp.Update
               rsTmp.MoveNext
            Loop
            rsTmp.Close
            lcTexto3 = "..comienza a jalar boletas servicios..."
            'Boletas Servicios
            Set rsreporte = mo_ReglasFacturacion.ServicioConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
            lRecordCount = rsreporte.RecordCount
            If lRecordCount > 0 Then
               mo_ProgressRpt.Min = 0
               mo_ProgressRpt.Max = lRecordCount
               mo_ProgressRpt.Value = 0
               mo_ProgressRpt.ShowText = True
               mo_ProgressRpt.Color = vbGreen
               lnLineas = 0
               '
               rsreporte.MoveFirst
               Do While Not rsreporte.EOF
                  LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
                  lnPagoCta = rsreporte.Fields!Adelantos
                  lnExoneracion = rsreporte.Fields!exoneraciones
                  If rsreporte.Fields!idEstadoComprobante = 9 Then
                     lnAnulado = rsreporte.Fields!TotalPagado
                     lnImptotal = 0
                  Else
                     lnAnulado = 0
                     lnImptotal = rsreporte.Fields!TotalPagado
                  End If
                  '
                  lnIdTipoServicio = 0
                  If rsreporte.Fields!idCuentaAtencion > 0 Then
                    lcTexto3 = "..comienza a jalar tipo servicios..."
                    Set rsTmp = mo_ReglasAdmision.FacturacionCuentasAtencionXnroCuentaConexion(rsreporte.Fields!idCuentaAtencion, moConexion)
                    If rsTmp.RecordCount > 0 Then
                       lnIdTipoServicio = rsTmp.Fields!idTipoServicio
                    End If
                    rsTmp.Close
                  End If
                  lbPrimeraVez = True
                  Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                        lnLineas = lnLineas + 1
                        mo_ProgressRpt.Value = lnLineas
                        '
                        If lbPrimeraVez = True Then
                           lbPrimeraVez = False
                           lbEncontroDato = False
                           lnIdCentroCosto = 0
                           '*****Consulta Externa (menos Medicina Fisica y Rehabilitacion)
                           If lbEncontroDato = False And lnIdTipoServicio = 1 Then
                                lcTexto3 = "..comienza: Consulta Externa (menos Medicina Fisica y Rehabilitacion)..."
                                Set rsTmp = mo_ReglasCaja.FactOrdenServicioXidOrden(rsreporte.Fields!IdOrden, moConexion)
                                If rsTmp.RecordCount > 0 Then
                                    If rsTmp.Fields!idPuntoCarga = 6 And rsTmp.Fields!idServicioPaciente <> 68 Then
                                       lnIdCentroCosto = 1003
                                       lbEncontroDato = True
                                    End If
                                End If
                                rsTmp.Close
                           End If
                           '*****Consulta Externa (solo Medicina Fisica y Rehabilitacion)
                           If lbEncontroDato = False And lnIdTipoServicio = 1 Then
                                lcTexto3 = "..comienza: Consulta Externa (solo Medicina Fisica y Rehabilitacion)..."
                                Set rsTmp = mo_ReglasCaja.FactOrdenServicioXidOrden(rsreporte.Fields!IdOrden, moConexion)
                                If rsTmp.RecordCount > 0 Then
                                    If rsTmp.Fields!idPuntoCarga = 6 And rsTmp.Fields!idServicioPaciente = 68 Then
                                       lnIdCentroCosto = 1004
                                       lbEncontroDato = True
                                    End If
                                End If
                                rsTmp.Close
                           End If
                           '*****Hospitalizacion
                           If lbEncontroDato = False And lnIdTipoServicio = 3 Then
                              lnCama = 0: lnSoloSOP = 0: lnResto = 0
                              lcTexto3 = "..comienza: Hospitalizacion..."
                              Set rsTmp = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumento(rsreporte.Fields!nroSerie, rsreporte.Fields!nrodocumento, moConexion)
                              If rsTmp.RecordCount > 0 Then
                                 '
                                 Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsreporte.Fields!IdComprobantePago, moConexion)
                                 '
                                 lbEncontroDato = True
                                 rsTmp.MoveFirst
                                 Do While Not rsTmp.EOF
                                    If rsTmp.Fields!idProducto = Val(lcParametro202) Then
                                        'id Estancia (tabla parametros)
                                        lnCama = lnCama + rsTmp.Fields!Total
                                    Else
                                        lcTexto3 = "..comienza: Hospitalizacion-NO cama..."
                                        Set rsTmp1 = mo_ReglasFacturacion.FactOrdenServicioXidOrdenConexion(rsTmp.Fields!IdOrden, moConexion)
                                        If rsTmp1.RecordCount > 0 Then
                                            If rsTmp1.Fields!idPuntoCarga = 573 Then
                                               lnSoloSOP = lnSoloSOP + rsTmp.Fields!Total
                                            Else
                                               lnResto = lnResto + rsTmp.Fields!Total
                                            End If
                                        End If
                                        rsTmp1.Close
                                    End If
                                    rsTmp.MoveNext
                                 Loop
                                 '
                                 oRsExoneradoBoleta.Close
                              End If
                              rsTmp.Close
                              'comprueba que cuadre totales y subTotales
                              If (lnCama + lnSoloSOP + lnResto) > rsreporte.Fields!TotalPagado Then
                                     lnQueda = ((lnCama + lnSoloSOP + lnResto) - rsreporte.Fields!TotalPagado)
                                     If (lnCama - lnQueda) >= 0 Then
                                         lnCama = lnCama - lnQueda
                                     Else
                                         lnQueda = lnQueda - lnCama
                                         lnCama = 0
                                         If (lnSoloSOP - lnQueda) >= 0 Then
                                             lnSoloSOP = lnSoloSOP - lnQueda
                                         Else
                                             lnQueda = lnQueda - lnSoloSOP
                                             lnSoloSOP = 0
                                             lnResto = lnResto - lnQueda
                                         End If
                                     End If
                              ElseIf (lnCama + lnSoloSOP + lnResto) < rsreporte.Fields!TotalPagado Then
                                     lnCama = lnCama + (rsreporte.Fields!TotalPagado - (lnCama + lnSoloSOP + lnResto))
                              End If
                              '
                              If rsreporte.Fields!idEstadoComprobante <> 9 Then
                                    If rsreporte.Fields!exoneraciones > 0 Then
                                        lnAnuladoCama = 0
                                        lnExoneracionCama = (rsreporte.Fields!exoneraciones / 3)
                                        lnImptotalCama = lnCama
                                        '
                                        lnAnuladoSOP = 0
                                        lnExoneracionSOP = (rsreporte.Fields!exoneraciones / 3)
                                        lnImptotalSOP = lnSoloSOP
                                        '
                                        lnAnuladoResto = 0
                                        lnExoneracionResto = (rsreporte.Fields!exoneraciones / 3)
                                        lnImptotalResto = lnResto
                                    Else
                                        lnAnuladoCama = 0
                                        lnExoneracionCama = 0
                                        lnImptotalCama = lnCama
                                        '
                                        lnAnuladoSOP = 0
                                        lnExoneracionSOP = 0
                                        lnImptotalSOP = lnSoloSOP
                                        '
                                        lnAnuladoResto = 0
                                        lnExoneracionResto = 0
                                        lnImptotalResto = lnResto
                                    End If
                              Else
                                    lnAnuladoCama = lnCama
                                    lnExoneracionCama = 0
                                    lnImptotalCama = 0
                                    '
                                    lnAnuladoSOP = lnSoloSOP
                                    lnExoneracionSOP = 0
                                    lnImptotalSOP = 0
                                    '
                                    lnAnuladoResto = lnResto
                                    lnExoneracionResto = 0
                                    lnImptotalResto = 0
                              End If
                              'Hospitalizacion-solo cama
                              If lnCama <> 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1005"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoCama
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionCama
                                      mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalCama
                                      mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalCama
                                      mrs_Tmp.Update
                                    End If
                              ElseIf lnExoneracionCama > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1005"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionCama
                                      mrs_Tmp.Update
                                    End If
                              End If
                              'Hospitalizacion-solo SOP
                              If lnSoloSOP <> 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1006"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoSOP
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionSOP
                                      mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalSOP
                                      mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalSOP
                                      mrs_Tmp.Update
                                    End If
                              ElseIf lnExoneracionSOP > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1006"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionSOP
                                      mrs_Tmp.Update
                                    End If
                              End If
                              'Hospitalizacion-resto (sin SOP, sin CAMA)
                              If lnResto <> 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1008"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoResto
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionResto
                                      mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalResto
                                      mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalResto
                                      mrs_Tmp.Update
                                    End If
                              ElseIf lnExoneracionResto > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1008"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionResto
                                      mrs_Tmp.Update
                                    End If
                              End If
                           End If
                           '*****Emergencia
                           lcTexto3 = "..comienza: Emergencia..."
                           If lbEncontroDato = False And (lnIdTipoServicio = 2 Or lnIdTipoServicio = 4) Then
                              lnConsulta = 0: lnResto = 0
                              Set rsTmp = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumento(rsreporte.Fields!nroSerie, rsreporte.Fields!nrodocumento, moConexion)
                              If rsTmp.RecordCount > 0 Then
                                 '
                                 Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsreporte.Fields!IdComprobantePago, moConexion)
                                 '
                                 lbEncontroDato = True
                                 rsTmp.MoveFirst
                                 Do While Not rsTmp.EOF
                                    If rsTmp.Fields!idProducto = mo_ReglasFacturacion.ObtenerCodigoDeConsultaDeEmergencia() Then
                                       'id consulta de emergencia (tabla parametros)
                                        lnConsulta = lnConsulta + rsTmp.Fields!Total
                                    Else
                                        lnResto = lnResto + rsTmp.Fields!Total
                                    End If
                                    rsTmp.MoveNext
                                 Loop
                                 '
                                 oRsExoneradoBoleta.Close
                              End If
                              rsTmp.Close
                              '
                              If (lnConsulta + lnResto) > rsreporte.Fields!TotalPagado Then
                                   lnQueda = ((lnConsulta + lnResto) - rsreporte.Fields!TotalPagado)
                                   If (lnConsulta - lnQueda) >= 0 Then
                                       lnConsulta = lnConsulta - lnQueda
                                   Else
                                       lnQueda = lnQueda - lnConsulta
                                       lnConsulta = 0
                                       lnResto = lnResto - lnQueda
                                   End If
                              ElseIf (lnCama + lnSoloSOP + lnResto) < rsreporte.Fields!TotalPagado Then
                                   lnConsulta = lnConsulta + (rsreporte.Fields!TotalPagado - (lnConsulta + lnResto))
                              End If
                              If rsreporte.Fields!idEstadoComprobante <> 9 Then
                                    If rsreporte.Fields!exoneraciones > 0 Then
                                        lnAnuladoConsulta = 0
                                        lnExoneracionConsulta = (rsreporte.Fields!exoneraciones / 2)
                                        lnImptotalConsulta = lnConsulta
                                        '
                                        lnAnuladoResto = 0
                                        lnExoneracionResto = (rsreporte.Fields!exoneraciones / 2)
                                        lnImptotalResto = lnResto
                                    Else
                                        lnAnuladoConsulta = 0
                                        lnExoneracionConsulta = 0
                                        lnImptotalConsulta = lnConsulta
                                        '
                                        lnAnuladoResto = 0
                                        lnExoneracionResto = 0
                                        lnImptotalResto = lnResto
                                    End If
                              Else
                                    lnAnuladoConsulta = lnConsulta
                                    lnExoneracionConsulta = 0
                                    lnImptotalConsulta = 0
                                    '
                                    lnAnuladoResto = lnResto
                                    lnExoneracionResto = 0
                                    lnImptotalResto = 0
                              End If
                              'Emergencia-solo CONSULTA
                              If lnConsulta > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1013"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                      mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                      mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                      mrs_Tmp.Update
                                    End If
                              ElseIf lnExoneracionConsulta > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1013"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                      mrs_Tmp.Update
                                    End If
                              End If
                              'Emergencia-resto (sin CONSULTA)
                              If lnResto > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1010"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoResto
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionResto
                                      mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalResto
                                      mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalResto
                                      mrs_Tmp.Update
                                    End If
                              ElseIf lnExoneracionResto > 0 Then
                                    mrs_Tmp.MoveFirst
                                    mrs_Tmp.Find "idCentroCosto=1010"
                                    If Not mrs_Tmp.EOF Then
                                      mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionResto
                                      mrs_Tmp.Update
                                    End If
                              End If
                           End If
                           '*****Reembolsos: Farmacia y/o Servicios
                           If lbEncontroDato = False Then
                              If lnAnulado = 0 Then
                                    lcTexto3 = "..comienza: Reembolsos: Farmacia y/o Servicios..."
                                    Set rsTmp = mo_ReglasFacturacion.FacturacionReembolsosXboletaConexion(rsreporte.Fields!IdComprobantePago, moConexion)
                                    If rsTmp.RecordCount > 0 Then
                                        lnResto = 0   'total reembolso farmacia
                                        Do While Not rsTmp.EOF
                                           lnResto = lnResto + rsTmp.Fields!ReembolsoPagadoFarmacia
                                           rsTmp.MoveNext
                                        Loop
                                        '*****Farmacia
                                        If lnResto > 0 Then
                                            lnExoneracionConsulta = 0
                                            If lnAnulado = 0 Then
                                               lnImptotalConsulta = lnResto
                                               lnAnuladoConsulta = 0
                                            Else
                                               lnImptotalConsulta = 0
                                               lnAnuladoConsulta = lnResto
                                            End If
                                            mrs_Tmp.MoveFirst
                                            mrs_Tmp.Find "idCentroCosto=1009"
                                            If Not mrs_Tmp.EOF Then
                                              mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                              mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                              mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                              mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                              mrs_Tmp.Update
                                              lbEncontroDato = True
                                              lnIdCentroCosto = 0
                                            End If
                                        End If
                                        '*****Servicio por REEMBOLSO
                                        lnResto = rsreporte.Fields!TotalPagado - lnResto
                                        If lnResto > 0 Then
                                            lnExoneracionConsulta = 0
                                            If lnAnulado = 0 Then
                                               lnImptotalConsulta = lnResto
                                               lnAnuladoConsulta = 0
                                            Else
                                               lnImptotalConsulta = 0
                                               lnAnuladoConsulta = lnResto
                                            End If
                                            mrs_Tmp.MoveFirst
                                            mrs_Tmp.Find "idCentroCosto=1015"
                                            If Not mrs_Tmp.EOF Then
                                              mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                              mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                              mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                              mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                              mrs_Tmp.Update
                                              lbEncontroDato = True
                                              lnIdCentroCosto = 0
                                            End If
                                        End If
                                    End If
                                    rsTmp.Close
                                End If
                           End If
                           '*****Laboratorio
                           If lbEncontroDato = False Then
                                lcTexto3 = "..comienza: Laboratorio..."
                                Set rsTmp = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & rsreporte.Fields!idProducto, moConexion)
                                If rsTmp.RecordCount > 0 Then
                                If rsTmp.Fields!idPuntoCarga = 2 Or rsTmp.Fields!idPuntoCarga = 3 Then
                                   lnIdCentroCosto = 1001
                                   lbEncontroDato = True
                                End If
                                End If
                                rsTmp.Close
                           End If
                           '*****Imagenes
                           If lbEncontroDato = False Then
                                lcTexto3 = "..comienza: Imagenes..."
                                Set rsTmp = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & rsreporte.Fields!idProducto, moConexion)
                                If rsTmp.RecordCount > 0 Then
                                If rsTmp.Fields!idPuntoCarga >= 20 And rsTmp.Fields!idPuntoCarga <= 23 Then
                                   lnIdCentroCosto = 1002
                                   lbEncontroDato = True
                                End If
                                End If
                                rsTmp.Close
                           End If
                           '*****Procedimientos Administrativos
                           If lbEncontroDato = False Then
                                If rsreporte.Fields!IdServicioGrupo = 5 Then
                                   lnIdCentroCosto = 1011
                                   lbEncontroDato = True
                                End If
                           End If
                           '*****Otros SALUD
                           If lbEncontroDato = False Then
                              lnIdCentroCosto = 999
                           End If
                        End If
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
                  Loop
                  If lnIdCentroCosto > 0 Then
                        mrs_Tmp.MoveFirst
                        mrs_Tmp.Find "idCentroCosto=" & lnIdCentroCosto
                        If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
                          mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotal
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotal
                          mrs_Tmp.Update
                        End If
                  End If
                  lcTexto3 = "..comienza: termino loop..."
               Loop
            End If
            'MEDICAMENTOS emitidos en CAJA SERVICIO
            lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
            Set rsreporte = Nothing
            Set rsreporte = mo_ReglasFacturacion.FarmaciaConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
            lRecordCount = rsreporte.RecordCount
            If lRecordCount > 0 Then
               mo_ProgressRpt1.Min = 0
               mo_ProgressRpt1.Max = lRecordCount
               mo_ProgressRpt1.Value = 0
               mo_ProgressRpt1.ShowText = True
               mo_ProgressRpt1.Color = vbGreen
               lnLineas = 0
               '
               rsreporte.MoveFirst
               lnExoneracion = 0: lnAnulado = 0: lnImptotal = 0
               Do While Not rsreporte.EOF
                  LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
                  lnExoneracion = lnExoneracion + rsreporte.Fields!exoneraciones
                  If rsreporte.Fields!idEstadoComprobante = 9 Then
                     lnAnulado = lnAnulado + rsreporte.Fields!TotalPagado
                  Else
                     lnImptotal = lnImptotal + rsreporte.Fields!TotalPagado
                  End If
                  Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                        lnLineas = lnLineas + 1
                        mo_ProgressRpt1.Value = lnLineas
                        '
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
                  Loop
               Loop
               mrs_Tmp.MoveFirst
               mrs_Tmp.Find "idCentroCosto=1009"  'Farmacia
               mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
               mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
               mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotal
               mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotal
               mrs_Tmp.Update
            End If
        End If
        lcTexto3 = "..comienza: emision reporte..."
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe informacion", vbInformation, "Consolidado Centro Costos"
        Else
            'Reporte
            CrvReportes.EnableExportButton = False
            mrs_Tmp.Sort = "CentroCosto"
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\EconResumenCCosto.rpt", 1)
            
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                Case "subTitulo"
                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
'                Case "FechaHoraImpresion"
'                    crParamDef.AddCurrentValue (lcBuscaParametro.RetornaFechaHoraServidorSQL)
                Case "fecha"
                  crParamDef.AddCurrentValue ("Fecha Impresión: " & lcBuscaParametro.RetornaFechaServidorSQL)
                Case "hora"
                  crParamDef.AddCurrentValue ("Hora Impresión: " & lcBuscaParametro.RetornaHoraServidorSQL1)
                Case "pc"
                    crParamDef.AddCurrentValue ("PC: " & sighEntidades.RetornaNombrePC)
                Case "user"
                    crParamDef.AddCurrentValue ("Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighEntidades.Usuario))
                Case "lcEESS"
                    crParamDef.AddCurrentValue (lcTitEESS)
                Case "lcEESSdireccion"
                    crParamDef.AddCurrentValue (lcTitDireccion)
                Case "lcEESStelefono"
                    crParamDef.AddCurrentValue (lcTitTelefono)
                End Select
            Next
            crReport.Database.SetDataSource mrs_Tmp
        End If
        moConexion.Close
    Case "DetalleCentroCosto"
        moConexion.Open sighEntidades.CadenaConexion
        moConexion.CursorLocation = adUseClient
        ResumenXcentroCosto mrs_Tmp
        
        If mb_DetallaProcAdmyOtrosServ = True Then
           lc_TextoDelFiltro = lc_TextoDelFiltro & " (PA.OS.desag.)"
        ElseIf mb_ConOtrosSaludDesagregado = True Then
           lc_TextoDelFiltro = lc_TextoDelFiltro & " (O.S.desag.)"
        End If
        moConexion.Close
        If mrs_Tmp.RecordCount > 0 Then
            mrs_Tmp.Sort = "producto"
            CrvReportes.EnableExportButton = False
            'Reporte
            mflgContinuar = True
            If mb_TotalizarXconsultorio = True Then
               Set crReport = crApp.OpenReport(App.Path & "\plantillas\EconDetalleCentroCostoS.rpt", 1)
            Else
               Set crReport = crApp.OpenReport(App.Path & "\plantillas\EconDetalleCentroCosto.rpt", 1)
            End If
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                Case "subTitulo"
                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                Case "FechaHoraImpresion"
                    crParamDef.AddCurrentValue (lcBuscaParametro.RetornaFechaHoraServidorSQL)
                Case "lcEESS"
                    crParamDef.AddCurrentValue (lcTitEESS)
                Case "lcEESSdireccion"
                    crParamDef.AddCurrentValue (lcTitDireccion)
                Case "lcEESStelefono"
                    crParamDef.AddCurrentValue (lcTitTelefono)
                End Select
            Next
            crReport.Database.SetDataSource mrs_Tmp
        End If
    Case "HcXmedicoXpagina"
    End Select
    lcTexto3 = "..comienza: salida de reporte..."
    If mflgContinuar = True Then
       Select Case ln_DestinoReporte
       Case sghImpresora
            crReport.PrintOut
       Case sghPantalla
            CrvReportes.ReportSource = crReport
            CrvReportes.ViewReport
            CrvReportes.Zoom 120
       Case sghExcel
            crReport.ExportOptions.DestinationType = crEDTDiskFile
            crReport.ExportOptions.FormatType = crEFTExcel70
            crReport.ExportOptions.DiskFileName = lcParametro269
            crReport.Export (False)
            MsgBox "Se generó el archivo " & lcParametro269
       End Select
 '
       mo_ReglasComunes.grabaTablaAuditoria (crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))   'debb-27/05/2015
          
    End If
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    Set oConexion = Nothing
    Set mo_ReglasImagenes = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set rsreporte = Nothing
    Set rsTmp = Nothing
    Set oRsExoneradoBoleta = Nothing
    Set oDoImagMovimientoIngresos = Nothing
    Set oDoImagMovimiento = Nothing
    LimpiarVariablesDeMemoria
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", _
            vbInformation + vbOKOnly
    Else
        MsgBox Err.Description & Chr(13) & lcTexto3, vbInformation + vbOKOnly
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
    LimpiarVariablesDeMemoria
End Sub





Private Sub Form_Resize()
    CrvReportes.Top = 0
    CrvReportes.Left = 0
    CrvReportes.Height = ScaleHeight
    CrvReportes.Width = ScaleWidth
End Sub

Sub GenerarRecordsetTemporal(lcReporte As String)
    With mrs_Tmp
         Select Case lcReporte
         Case "ImpresionPreCuenta"
                .Fields.Append "FechaIngreso", adVarChar, 100, adFldIsNullable
                .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
                .Fields.Append "Interconsulta", adVarChar, 100, adFldIsNullable
                .Fields.Append "ColaTipoS", adVarChar, 100, adFldIsNullable
         Case "ResumenCCosto"
                .Fields.Append "idCentroCosto", adInteger
                .Fields.Append "CentroCosto", adVarChar, 200, adFldIsNullable
                .Fields.Append "ImpAnulado", adDouble
                .Fields.Append "ImpExonerado", adDouble
                .Fields.Append "ImpNormal", adDouble
                .Fields.Append "ImpCancelado", adDouble
         Case "ResumenCentroCosto"
                .Fields.Append "idCentroCosto", adInteger
                .Fields.Append "CentroCosto", adVarChar, 200, adFldIsNullable
                .Fields.Append "SubTotal", adDouble
                .Fields.Append "ImpExonerado", adDouble
                .Fields.Append "ImpAnulado", adDouble
                .Fields.Append "PagoCta", adDouble
                .Fields.Append "ImpTotal", adDouble
         Case "DetalleCentroCosto"
                .Fields.Append "idProducto", adInteger
                .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "Producto", adVarChar, 255, adFldIsNullable
                .Fields.Append "Precio", adDouble
                .Fields.Append "SubTotal", adDouble
                .Fields.Append "canSubTotal", adDouble
                .Fields.Append "ImpExonerado", adDouble
                .Fields.Append "CanExonerado", adInteger
                .Fields.Append "ImpAnulado", adDouble
                .Fields.Append "CanAnulado", adInteger
                .Fields.Append "PagoCta", adDouble
                .Fields.Append "ImpTotal", adDouble
                .Fields.Append "CanTotal", adInteger
                .Fields.Append "Consultorio", adVarChar, 255, adFldIsNullable
         End Select
         .LockType = adLockOptimistic
         .Open
    End With
End Sub



Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mrs_Tmp = Nothing
    Set mrs_Tmp1 = Nothing
End Sub

Sub ResumenXcentroCosto(mrs_Tmp As Recordset)
        Dim lnExoneracionCamaSP As Double, lnExoneracionSOPSP As Double, lnExoneracionRestoSP As Double, lnExoneracionItemSP As Double, lnExoneracionConsultaSP As Double
        Dim lcTexto3 As String, LcTexto2 As String, LcTexto1 As String
        Dim lRecordCount As Long, lnLineas As Long, lnIdTipoServicio As Long, lnIdCentroCosto As Long
        Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double, lnImptotal As Double
        Dim lnCama As Double, lnSoloSOP As Double, lnResto As Double, lnQueda As Double
        Dim lnAnuladoCama As Double, lnImptotalCama As Double, lnExoneracionConsulta As Double
        Dim lnAnuladoSOP As Double, lnImptotalSOP As Double
        Dim lnAnuladoResto As Double, lnImptotalResto As Double
        Dim lnConsulta As Double, lnAnuladoConsulta As Double
        Dim lnImptotalConsulta As Double
        Dim lnImpLaboratorio As Double, lnExoneracionLaboratorioSP As Double
        Dim lnImpImagenes As Double, lnExoneracionImagenesSP As Double
        Dim lnAnuladoLaboratorio As Double, lnImpTotalLaboratorio As Double
        Dim lnAnuladoImagenes As Double, lnImpTotalImagenes As Double
        Dim lbPrimeraVez As Boolean, lbEncontroDato As Boolean
        Dim rsTmp As New Recordset
        Dim rsreporte As New Recordset
        Dim oRsExoneradoBoleta As New Recordset
        Dim rsTmp1 As New Recordset
        Dim rsTmp2 As New Recordset
        Dim rsTmp3 As New Recordset
        Dim oRsTmpReemb1 As New Recordset
        Dim rsTmp999 As New Recordset
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long, lnIdProductoEmergencia As Long
        Dim lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long
        Dim lnImptotalDetall As Double, lnCanTotalDetall As Long
        Dim lnIdProductoDetall As Long, lnCodigoDetall As String, lnDescripcionDetall As String, lnPrecioDetall As Double
        Dim lbHallo As Boolean, lnIdEstadoComprobante As Long
        Dim lnExoneraciones100  As Double, lnTotal100 As Double, lnImporteNeto As Double
        Dim lnImporteXcuenta As Double, lnImporteXitem As Double, lnTotalGrabado As Double
        Dim lnIdCentroCosto1 As Long, lbEntroAlDetalle As Boolean, wxParametro549 As String
        
        On Error GoTo ErrRXCC
        '
        wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
        Set rsTmp3 = mo_ReglasCaja.CajaCajaSegunFiltro("")
        '
        lnIdProductoEmergencia = mo_ReglasFacturacion.ObtenerCodigoDeConsultaDeEmergencia()
        mo_ReglasComunes.ActualizaCentroCostosParaItems IIf(mb_ConOtrosSaludDesagregado = True, True, False), moConexion
        '
        lcTexto3 = "..comienza a generar Temporal..."
        GenerarRecordsetTemporal lc_TipoReporte
        lcTexto3 = "..llena Temporal..."
        If lc_TipoReporte <> "DetalleCentroCosto" Then
            'Llena Temporal con Centro de Costos
            Set rsTmp = mo_ReglasComunes.CentrosCostoSeleccionarTodos(moConexion)
            If rsTmp.RecordCount = 0 Then
               MsgBox "Llene tabla CENTRO DE COSTOS"
               Exit Sub
            End If
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!IdCentroCosto = rsTmp.Fields!IdCentroCosto
               mrs_Tmp.Fields!CentroCosto = rsTmp.Fields!descripcion
               mrs_Tmp.Update
               rsTmp.MoveNext
            Loop
            rsTmp.Close
        End If
        lcTexto3 = "..comienza a jalar boletas servicios..."
        'Boletas Servicios
        Set rsreporte = mo_ReglasFacturacion.ServicioConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        rsreporte.Filter = IIf(mb_tieneCredito = True, "TieneCredito<>null", "TieneCredito=null")
        lRecordCount = rsreporte.RecordCount
        If lRecordCount > 0 Then
           Set rsTmp999 = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumentoFechas(mda_FechaInicio, mda_FechaFin, moConexion)
           If lc_TipoReporte <> "DetalleCentroCosto" Then
                mo_ProgressRpt.Min = 0
                mo_ProgressRpt.Max = lRecordCount
                mo_ProgressRpt.Value = 0
                mo_ProgressRpt.ShowText = True
                mo_ProgressRpt.Color = vbGreen
           End If
           lnLineas = 0
           '
           rsreporte.MoveFirst
           Do While Not rsreporte.EOF
If Trim(rsreporte.Fields!nroSerie) = "003" And Trim(rsreporte.Fields!nrodocumento) = "274806" Then
LcTexto2 = ""
End If
              LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
              lnPagoCta = rsreporte.Fields!Adelantos
              lnExoneracion = rsreporte.Fields!exoneraciones
              lnIdEstadoComprobante = rsreporte.Fields!idEstadoComprobante
              If rsreporte.Fields!idEstadoComprobante = 9 Then
                 lnAnulado = rsreporte.Fields!TotalPagado
                 lnImptotal = 0
              Else
                 lnAnulado = 0
                 lnImptotal = rsreporte.Fields!TotalPagado
              End If
              lcServicioPaciente = ""
              lbEntroAlDetalle = True
              '
              lnIdTipoServicio = 0
              If rsreporte.Fields!idCuentaAtencion > 0 Then
                lcTexto3 = "..comienza a jalar tipo servicios..."
                Set rsTmp = mo_ReglasAdmision.FacturacionCuentasAtencionXnroCuentaConexion(rsreporte.Fields!idCuentaAtencion, moConexion)
                If rsTmp.RecordCount > 0 Then
                   lnIdTipoServicio = rsTmp.Fields!idTipoServicio
                End If
                rsTmp.Close
              End If
              lbPrimeraVez = True
              '
              Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsreporte.Fields!IdComprobantePago, moConexion)
              '
              Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                    lnLineas = lnLineas + 1
                    If lc_TipoReporte <> "DetalleCentroCosto" Then
                       mo_ProgressRpt.Value = lnLineas
                    End If
                    '
                    If lbPrimeraVez = True Then
                       lbPrimeraVez = False
                       lbEncontroDato = False
                       lnIdCentroCosto = 0
                       '*****Consulta Externa (menos Medicina Fisica y Rehabilitacion)
                       If lbEncontroDato = False And lnIdTipoServicio = 1 Then
                            lcTexto3 = "..comienza: Consulta Externa (menos Medicina Fisica y Rehabilitacion)..."
                            Set rsTmp = mo_ReglasFacturacion.FactOrdenServicioPoridOrdenConexion(rsreporte.Fields!IdOrden, moConexion)
                            If rsTmp.RecordCount > 0 Then
                                If rsTmp.Fields!idPuntoCarga = 6 And rsTmp.Fields!idServicioPaciente <> 68 Then
                                   lcServicioPaciente = rsTmp.Fields!dServicioPaciente
                                   lnIdCentroCosto = 1003
                                   lbEncontroDato = True
                                End If
                            End If
                            rsTmp.Close
                       End If
                       '*****Consulta Externa (solo Medicina Fisica y Rehabilitacion)
                       If lbEncontroDato = False And lnIdTipoServicio = 1 Then
                            lcTexto3 = "..comienza: Consulta Externa (solo Medicina Fisica y Rehabilitacion)..."
                            Set rsTmp = mo_ReglasFacturacion.FactOrdenServicioPoridOrdenConexion(rsreporte.Fields!IdOrden, moConexion)
                            If rsTmp.RecordCount > 0 Then
                                If rsTmp.Fields!idPuntoCarga = 6 And rsTmp.Fields!idServicioPaciente = 68 Then
                                   lcServicioPaciente = rsTmp.Fields!dServicioPaciente
                                   lnIdCentroCosto = 1004
                                   lbEncontroDato = True
                                End If
                            End If
                            rsTmp.Close
                       End If
                       '*****Hospitalizacion
                       If lbEncontroDato = False And lnIdTipoServicio = 3 Then
                          lnCama = 0: lnSoloSOP = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                          lcTexto3 = "..comienza: Hospitalizacion..."
                          'Set rsTmp = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumento(rsReporte.Fields!nroSerie, rsReporte.Fields!NroDocumento, moConexion)
                          rsTmp999.Filter = "idComprobantePago=" & rsreporte.Fields!IdComprobantePago
                          If rsTmp999.RecordCount > 0 Then
                             '
                             lnExoneracionCamaSP = 0: lnExoneracionSOPSP = 0: lnExoneracionRestoSP = 0
                             lnExoneracionLaboratorioSP = 0: lnExoneracionImagenesSP = 0
                             '
                             lbEncontroDato = True
                             rsTmp999.MoveFirst
                             Do While Not rsTmp999.EOF
                                '
                                lnExoneracionItemSP = 0
                                If oRsExoneradoBoleta.RecordCount > 0 Then
                                   oRsExoneradoBoleta.MoveFirst
                                   Do While Not oRsExoneradoBoleta.EOF
                                      If oRsExoneradoBoleta.Fields!IdOrden = rsTmp999.Fields!IdOrden And oRsExoneradoBoleta.Fields!idProducto = rsTmp999.Fields!idProducto Then
                                         lnExoneracionItemSP = lnExoneracionItemSP + oRsExoneradoBoleta.Fields!TotalFinanciado
                                      End If
                                      oRsExoneradoBoleta.MoveNext
                                   Loop
                                End If
                                '
                                lcTexto3 = "..comienza: Hospitalizacion-NO cama..."
                                Set rsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPoridOrdenConexion(rsTmp999.Fields!IdOrden, moConexion)
                                lbHallo = False
                                If rsTmp1.RecordCount > 0 Then
                                   If rsTmp1.Fields!idPuntoCarga = 9 Then
                                        lbHallo = True
                                        'Estancia
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnCama = lnCama + lnImporteNeto
                                        lnExoneracionCamaSP = lnExoneracionCamaSP + lnExoneracionItemSP
                                        '
                                        If lc_TipoReporte = "DetalleCentroCosto" Then
                                            lcServicioPaciente = rsTmp1.Fields!dServicioPaciente
                                            If lnExoneracionItemSP > 0 Then
                                               lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                            Else
                                               lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                            End If
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6      '***Devolucion
                                               lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            Case 9      '***Anulado
                                               lnImptotalDetall = 0: lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                            Case Else   '***Pagado
                                               lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            End Select
                                            lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                            lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                            GrabaDetalleEnTmp 1005, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, _
                                                              lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, _
                                                              lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, _
                                                              lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, _
                                                              rsTmp999.Fields!Total
                                        End If
                                    End If
                                End If
                                If lbHallo = False And rsTmp1.RecordCount > 0 Then
                                   If (rsTmp1.Fields!idPuntoCarga = 2 Or rsTmp1.Fields!idPuntoCarga = 3 Or rsTmp1.Fields!idPuntoCarga = 11) Then
                                        lbHallo = True
                                        'Laboratorio
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnImpLaboratorio = lnImpLaboratorio + lnImporteNeto
                                        lnExoneracionLaboratorioSP = lnExoneracionLaboratorioSP + lnExoneracionItemSP
                                        '
                                        If lc_TipoReporte = "DetalleCentroCosto" Then
                                            If lnExoneracionItemSP > 0 Then
                                               lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                            Else
                                               lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                            End If
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6      '***Devolucion
                                               lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            Case 9      '***Anulado
                                               lnImptotalDetall = 0: lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                            Case Else   '***Pagado
                                               lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            End Select
                                            lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                            lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                            GrabaDetalleEnTmp 1001, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                        End If
                                   End If
                                End If
                                If lbHallo = False And rsTmp1.RecordCount > 0 Then
                                   If (rsTmp1.Fields!idPuntoCarga >= 20 And rsTmp1.Fields!idPuntoCarga <= 23) Then
                                        lbHallo = True
                                        'Imagenes
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnImpImagenes = lnImpImagenes + lnImporteNeto
                                        lnExoneracionImagenesSP = lnExoneracionImagenesSP + lnExoneracionItemSP
                                        '
                                        If lc_TipoReporte = "DetalleCentroCosto" Then
                                            If lnExoneracionItemSP > 0 Then
                                               lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                            Else
                                               lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                            End If
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6      '***Devolucion
                                               lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            Case 9      '***Anulado
                                               lnImptotalDetall = 0: lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                            Case Else   '***Pagado
                                               lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            End Select
                                            lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                            lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                            GrabaDetalleEnTmp 1002, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                        End If
                                   End If
                                End If
                                If lbHallo = False And rsTmp1.RecordCount > 0 Then
                                    If rsTmp1.Fields!idServicioPaciente = 73 Then
                                        lbHallo = True
                                        'Sala de Operaciones (Centro Quirurgico)
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnSoloSOP = lnSoloSOP + lnImporteNeto
                                        lnExoneracionSOPSP = lnExoneracionSOPSP + lnExoneracionItemSP
                                        '
                                        If lc_TipoReporte = "DetalleCentroCosto" Then
                                            lcServicioPaciente = rsTmp1.Fields!dServicioPaciente
                                            If lnExoneracionItemSP > 0 Then
                                               lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                            Else
                                               lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                            End If
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6      '***Devolucion
                                               lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            Case 9      '***Anulado
                                               lnImptotalDetall = 0: lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                            Case Else   '***Pagado
                                               lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                            End Select
                                            lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                            lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                            GrabaDetalleEnTmp 1006, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                            
                                        End If
                                    End If
                                End If
                                If lbHallo = False Then
                                    lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                    lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                    lnResto = lnResto + lnImporteNeto
                                    lnExoneracionRestoSP = lnExoneracionRestoSP + lnExoneracionItemSP
                                    '
                                    If lc_TipoReporte = "DetalleCentroCosto" Then
                                        lcServicioPaciente = IIf(rsTmp1.RecordCount > 0, rsTmp1.Fields!dServicioPaciente, "")
                                        If lnExoneracionItemSP > 0 Then
                                           lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                        Else
                                           lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                        End If
                                        Select Case rsreporte.Fields!idEstadoComprobante
                                        Case 6      '***Devolucion
                                           lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                        Case 9      '***Anulado
                                           lnImptotalDetall = 0: lnCanTotalDetall = 0
                                           lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                        Case Else   '***Pagado
                                           lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                        End Select
                                        lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                        lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                        GrabaDetalleEnTmp 1008, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                    End If
                                End If
                                rsTmp1.Close
                                rsTmp999.MoveNext
                             Loop
                          End If
                          'rsTmp.Close
                          'Comprueba cuadre de exoneraciones NO PRORATEADAS
                          lnExoneraciones100 = lnExoneracionCamaSP + lnExoneracionSOPSP + lnExoneracionRestoSP + lnExoneracionLaboratorioSP + lnExoneracionImagenesSP
                          If lnExoneraciones100 > lnExoneracion Then
                             lnExoneracionCamaSP = lnExoneracionCamaSP - ((lnExoneraciones100) - lnExoneracion)
                          ElseIf lnExoneraciones100 < lnExoneracion Then
                             lnExoneracionCamaSP = lnExoneracionCamaSP + (lnExoneracion - lnExoneraciones100)
                          End If
                          'comprueba que cuadre totales y subTotales
                          lnTotal100 = lnCama + lnSoloSOP + lnResto + lnImpLaboratorio + lnImpImagenes
                          If lnTotal100 > rsreporte.Fields!TotalPagado Then
                             lnCama = lnCama - (lnTotal100 - rsreporte.Fields!TotalPagado)
                          ElseIf lnTotal100 < rsreporte.Fields!TotalPagado Then
                             lnCama = lnCama + (rsreporte.Fields!TotalPagado - lnTotal100)
                          End If
                          '
                          Select Case rsreporte.Fields!idEstadoComprobante
                          Case 6     '*****Devolucion
                                lnAnuladoCama = 0
                                lnImptotalCama = -lnCama
                                '
                                lnAnuladoSOP = 0
                                lnImptotalSOP = -lnSoloSOP
                                '
                                lnAnuladoResto = 0
                                lnImptotalResto = -lnResto
                                '
                                lnAnuladoLaboratorio = 0
                                lnImpTotalLaboratorio = -lnImpLaboratorio
                                '
                                lnAnuladoImagenes = 0
                                lnImpTotalImagenes = -lnImpImagenes
                          Case 9            '*******Anulado
                                lnAnuladoCama = lnCama
                                lnImptotalCama = 0
                                '
                                lnAnuladoSOP = lnSoloSOP
                                lnImptotalSOP = 0
                                '
                                lnAnuladoResto = lnResto
                                lnImptotalResto = 0
                                '
                                lnAnuladoLaboratorio = lnImpLaboratorio
                                lnImpTotalLaboratorio = 0
                                '
                                lnAnuladoImagenes = lnImpImagenes
                                lnImpTotalImagenes = 0
                          Case Else        '*****Pagado
                                lnAnuladoCama = 0
                                lnImptotalCama = lnCama
                                '
                                lnAnuladoSOP = 0
                                lnImptotalSOP = lnSoloSOP
                                '
                                lnAnuladoResto = 0
                                lnImptotalResto = lnResto
                                '
                                lnAnuladoLaboratorio = 0
                                lnImpTotalLaboratorio = lnImpLaboratorio
                                '
                                lnAnuladoImagenes = 0
                                lnImpTotalImagenes = lnImpImagenes
                          End Select
                          If lc_TipoReporte <> "DetalleCentroCosto" Then
                                'Hospitalizacion-solo cama
                                If lnCama <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1005"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoCama
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionCamaSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalCama
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalCama
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionCamaSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1005"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionCamaSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Hospitalizacion-solo SOP
                                If lnSoloSOP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1006"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoSOP
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionSOPSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalSOP
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalSOP
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionSOPSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1006"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionSOPSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Hospitalizacion-resto (sin SOP, sin CAMA)
                                If lnResto <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1008"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoResto
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionRestoSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalResto
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalResto
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionRestoSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1008"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionRestoSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Laboratorio
                                If lnImpLaboratorio <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1001"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoLaboratorio
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionLaboratorioSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImpTotalLaboratorio
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImpTotalLaboratorio
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionLaboratorioSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1001"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionLaboratorioSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Imagenes
                                If lnImpImagenes <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1002"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoImagenes
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionImagenesSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImpTotalImagenes
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImpTotalImagenes
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionImagenesSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1002"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionImagenesSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                          End If
                       End If
                       '*****Emergencia
                       lcTexto3 = "..comienza: Emergencia..."
                       If lbEncontroDato = False And (lnIdTipoServicio = 2 Or lnIdTipoServicio = 4) Then
                          lnConsulta = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                          'Set rsTmp = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumento(rsReporte.Fields!nroSerie, rsReporte.Fields!NroDocumento, moConexion)
                          rsTmp999.Filter = "idComprobantePago=" & rsreporte.Fields!IdComprobantePago
                          
                          If rsTmp999.RecordCount > 0 Then
                             '
                             lnExoneracionConsultaSP = 0: lnExoneracionRestoSP = 0
                             lnExoneracionLaboratorioSP = 0: lnExoneracionImagenesSP = 0
                             '
                             lbEncontroDato = True
                             rsTmp999.MoveFirst
                             Do While Not rsTmp999.EOF
                                '
                                lnExoneracionItemSP = 0
                                If oRsExoneradoBoleta.RecordCount > 0 Then
                                   oRsExoneradoBoleta.MoveFirst
                                   Do While Not oRsExoneradoBoleta.EOF
                                      If oRsExoneradoBoleta.Fields!IdOrden = rsTmp999.Fields!IdOrden And oRsExoneradoBoleta.Fields!idProducto = rsTmp999.Fields!idProducto Then
                                         lnExoneracionItemSP = lnExoneracionItemSP + oRsExoneradoBoleta.Fields!TotalFinanciado
                                      End If
                                      oRsExoneradoBoleta.MoveNext
                                   Loop
                                End If
                                '
                                Set rsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPoridOrdenConexion(rsTmp999.Fields!IdOrden, moConexion)
                                lbHallo = False
                                If rsTmp999.Fields!idProducto = lnIdProductoEmergencia Then
                                    lbHallo = True
                                    'consulta de emergencia (tabla parametros)
                                    lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                    lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                    lnConsulta = lnConsulta + lnImporteNeto
                                    lnExoneracionConsultaSP = lnExoneracionConsultaSP + lnExoneracionItemSP
                                    '
                                    If lc_TipoReporte = "DetalleCentroCosto" Then
                                        lcServicioPaciente = IIf(rsTmp1.RecordCount > 0, rsTmp1.Fields!dServicioPaciente, "")
                                        If lnExoneracionItemSP > 0 Then
                                           lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                        Else
                                           lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                        End If
                                        Select Case rsreporte.Fields!idEstadoComprobante
                                        Case 6      '***Devolucion
                                           lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                        Case 9      '***Anulado
                                           lnImptotalDetall = 0: lnCanTotalDetall = 0
                                           lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                        Case Else   '***Pagado
                                           lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
'                                           If lnPagoCta > 0 Then
'                                               lnImptotalDetall = lnImptotalDetall - lnPagoCta
'                                               lnPagoCta = 0
'                                           End If
                                        End Select
                                        lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                        lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                        GrabaDetalleEnTmp 1013, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                        
                                    End If
                                End If
                                
                                If lbHallo = False And rsTmp1.RecordCount > 0 Then
                                   If (rsTmp1.Fields!idPuntoCarga = 2 Or rsTmp1.Fields!idPuntoCarga = 3 Or rsTmp1.Fields!idPuntoCarga = 11) Then
                                        lbHallo = True
                                        'Laboratorio
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnImpLaboratorio = lnImpLaboratorio + lnImporteNeto
                                        lnExoneracionLaboratorioSP = lnExoneracionLaboratorioSP + lnExoneracionItemSP
                                         '
                                         If lc_TipoReporte = "DetalleCentroCosto" Then
                                             If lnExoneracionItemSP > 0 Then
                                                lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                             Else
                                                lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                             End If
                                             Select Case rsreporte.Fields!idEstadoComprobante
                                             Case 6      '***Devolucion
                                                lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                                lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                             Case 9      '***Anulado
                                                lnImptotalDetall = 0: lnCanTotalDetall = 0
                                                lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                             Case Else   '***Pagado
                                                lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                                lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                             End Select
                                             lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                             lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                             GrabaDetalleEnTmp 1001, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                         End If
                                    End If
                                End If
                                If lbHallo = False And rsTmp1.RecordCount > 0 Then
                                   If (rsTmp1.Fields!idPuntoCarga >= 20 And rsTmp1.Fields!idPuntoCarga <= 23) Then
                                        lbHallo = True
                                        'Imagenes
                                        lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                        lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                        lnImpImagenes = lnImpImagenes + lnImporteNeto
                                        lnExoneracionImagenesSP = lnExoneracionImagenesSP + lnExoneracionItemSP
                                         '
                                         If lc_TipoReporte = "DetalleCentroCosto" Then
                                             If lnExoneracionItemSP > 0 Then
                                                lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                             Else
                                                lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                             End If
                                             Select Case rsreporte.Fields!idEstadoComprobante
                                             Case 6      '***Devolucion
                                                lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                                lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                             Case 9      '***Anulado
                                                lnImptotalDetall = 0: lnCanTotalDetall = 0
                                                lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                             Case Else   '***Pagado
                                                lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                                lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                             End Select
                                             lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                             lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                             GrabaDetalleEnTmp 1002, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                         End If
                                    End If
                                End If
                                If lbHallo = False Then
                                    lnImporteNeto = rsTmp999.Fields!Importe - lnExoneracionItemSP
                                    lnImporteNeto = ProrrateaAdelantos(lnImporteNeto, rsreporte.Fields!Adelantos, rsreporte.Fields!TotalPagado)
                                    lnResto = lnResto + lnImporteNeto
                                    lnExoneracionRestoSP = lnExoneracionRestoSP + lnExoneracionItemSP
                                    '
                                    If lc_TipoReporte = "DetalleCentroCosto" Then
                                        lcServicioPaciente = IIf(rsTmp1.RecordCount > 0, rsTmp1.Fields!dServicioPaciente, "")
                                        If lnExoneracionItemSP > 0 Then
                                           lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                        Else
                                           lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                        End If
                                        Select Case rsreporte.Fields!idEstadoComprobante
                                        Case 6      '***Devolucion
                                           lnImptotalDetall = -(lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                        Case 9      '***Anulado
                                           lnImptotalDetall = 0: lnCanTotalDetall = 0
                                           lnAnuladoImpDetall = (lnImporteNeto): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                        Case Else   '***Pagado
                                           lnImptotalDetall = (lnImporteNeto): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                        End Select
                                        lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                        lnDescripcionDetall = rsTmp999.Fields!NombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                        GrabaDetalleEnTmp 1010, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total
                                    End If
                                End If
                                rsTmp1.Close
                                rsTmp999.MoveNext
                             Loop
                          End If
                          'rsTmp.Close
                          'Comprueba cuadre de exoneraciones NO PRORATEADAS
                          lnExoneraciones100 = lnExoneracionConsultaSP + lnExoneracionRestoSP + lnExoneracionLaboratorioSP + lnExoneracionImagenesSP
                          If lnExoneraciones100 > lnExoneracion Then
                               lnExoneracionRestoSP = lnExoneracionRestoSP - ((lnExoneraciones100) - lnExoneracion)
                          ElseIf lnExoneraciones100 < lnExoneracion Then
                               lnExoneracionRestoSP = lnExoneracionRestoSP + (lnExoneracion - lnExoneraciones100)
                          End If
                          '
                          lnTotal100 = lnConsulta + lnResto + lnImpLaboratorio + lnImpImagenes
                          If lnTotal100 > rsreporte.Fields!TotalPagado Then
                               lnResto = lnResto - ((lnTotal100) - rsreporte.Fields!TotalPagado)
                          ElseIf lnTotal100 < rsreporte.Fields!TotalPagado Then
                               lnResto = lnResto + (rsreporte.Fields!TotalPagado - lnTotal100)
                          End If
                          Select Case rsreporte.Fields!idEstadoComprobante
                          Case 6      '***Devoluciones
                                lnAnuladoConsulta = 0
                                lnImptotalConsulta = -lnConsulta
                                '
                                lnAnuladoResto = 0
                                lnImptotalResto = -lnResto
                                '
                                lnAnuladoLaboratorio = 0
                                lnImpTotalLaboratorio = -lnImpLaboratorio
                                '
                                lnAnuladoImagenes = 0
                                lnImpTotalImagenes = -lnImpImagenes
                          Case 9    '***anulado
                                lnAnuladoConsulta = lnConsulta
                                lnImptotalConsulta = 0
                                '
                                lnAnuladoResto = lnResto
                                lnImptotalResto = 0
                                '
                                lnAnuladoLaboratorio = lnImpLaboratorio
                                lnImpTotalLaboratorio = 0
                                '
                                lnAnuladoImagenes = lnImpImagenes
                                lnImpTotalImagenes = 0
                          Case Else  '***Pagado
                                lnAnuladoConsulta = 0
                                lnImptotalConsulta = lnConsulta
                                '
                                lnAnuladoResto = 0
                                lnImptotalResto = lnResto
                                '
                                lnAnuladoLaboratorio = 0
                                lnImpTotalLaboratorio = lnImpLaboratorio
                                '
                                lnAnuladoImagenes = 0
                                lnImpTotalImagenes = lnImpImagenes
                          End Select
                          If lc_TipoReporte <> "DetalleCentroCosto" Then
                                'Emergencia-solo CONSULTA
                                If lnConsulta > 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1013"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsultaSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionConsultaSP > 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1013"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsultaSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Emergencia-resto (sin CONSULTA)
                                If lnResto > 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1010"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoResto
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionRestoSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalResto
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalResto
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionRestoSP > 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1010"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionRestoSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Laboratorio
                                If lnImpLaboratorio <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1001"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoLaboratorio
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionLaboratorioSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImpTotalLaboratorio
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImpTotalLaboratorio
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionLaboratorioSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1001"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionLaboratorioSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                                'Imagenes
                                If lnImpImagenes <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1002"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoImagenes
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionImagenesSP
                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImpTotalImagenes
                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImpTotalImagenes
                                        mrs_Tmp.Update
                                      End If
                                ElseIf lnExoneracionImagenesSP <> 0 Then
                                      mrs_Tmp.MoveFirst
                                      mrs_Tmp.Find "idCentroCosto=1002"
                                      If Not mrs_Tmp.EOF Then
                                        mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionImagenesSP
                                        mrs_Tmp.Update
                                      End If
                                End If
                           End If
                       End If
                       '*****Reembolsos: Farmacia y/o Servicios
                       If lbEncontroDato = False Then
                          If lnAnulado = 0 Then
                                lcTexto3 = "..comienza: Reembolsos: Farmacia y/o Servicios..."
                                Set rsTmp = mo_ReglasFacturacion.ReembolsoDetalleSeleccionaPorIdComprobantePago(rsreporte.Fields!IdComprobantePago, moConexion)
                                If rsTmp.RecordCount > 0 Then
                                    lnResto = 0   'total reembolso farmacia
                                    Do While Not rsTmp.EOF
                                       lnResto = lnResto + rsTmp.Fields!ReembolsoPagadoFarmacia
                                       rsTmp.MoveNext
                                    Loop
                                    If lc_TipoReporte <> "DetalleCentroCosto" Then
                                        '*****Farmacia
                                        If lnResto > 0 Then
                                            lnExoneracionConsulta = 0
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6    '**devolucion
                                                    lnImptotalConsulta = -lnResto
                                                    lnAnuladoConsulta = 0
                                            Case 9    '***anulado
                                                    lnImptotalConsulta = 0
                                                    lnAnuladoConsulta = lnResto
                                            Case Else '**pagado
                                                    lnImptotalConsulta = lnResto
                                                    lnAnuladoConsulta = 0
                                            End Select
                                            mrs_Tmp.MoveFirst
                                            mrs_Tmp.Find "idCentroCosto=1009"
                                            If Not mrs_Tmp.EOF Then
                                              mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                              mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                              mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                              mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                              mrs_Tmp.Update
                                              lbEncontroDato = True
                                              lnIdCentroCosto = 0
                                            End If
                                        End If
                                        '*****Servicio por REEMBOLSO
                                        lnResto = rsreporte.Fields!TotalPagado - lnResto
                                        If lnResto > 0 Then
                                            '*****18/7/11
                                            lnTotalGrabado = 0
                                            rsTmp.MoveFirst
                                            Do While Not rsTmp.EOF
                                               If rsTmp.Fields!ReembolsoPagadoServicio > 0 Then
                                                       Set oRsTmpReemb1 = mo_ReglasFacturacion.ReembolsoDetalleItemSeleccionarPorIdCuenta(rsTmp.Fields!idCuentaAtencion, moConexion)
                                                       If oRsTmpReemb1.RecordCount > 0 Then
                                                             oRsTmpReemb1.MoveFirst
                                                             lnImporteXcuenta = 0
                                                             Do While Not oRsTmpReemb1.EOF
                                                                  lnImporteXcuenta = lnImporteXcuenta + oRsTmpReemb1.Fields!TotalFinanciado
                                                                  oRsTmpReemb1.MoveNext
                                                             Loop
                                                             '
                                                             oRsTmpReemb1.MoveFirst
                                                             Do While Not oRsTmpReemb1.EOF
                                                                    If oRsTmpReemb1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or oRsTmpReemb1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or oRsTmpReemb1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Then
                                                                       'Laboratorio
                                                                       lnIdCentroCosto = 1001
                                                                    ElseIf oRsTmpReemb1.Fields!idPuntoCarga >= sghPuntosCargaBasicos.sghPtoCargaEcogGeneral And oRsTmpReemb1.Fields!idPuntoCarga <= sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica Then
                                                                       'Imagenes
                                                                       lnIdCentroCosto = 1002
                                                                    ElseIf oRsTmpReemb1.Fields!idTipoServicio = sghTipoServicio.sghConsultaExterna And oRsTmpReemb1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionCE And oRsTmpReemb1.Fields!idServicioPaciente <> 68 Then
                                                                       'Consulta Externa
                                                                       lnIdCentroCosto = 1003
                                                                    ElseIf oRsTmpReemb1.Fields!idTipoServicio = sghTipoServicio.sghHospitalizacion Then
                                                                       'Hospitalizacion
                                                                        If oRsTmpReemb1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion Then
                                                                           'Hospitalizacion-CAMA
                                                                           lnIdCentroCosto = 1005
                                                                        ElseIf oRsTmpReemb1.Fields!idServicioPaciente = 73 Then
                                                                           'Hospitalizacion-SALA OPERACIONES
                                                                           lnIdCentroCosto = 1006
                                                                        Else
                                                                           'Procedimientos Hospitalarios
                                                                           lnIdCentroCosto = 1008
                                                                        End If
                                                                    ElseIf oRsTmpReemb1.Fields!idTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Or oRsTmpReemb1.Fields!idTipoServicio = sghTipoServicio.sghEmergenciaObservacion Then
                                                                        'Emergencia
                                                                         If oRsTmpReemb1.Fields!idProducto = lnIdProductoEmergencia Then
                                                                            'Emergencia-CONSULTA
                                                                            lnIdCentroCosto = 1013
                                                                         Else
                                                                            'Emergencia-OTROS PROCEDIMIENTOS
                                                                            lnIdCentroCosto = 1010
                                                                         End If
                                                                    ElseIf oRsTmpReemb1.Fields!IdServicioGrupo = 5 Then
                                                                        'Procedimientos Administrativos
                                                                        lnIdCentroCosto = 1011
                                                                    Else
                                                                        'No se hallo en Hosp/Emerg/Cons.Ext =>Otros SALUD
                                                                        lnIdCentroCosto = 999
                                                                    End If
                                                                    '
                                                                    lnImporteXitem = ProrratearCuentaReembolsada(rsTmp.Fields!ReembolsoPagadoServicio, lnImporteXcuenta, oRsTmpReemb1.Fields!TotalFinanciado)
                                                                    lnTotalGrabado = lnTotalGrabado + lnImporteXitem
                                                                    lnExoneracionConsulta = 0
                                                                    Select Case rsreporte.Fields!idEstadoComprobante
                                                                     Case 6    '**devolucion
                                                                             lnImptotalConsulta = -lnImporteXitem
                                                                             lnAnuladoConsulta = 0
                                                                     Case 9    '***anulado
                                                                             lnImptotalConsulta = 0
                                                                             lnAnuladoConsulta = lnImporteXitem
                                                                     Case Else '***pagado
                                                                             lnImptotalConsulta = lnImporteXitem
                                                                             lnAnuladoConsulta = 0
                                                                    End Select
                                                                    '
                                                                    mrs_Tmp.MoveFirst
                                                                    mrs_Tmp.Find "idCentroCosto=" & lnIdCentroCosto
                                                                    If Not mrs_Tmp.EOF Then
                                                                       mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                                                       mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                                                       mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                                                       mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                                                       mrs_Tmp.Update
                                                                       lbEncontroDato = True
                                                                       lnIdCentroCosto = 0
                                                                    End If
                                                                    '
                                                                    oRsTmpReemb1.MoveNext
                                                             Loop
                                                             oRsTmpReemb1.Close
                                                      Else
                                                          'No se hallo en Hosp/Emerg/Cons.Ext =>Otros SALUD
                                                          lnIdCentroCosto = 999
                                                          '
                                                          lnImporteXitem = rsTmp.Fields!ReembolsoPagadoServicio
                                                          '
                                                          lnTotalGrabado = lnTotalGrabado + lnImporteXitem
                                                          lnExoneracionConsulta = 0
                                                          Select Case rsreporte.Fields!idEstadoComprobante
                                                          Case 6    '**devolucion
                                                                  lnImptotalConsulta = -lnImporteXitem
                                                                  lnAnuladoConsulta = 0
                                                          Case 9    '***anulado
                                                                  lnImptotalConsulta = 0
                                                                  lnAnuladoConsulta = lnImporteXitem
                                                          Case Else '***pagado
                                                                  lnImptotalConsulta = lnImporteXitem
                                                                  lnAnuladoConsulta = 0
                                                          End Select
                                                          '
                                                          mrs_Tmp.MoveFirst
                                                          mrs_Tmp.Find "idCentroCosto=" & lnIdCentroCosto
                                                          If Not mrs_Tmp.EOF Then
                                                             mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                                             mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                                             mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                                             mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                                             mrs_Tmp.Update
                                                             lbEncontroDato = True
                                                             lnIdCentroCosto = 0
                                                          End If
                                                     End If
                                               End If
                                               rsTmp.MoveNext
                                            Loop
                                            If lnTotalGrabado <> lnResto Then
                                                'Se descuadro por DECIMAS
                                                lnImporteXitem = lnResto - lnTotalGrabado
                                                lnExoneracionConsulta = 0
                                                Select Case rsreporte.Fields!idEstadoComprobante
                                                  Case 6    '**devolucion
                                                          lnImptotalConsulta = -lnImporteXitem
                                                          lnAnuladoConsulta = 0
                                                  Case 9    '***anulado
                                                          lnImptotalConsulta = 0
                                                          lnAnuladoConsulta = lnImporteXitem
                                                  Case Else '***pagado
                                                          lnImptotalConsulta = lnImporteXitem
                                                          lnAnuladoConsulta = 0
                                                End Select
                                                '
                                                mrs_Tmp.MoveFirst
                                                mrs_Tmp.Find "idCentroCosto=999"  'Otros SALUD
                                                If Not mrs_Tmp.EOF Then
                                                    mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                                    mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                                    mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                                    mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                                    mrs_Tmp.Update
                                                    lbEncontroDato = True
                                                    lnIdCentroCosto = 0
                                                End If
                                               
                                            End If
                                            '*****18/7/11
                                        End If
                                     End If
                                     '
                                     If lc_TipoReporte = "DetalleCentroCosto" Then
                                        'Reembolso Farmacia
                                        lnExoneracionImpDetall = 0: lnExoneracionCanDetall = 0
                                        Select Case rsreporte.Fields!idEstadoComprobante
                                        Case 6      '***Devolucion
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                           lnImptotalDetall = -lnResto: lnCanTotalDetall = 1
                                        Case 9      '***Anulado
                                           lnAnuladoImpDetall = lnResto: lnAnuladoCanDetall = 1
                                           lnImptotalDetall = 0: lnCanTotalDetall = 0
                                        Case Else   '***Pagado
                                           lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                           lnImptotalDetall = lnResto: lnCanTotalDetall = 1
                                        End Select
                                        lnIdProductoDetall = lcParametro252
                                        Set rsTmp2 = mo_ReglasComunes.CatalogoServiciosSeleccionarXidentificador(lnIdProductoDetall, moConexion)
                                        If rsTmp2.RecordCount > 0 Then
                                           lnCodigoDetall = rsTmp2.Fields!Codigo: lnDescripcionDetall = rsTmp2.Fields!NombreMINSA
                                        Else
                                           lnCodigoDetall = "": lnDescripcionDetall = ""
                                        End If
                                        rsTmp2.Close
                                        lnPrecioDetall = 1
                                        GrabaDetalleEnTmp 1009, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, 1, lnResto
                                        'Reembolso Servicio
                                        lnResto = rsreporte.Fields!TotalPagado - lnResto
                                        If lnResto <> 0 Then
                                            lnExoneracionImpDetall = 0: lnExoneracionCanDetall = 0
                                            Select Case rsreporte.Fields!idEstadoComprobante
                                            Case 6      '***Devolucion
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                               lnImptotalDetall = -lnResto: lnCanTotalDetall = 1
                                            Case 9      '***Anulado
                                               lnAnuladoImpDetall = lnResto: lnAnuladoCanDetall = 1
                                               lnImptotalDetall = 0: lnCanTotalDetall = 0
                                            Case Else   '***Pagado
                                               lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                               lnImptotalDetall = lnResto: lnCanTotalDetall = 1
                                            End Select
                                            lnIdProductoDetall = lcParametro251
                                            Set rsTmp2 = mo_ReglasComunes.CatalogoServiciosSeleccionarXidentificador(lnIdProductoDetall, moConexion)
                                            If rsTmp2.RecordCount > 0 Then
                                               lnCodigoDetall = rsTmp2.Fields!Codigo: lnDescripcionDetall = rsTmp2.Fields!NombreMINSA
                                            Else
                                               lnCodigoDetall = "": lnDescripcionDetall = ""
                                            End If
                                            rsTmp2.Close
                                            lnPrecioDetall = 1
                                            GrabaDetalleEnTmp 1015, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, 1, lnResto
                                        End If
                                     End If
                                End If
                                rsTmp.Close
                            End If
                       End If
                       '*****Laboratorio
                       If lbEncontroDato = False Then
                            lcTexto3 = "..comienza: Laboratorio..."
                            Set rsTmp = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & rsreporte.Fields!idProducto, moConexion)
                            If rsTmp.RecordCount > 0 Then
                                If mb_ConOtrosSaludDesagregado = True Then
                                    rsTmp.MoveFirst
                                    Do While Not rsTmp.EOF
                                        If rsTmp.Fields!idPuntoCarga = 2 Or rsTmp.Fields!idPuntoCarga = 3 Then
                                           lnIdCentroCosto = 1001
                                           lbEncontroDato = True
                                           Exit Do
                                        End If
                                        rsTmp.MoveNext
                                    Loop
                                Else
                                    If rsTmp.Fields!idPuntoCarga = 2 Or rsTmp.Fields!idPuntoCarga = 3 Then
                                       lnIdCentroCosto = 1001
                                       lbEncontroDato = True
                                    End If
                                End If
                            End If
                            rsTmp.Close
                       End If
                       '*****Imagenes
                       If lbEncontroDato = False Then
                            lcTexto3 = "..comienza: Imagenes..."
                            Set rsTmp = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & rsreporte.Fields!idProducto, moConexion)
                            If rsTmp.RecordCount > 0 Then
                                If mb_ConOtrosSaludDesagregado = True Then
                                   rsTmp.MoveFirst
                                   Do While Not rsTmp.EOF
                                        If rsTmp.Fields!idPuntoCarga >= 20 And rsTmp.Fields!idPuntoCarga <= 23 Then
                                           lnIdCentroCosto = 1002
                                           lbEncontroDato = True
                                           Exit Do
                                        End If
                                        rsTmp.MoveNext
                                   Loop
                                Else
                                    If rsTmp.Fields!idPuntoCarga >= 20 And rsTmp.Fields!idPuntoCarga <= 23 Then
                                       lnIdCentroCosto = 1002
                                       lbEncontroDato = True
                                    End If
                                End If
                            End If
                            rsTmp.Close
                       End If
                       '*****Procedimientos Administrativos
                       If lbEncontroDato = False Then
                            If rsreporte.Fields!IdServicioGrupo = 5 Then
                               lnIdCentroCosto = 1011
                               If rsreporte!idProducto = Val(wxParametro549) Then
                                    rsTmp3.MoveFirst
                                    rsTmp3.Find "idCaja=" & rsreporte!IdCaja
                                    If Not rsTmp3.EOF Then
                                       If Not IsNull(rsTmp3!IdCentroCosto) Then
                                          lnIdCentroCosto = rsTmp3!IdCentroCosto
                                       End If
                                    End If
                                  
                               End If
                               lbEncontroDato = True
                            End If
                       End If
                       '*****Otros SALUD
                       If lbEncontroDato = False Then
                          lnIdCentroCosto = 999
                       End If
                    End If
                    If (lnIdCentroCosto = 999 Or lnIdCentroCosto = 1011) Then
                       'solo DETALLE para: *****Procedimientos Administrativos
                       '                   *****Otros SALUD
                       If mb_DetallaProcAdmyOtrosServ = True And lbEntroAlDetalle = True Then
                            lbEntroAlDetalle = False
                            lnIdCentroCosto1 = lnIdCentroCosto
                            Set rsTmp = mo_ReglasFacturacion.FacturacionServicioPagosXidComprobantePagoConexion(rsreporte.Fields!IdComprobantePago, moConexion)
                            If rsTmp.RecordCount > 0 Then
                               rsTmp.MoveFirst
                               Do While Not rsTmp.EOF
                                  If rsTmp.Fields!IdServicioSubGrupo = 2 Then
                                     'Imagenes
                                     lnIdCentroCosto = 1002
                                  ElseIf rsTmp.Fields!IdServicioSubGrupo = 3 Then
                                     'laboratorio
                                     lnIdCentroCosto = 1001
                                  ElseIf Val(rsTmp.Fields!Codigo) = 99281 Then
                                     lnIdCentroCosto = 1013            'Emergencia-CONSULTA
                                  ElseIf Val(rsTmp.Fields!Codigo) = 99221 Then
                                     lnIdCentroCosto = 1005            'Hospitalizacion cama
                                  ElseIf Val(rsTmp.Fields!Codigo) = 99201 Then
                                     lnIdCentroCosto = 1003            'Consulta Externa
                                  Else
                                     lnIdCentroCosto = lnIdCentroCosto1
                                     If lnIdCentroCosto1 = 999 Then
                                        Set rsTmp2 = mo_ReglasComunes.FactCatalogoServiciosPtosSeleccionar(" where idProducto=" & rsTmp.Fields!idProducto, moConexion)
                                        If rsTmp2.RecordCount > 0 Then
                                           Select Case rsTmp2.Fields!idPuntoCarga
                                           Case sghPuntosCargaBasicos.sghPtoCargaAdmisionCE
                                                lnIdCentroCosto = 1003            'Consulta Externa
                                           Case sghPuntosCargaBasicos.sghPtoCargaAdmisionEmergencia
                                                If oRsTmpReemb1.Fields!idProducto = lnIdProductoEmergencia Then
                                                   lnIdCentroCosto = 1013            'Emergencia-CONSULTA
                                                Else
                                                   lnIdCentroCosto = 1010            'Emergencia-OTROS PROCEDIMIENTOS
                                                End If
                                           Case sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion
                                                lnIdCentroCosto = 1008            'Procedimientos Hospitalarios
                                            
                                           End Select
                                        End If
                                        rsTmp2.Close
                                     End If
                                     
                                  End If
                                  lnImporteXitem = rsTmp.Fields!Total
                                  '
                                  lnExoneracionConsulta = 0: lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                  If lnExoneracion > 0 Then
                                     If rsreporte.Fields!idEstadoComprobante = 9 Then
                                        lnExoneracionConsulta = ProrratearCuentaReembolsada(lnExoneracion, lnAnulado + lnExoneracion, rsTmp.Fields!Total)
                                     Else
                                        lnExoneracionConsulta = ProrratearCuentaReembolsada(lnExoneracion, lnImptotal + lnExoneracion, rsTmp.Fields!Total)
                                     End If
                                     lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionConsulta
                                     lnImporteXitem = lnImporteXitem - lnExoneracionImpDetall
                                  End If
                                  '
                                  Select Case rsreporte.Fields!idEstadoComprobante
                                  Case 6    '**devolucion
                                          lnImptotalConsulta = -lnImporteXitem: lnAnuladoConsulta = 0
                                          '
                                          lnImptotalDetall = -(lnImporteXitem): lnCanTotalDetall = rsTmp.Fields!Cantidad
                                          lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                  Case 9    '***anulado
                                          lnImptotalConsulta = 0: lnAnuladoConsulta = lnImporteXitem
                                          '
                                          lnImptotalDetall = 0: lnCanTotalDetall = 0
                                          lnAnuladoImpDetall = (lnImporteXitem): lnAnuladoCanDetall = rsTmp.Fields!Cantidad
                                  Case Else '***pagado
                                          lnImptotalConsulta = lnImporteXitem: lnAnuladoConsulta = 0
                                          '
                                          lnImptotalDetall = (lnImporteXitem): lnCanTotalDetall = rsTmp.Fields!Cantidad
                                          lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                  End Select
                                  '
                                  If lc_TipoReporte = "DetalleCentroCosto" Then
                                     GrabaDetalleEnTmp lnIdCentroCosto, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, _
                                                       lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, _
                                                       rsTmp.Fields!idProducto, rsTmp.Fields!Codigo, _
                                                       rsTmp.Fields!nombre, rsTmp.Fields!precio, _
                                                       rsTmp.Fields!Cantidad, rsTmp.Fields!Total
                                  Else
                                     mrs_Tmp.MoveFirst
                                     mrs_Tmp.Find "idCentroCosto=" & lnIdCentroCosto
                                     If Not mrs_Tmp.EOF Then
                                       mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoConsulta
                                       mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionConsulta
                                       mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalConsulta
                                       mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalConsulta
                                       mrs_Tmp.Update
                                     End If
                                  End If
                                  rsTmp.MoveNext
                               Loop
                            End If
                            rsTmp.Close
                            lnIdCentroCosto = 0
                       End If
                       
                    End If
                    '**** ***** ***** *****
                    If lc_TipoReporte = "DetalleCentroCosto" Then
                        If lnIdCentroCosto > 0 And ml_IdCentroCostos = lnIdCentroCosto Then
                            '
                            lnExoneracionImpDetall = 0
                            lnExoneracionCanDetall = 0
                            If oRsExoneradoBoleta.RecordCount > 0 Then
                               oRsExoneradoBoleta.MoveFirst
                               Do While Not oRsExoneradoBoleta.EOF
                                  If oRsExoneradoBoleta.Fields!IdOrden = rsreporte.Fields!IdOrden And oRsExoneradoBoleta.Fields!idProducto = rsreporte.Fields!idProducto Then
                                     lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!TotalFinanciado
                                     lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
                                  End If
                                  oRsExoneradoBoleta.MoveNext
                               Loop
                            End If
                            '
                            lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                            lnImptotalDetall = 0: lnCanTotalDetall = 0
                            Select Case rsreporte.Fields!idEstadoComprobante
                            Case 6      '***Devolucion
                               lnImptotalDetall = -(rsreporte.Fields!Total - lnExoneracionImpDetall)
                               lnCanTotalDetall = rsreporte.Fields!Cantidad
                            Case 9      '***Anulado
                               lnAnuladoImpDetall = rsreporte.Fields!Total - lnExoneracionImpDetall
                               lnAnuladoCanDetall = rsreporte.Fields!Cantidad
                            Case Else   '***Pagado
                               lnImptotalDetall = rsreporte.Fields!Total - lnExoneracionImpDetall
                               lnCanTotalDetall = rsreporte.Fields!Cantidad
                               If lnPagoCta > 0 Then
                                   lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                   lnPagoCta = 0
                               End If
                            End Select
                            '
                            GrabaDetalleEnTmp lnIdCentroCosto, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, _
                                              lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, _
                                              rsreporte.Fields!idProducto, rsreporte.Fields!Codigo, rsreporte.Fields!NombreMINSA, _
                                              rsreporte.Fields!precio, rsreporte.Fields!Cantidad, rsreporte.Fields!Total
                            
                        End If
                    End If
                    '**** ***** ***** *****
                    rsreporte.MoveNext
                    If rsreporte.EOF Then
                       Exit Do
                    End If
              Loop
              '
              oRsExoneradoBoleta.Close
              '
              If lc_TipoReporte <> "DetalleCentroCosto" Then
                    If lnIdEstadoComprobante = 6 Then
                       '***Devolucion
                       lnImptotal = -lnImptotal
                    End If
                    If lnIdCentroCosto > 0 Then
                          mrs_Tmp.MoveFirst
                          mrs_Tmp.Find "idCentroCosto=" & lnIdCentroCosto
                          If Not mrs_Tmp.EOF Then
                            mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
                            mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
                            mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotal
                            mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotal
                            mrs_Tmp.Update
                          End If
                    End If
                    lcTexto3 = "..comienza: termino loop..."
              End If
           Loop
        End If
        'MEDICAMENTOS emitidos en CAJA SERVICIO
        lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
        Set rsreporte = Nothing
        Set rsreporte = mo_ReglasFacturacion.FarmaciaConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        rsreporte.Filter = IIf(mb_tieneCredito = True, "TieneCredito<>null", "TieneCredito=null")
        lRecordCount = rsreporte.RecordCount
        If lRecordCount > 0 Then
           Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(moConexion)
           If lc_TipoReporte <> "DetalleCentroCosto" Then
                mo_ProgressRpt1.Min = 0
                mo_ProgressRpt1.Max = lRecordCount
                mo_ProgressRpt1.Value = 0
                mo_ProgressRpt1.ShowText = True
                mo_ProgressRpt1.Color = vbGreen
           End If
           lnLineas = 0
           '
           rsreporte.MoveFirst
           lnExoneracion = 0: lnAnulado = 0: lnImptotal = 0
           Do While Not rsreporte.EOF
              LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
              lnExoneracion = lnExoneracion + rsreporte.Fields!exoneraciones
              lnPagoCta = rsreporte.Fields!Adelantos
              If rsreporte.Fields!idEstadoComprobante = 9 Then
                 lnAnulado = lnAnulado + rsreporte.Fields!TotalPagado
              Else
                 lnImptotal = lnImptotal + rsreporte.Fields!TotalPagado
              End If
              'Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionBienesFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
              oRsExoneradoBoleta.Filter = "idComprobantePago=" & rsreporte.Fields!IdComprobantePago
              '
              Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                    lnLineas = lnLineas + 1
                    If lc_TipoReporte <> "DetalleCentroCosto" Then
                       mo_ProgressRpt1.Value = lnLineas
                    End If
                    If lc_TipoReporte = "DetalleCentroCosto" Then
                        If ml_IdCentroCostos = 1009 Then
                            '
                            lnExoneracionImpDetall = 0
                            lnExoneracionCanDetall = 0
                            If oRsExoneradoBoleta.RecordCount > 0 Then
                               oRsExoneradoBoleta.MoveFirst
                               Do While Not oRsExoneradoBoleta.EOF
                                  If oRsExoneradoBoleta.Fields!MovTipo = rsreporte.Fields!MovTipo And oRsExoneradoBoleta.Fields!movNumero = rsreporte.Fields!movNumero And oRsExoneradoBoleta.Fields!idProducto = rsreporte.Fields!idProducto Then
                                     lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!TotalFinanciado
                                     lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
                                  End If
                                  oRsExoneradoBoleta.MoveNext
                               Loop
                            End If
                            '
                            lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                            lnImptotalDetall = 0: lnCanTotalDetall = 0
                            If rsreporte.Fields!idEstadoComprobante = 9 Then
                               lnAnuladoImpDetall = rsreporte.Fields!TotalPagar - lnExoneracionImpDetall
                               lnAnuladoCanDetall = rsreporte.Fields!CantidadPagar
                            Else
                               lnImptotalDetall = rsreporte.Fields!TotalPagar - lnExoneracionImpDetall
                               lnCanTotalDetall = rsreporte.Fields!CantidadPagar
                               If lnPagoCta > 0 Then
                                   lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                   lnPagoCta = 0
                               End If
                            End If
                            '
                            GrabaDetalleEnTmp 1009, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, rsreporte.Fields!idProducto, rsreporte.Fields!Codigo, rsreporte.Fields!Producto, rsreporte.Fields!PrecioVenta, rsreporte.Fields!CantidadPagar, rsreporte.Fields!TotalPagar
                            
                        End If
                    End If
                    '
                    rsreporte.MoveNext
                    If rsreporte.EOF Then
                       Exit Do
                    End If
              Loop
           Loop
           If lc_TipoReporte <> "DetalleCentroCosto" Then
                mrs_Tmp.MoveFirst
                mrs_Tmp.Find "idCentroCosto=1009"  'Farmacia
                mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
                mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotal
                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotal
                mrs_Tmp.Update
           End If
        End If
        Set rsTmp = Nothing
        Set rsreporte = Nothing
        Set oRsExoneradoBoleta = Nothing
        Set rsTmp1 = Nothing
        Set rsTmp2 = Nothing
        Set oRsTmpReemb1 = Nothing
        Set rsTmp999 = Nothing
        Set mo_ReglasCaja = Nothing
        Set rsTmp3 = Nothing
        Exit Sub
ErrRXCC:
        MsgBox Err.Description & Chr(13) & "Boleta: " & LcTexto2
        Exit Sub
Resume
End Sub


Sub GrabaDetalleEnTmp(lnIdCentroCosto As Long, mrs_Tmp As Recordset, lnExoneracionImpDetall As Double, _
                      lnExoneracionCanDetall As Long, lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long, _
                      lnImptotalDetall As Double, lnCanTotalDetall As Long, lnIdProducto1 As Long, lcCodigo1 As String, _
                      lcDescripcion1 As String, lnPrecio1 As Double, lnCantidad1 As Long, lnTotal1 As Double)
        Dim lbPrimeraVez As Boolean
        If ml_IdCentroCostos = lnIdCentroCosto Then
            lbPrimeraVez = True
            If mrs_Tmp.RecordCount > 0 Then
               mrs_Tmp.MoveFirst
               Do While Not mrs_Tmp.EOF
                  If mb_TotalizarXconsultorio = True Then
                        If mrs_Tmp.Fields!idProducto = lnIdProducto1 And mrs_Tmp.Fields!Consultorio = lcServicioPaciente Then
                           lbPrimeraVez = False
                           Exit Do
                        End If
                  Else
                        If mrs_Tmp.Fields!idProducto = lnIdProducto1 Then
                           lbPrimeraVez = False
                           Exit Do
                        End If
                  End If
                  mrs_Tmp.MoveNext
               Loop
            End If
            If lbPrimeraVez = True Then
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!idProducto = lnIdProducto1
                  mrs_Tmp.Fields!Producto = lcDescripcion1
                  mrs_Tmp.Fields!Codigo = lcCodigo1
                  mrs_Tmp.Fields!precio = lnPrecio1
            End If
            mrs_Tmp.Fields!Subtotal = mrs_Tmp.Fields!Subtotal + lnTotal1
            mrs_Tmp.Fields!canSubtotal = mrs_Tmp.Fields!canSubtotal + lnCantidad1
            mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracionImpDetall
            mrs_Tmp.Fields!CanExonerado = mrs_Tmp.Fields!CanExonerado + lnExoneracionCanDetall
            mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnuladoImpDetall
            mrs_Tmp.Fields!CanAnulado = mrs_Tmp.Fields!CanAnulado + lnAnuladoCanDetall
            mrs_Tmp.Fields!PagoCta = mrs_Tmp.Fields!PagoCta + 0
            mrs_Tmp.Fields!ImpTotal = mrs_Tmp.Fields!ImpTotal + lnImptotalDetall
            mrs_Tmp.Fields!CanTotal = mrs_Tmp.Fields!CanTotal + lnCanTotalDetall
            mrs_Tmp.Fields!Consultorio = lcServicioPaciente
            mrs_Tmp.Update
        End If
End Sub


Function ProrrateaAdelantos(lnImporteNetoItem As Double, lnAdelantosBoleta As Double, lnTotalPagadoBoleta As Double) As Double
   If lnAdelantosBoleta > 0 And lnImporteNetoItem > 0 Then
      ProrrateaAdelantos = Round((lnImporteNetoItem * lnTotalPagadoBoleta) / (lnTotalPagadoBoleta + lnAdelantosBoleta), 2)
   Else
      ProrrateaAdelantos = lnImporteNetoItem
   End If
End Function


Function ProrratearCuentaReembolsada(lnImporteReembolsado As Double, lnImporteCuentaTotal As Double, lnImporteCuentaItem As Double) As Double
    ProrratearCuentaReembolsada = Round((lnImporteReembolsado * lnImporteCuentaItem) / lnImporteCuentaTotal, 2)
End Function

