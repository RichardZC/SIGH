VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rCrystal 
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "rCrystal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CrvReportes 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
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
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ln_DestinoReporte As sghImpresion
'aqui declara los objetos que contendra al rporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes    'debb-27/05/2015
Dim lc_TextoDelFiltro As String
Dim lc_TipoReporte As String
Dim lnIdPuntoCarga As Long
Dim lnOrdenadoPor As Long
Dim mrs_Tmp As New Recordset
Dim mrs_Tmp1 As New Recordset
Dim mrs_Tmp2 As New Recordset
Dim rsTmpSOAT As New Recordset
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_HoraInicio As String
Dim ml_HoraFin As String
Dim mb_ConsiderarSinMovimientos As Boolean
Dim mb_SeMuestraLotes As Boolean
Dim mb_StockMinimoMayorAcantidad As Boolean
Dim ml_idUsuario As Long
Dim ml_idProducto  As Long
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
Dim ml_IdPlan As Long
Dim lc_TipoServicioHosp As String
Dim mda_FechaSolicitudDesde As Date
Dim mda_FechaSolicitudHasta  As Date
Dim ml_lcTipoServicio As String
Dim ml_IncluyeHistoriasQueSalieron As Boolean
Dim ms_UltimosDigitosHC As String
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String
Property Let UltimosDigitosHC(sValue As String)
    ms_UltimosDigitosHC = sValue
End Property


Property Let IncluyeHistoriasQueSalieron(lValue As Boolean)
    ml_IncluyeHistoriasQueSalieron = lValue
End Property
Property Let lcTipoServicio(lValue As String)
    ml_lcTipoServicio = lValue
End Property
Property Let FechaSolicitudDesde(lValue As Date)
    mda_FechaSolicitudDesde = lValue
End Property
Property Let FechaSolicitudHasta(lValue As Date)
    mda_FechaSolicitudHasta = lValue
End Property

Property Let TipoServicioHosp(lValue As String)
    lc_TipoServicioHosp = lValue
End Property
Property Let IdPlan(lValue As Long)
    ml_IdPlan = lValue
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

Property Let Estado(lValue As Long)
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
    ml_idProducto = lValue
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

Property Let IdPuntoCarga(iValue As Long)
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
Property Let IdResponsable(iValue As Long)
   lnIdResponsable = iValue
End Property


Property Set RecordSet_mrs_Tmp(oValue As ADODB.Recordset)
    Set mrs_Tmp = oValue
'    Set mrs_Tmp = oValue.Clone
End Property
Property Let DestinoReporte(lValue As sghImpresion)
    ln_DestinoReporte = lValue
End Property


Private Sub Form_Activate()
    If Len(lc_TextoDelFiltro) > 250 Then
       lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
    End If


    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim rsReporte As New ADODB.Recordset
    Dim rsTmp As New Recordset
    Dim rsTmp10 As New Recordset
    
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
    Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
    Dim ldFechaPrincipio As Date
    Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim oConexion As New ADODB.Connection
    Dim lbPrimeraVez As Boolean
    Dim lcTexto1 As String, lcTexto2 As String, lcTexto3 As String, lcTexto4 As String
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim lnDebioPagar As Double, lnTotalPorCuenta As Double
    Dim lnIdTipoServicio As Long, lnIdDepartamento As Long, lnIdEspacialidad As Long, lnIdServicioPaciente As Long
    Dim lcDtipoServicio As String, lcDpto As String, lcDespecialidad As String, lcDServicio As String
    Dim lnSalieron As Integer, lnRetornaron As Integer, lnSalieronYnoRetornaron As Integer
    Dim lnTSalieron As Integer, lnTRetornaron As Integer, lnTSalieronYnoRetornaron As Integer
    Dim lcSql As String, lbEsCuentaPagante As Boolean
    Dim lbProcesaEnServidor As Boolean, lcHOraInicio As String, lcHoraFinal As String, lbContinuar As Boolean
    Dim lnCuenta As Long, lnIdFormaPago As Long
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
    lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    
    mflgContinuar = False
    Select Case lc_TipoReporte
    
    Case "EExoneracionXboleta"    'debb-21/03/2015 (inicio)
        mflgContinuar = True
        Set crReport = crApp.OpenReport(App.Path & "\plantillas\EExoneracionXboleta.rpt", 1)
        ' Parametros del reporte
        Set crParamDefs = crReport.ParameterFields
        For Each crParamDef In crParamDefs
            Select Case crParamDef.ParameterFieldName
            Case "subTitulo"
                crParamDef.AddCurrentValue (lc_TextoDelFiltro)
            Case "lcEESS"
                crParamDef.AddCurrentValue (lcTitEESS)
            Case "lcEESSdireccion"
                crParamDef.AddCurrentValue (lcTitDireccion)
            Case "lcEESStelefono"
                crParamDef.AddCurrentValue (lcTitTelefono)
            End Select
        Next
        crReport.Database.SetDataSource mrs_Tmp
        'debb-21/03/2015 (fin)
    Case "HcXmedico"
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_idUsuario, mda_FechaInicio, mda_FechaFin, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, (Val(ml_lcTipoServicio)), ml_IncluyeHistoriasQueSalieron)
        lc_TextoDelFiltro = " "
        If rsReporte.RecordCount > 0 Then
            GenerarRecordsetTemporal "ImpresionPreCuenta"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lbContinuar = True
               If Val(ms_UltimosDigitosHC) > 0 Then
                   If Right(Trim(Str(rsReporte!NroHistoriaClinica)), Len(ms_UltimosDigitosHC)) <> ms_UltimosDigitosHC Then
                      lbContinuar = False
                   End If
               End If
               If lbContinuar = True And rsReporte.Fields!HoraRequerida >= ml_HoraInicio And rsReporte.Fields!HoraRequerida <= ml_HoraFin Then
                    lnCuenta = 0: lnIdFormaPago = 0
                    lcTexto3 = "": lcTexto4 = ""
                    Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesPorIdAtencionConTurnoYfuentefinanciamiento(rsReporte!idAtencion, oConexion)
                    If mrs_Tmp2.RecordCount > 0 Then
                       lnCuenta = mrs_Tmp2!idCuentaAtencion
                       lcTexto3 = IIf(IsNull(mrs_Tmp2!dfuente), "", mrs_Tmp2!dfuente)
                       lcTexto4 = IIf(IsNull(mrs_Tmp2!dturno), "", mrs_Tmp2!dturno)
                       lnIdFormaPago = mrs_Tmp2!IdFormaPago
                    End If
                    mrs_Tmp2.Close
                    '
                    lcTexto1 = "": lcTexto2 = ""
                    If lnIdFormaPago = 1 Then
                        Set mrs_Tmp2 = mo_ReglasCaja.CajaComprobantesPagoXcuenta(lnCuenta, oConexion)
                        If mrs_Tmp2.RecordCount > 0 Then
                           lcTexto1 = mrs_Tmp2!nroSerie & " - " & mrs_Tmp2!NroDocumento
                           lcTexto2 = mrs_Tmp2!fechaCobranza
                        End If
                        mrs_Tmp2.Close
                    End If
                    '
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!Servicio = rsReporte.Fields!Servicio & "     (" & Trim(Str(rsReporte.Fields!idMedico)) & ")"
                    mrs_Tmp.Fields!Medico = rsReporte.Fields!nMedico
                    mrs_Tmp.Fields!FechaIngreso = Format(rsReporte.Fields!FechaRequerida, "dd/mm/yyyy") & " - " & rsReporte.Fields!HoraRequerida
                    mrs_Tmp.Fields!NroHistoriaClinica = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte.Fields!NroHistoriaClinica)), False) & _
                                                        IIf(Format(rsReporte!fechaCreacion, sighEntidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, " (n)", "")
                    mrs_Tmp.Fields!Interconsulta = rsReporte.Fields!FichaFamiliar
                    mrs_Tmp.Fields!Paciente = rsReporte.Fields!nPaciente
                    mrs_Tmp.Fields!Usuario = lcTexto3    'fuente financiamiento
                    mrs_Tmp.Fields!ColaTipoS = lcTexto1    'N° boleta
                    mrs_Tmp.Fields!fboleta = lcTexto2    'fecha boleta
                    If Not IsNull(rsReporte.Fields!NroDocumento) Then
                       mrs_Tmp.Fields!Cola = rsReporte.Fields!NroDocumento 'dni
                    End If
                    mrs_Tmp.Fields!turno = lcTexto4       'mañana/tarde...
                    
                    mrs_Tmp.Update
               End If
               rsReporte.MoveNext
            Loop
            mrs_Tmp.Sort = "servicio,fechaIngreso"
            lc_TextoDelFiltro = "F. requerida: " & Format(mda_FechaInicio, sighEntidades.DevuelveFechaSoloFormato_DMY_HM) & "  al " & Format(mda_FechaFin, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
            'Reporte
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\HCsolicitadasXmedicoFF.rpt", 1)
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
        oConexion.Close
    Case "HcXmedicoXpagina"
        Set rsReporte = mo_AdminReportes.ReporteHistoriasSolicitadasCEPorMedico(ml_idUsuario, mda_FechaInicio, mda_FechaFin, mda_FechaSolicitudDesde, mda_FechaSolicitudHasta, (Val(ml_lcTipoServicio)), ml_IncluyeHistoriasQueSalieron)
        lc_TextoDelFiltro = " "
        If rsReporte.RecordCount > 0 Then
            GenerarRecordsetTemporal "ImpresionPreCuenta"
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lbContinuar = True
               If Val(ms_UltimosDigitosHC) > 0 Then
                   If Right(Trim(Str(rsReporte!NroHistoriaClinica)), Len(ms_UltimosDigitosHC)) <> ms_UltimosDigitosHC Then
                      lbContinuar = False
                   End If
               End If
               If lbContinuar = True And rsReporte.Fields!HoraRequerida >= ml_HoraInicio And rsReporte.Fields!HoraRequerida <= ml_HoraFin Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!Servicio = rsReporte.Fields!Servicio & "     (" & Trim(Str(rsReporte.Fields!idMedico)) & ")"
                    mrs_Tmp.Fields!Medico = rsReporte.Fields!nMedico
                    mrs_Tmp.Fields!FechaIngreso = Format(rsReporte.Fields!FechaRequerida, "dd/mm/yyyy") & " - " & rsReporte.Fields!HoraRequerida
                    mrs_Tmp.Fields!NroHistoriaClinica = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsReporte.Fields!NroHistoriaClinica)), False) & _
                                                        IIf(Format(rsReporte!fechaCreacion, sighEntidades.DevuelveFechaSoloFormato_DMY) = lcBuscaParametro.RetornaFechaServidorSQL, " (n)", "")
                    mrs_Tmp.Fields!Interconsulta = rsReporte.Fields!FichaFamiliar
                    mrs_Tmp.Fields!Paciente = rsReporte.Fields!nPaciente
                    mrs_Tmp.Fields!Usuario = rsReporte.Fields!NroDocumento
                    mrs_Tmp.Update
               End If
               rsReporte.MoveNext
            Loop
            mrs_Tmp.Sort = "servicio,fechaIngreso"
            'Reporte
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\HCsolicitadasXmedico.rpt", 1)
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
    Case "EExoGeneral"
        'Proceso
        Set mrs_Tmp1 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorFechas(mda_FechaInicio, mda_FechaFin)
        mrs_Tmp1.Filter = "exoneraciones>0"
        If mrs_Tmp1.RecordCount = 0 Then
             mflgContinuar = False
        Else
             GenerarRecordsetTemporal lc_TipoReporte
             mrs_Tmp1.MoveFirst
             Do While Not mrs_Tmp1.EOF
                If mrs_Tmp1.Fields!IdTipoOrden = 1 Then
                   Set rsTmp = mo_ReglasFacturacion.FactOrdenServicioSeleccionarPorIdComprobantePago(mrs_Tmp1.Fields!IdComprobantePago)
                Else
                   Set rsTmp = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorIdComprobantePago(mrs_Tmp1.Fields!IdComprobantePago)
                End If
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    Do While Not rsTmp.EOF
                        lnDebioPagar = 0
                        If mrs_Tmp1.Fields!IdTipoOrden = 1 Then
                           Set mrs_Tmp2 = mo_ReglasFacturacion.FacturacionServicioPagosSeleccionarPorId(rsTmp.Fields!idOrdenPago)
                           If mrs_Tmp2.RecordCount > 0 Then
                              mrs_Tmp2.MoveFirst
                              Do While Not mrs_Tmp2.EOF
                                 lnDebioPagar = lnDebioPagar + mrs_Tmp2.Fields!Total
                                 mrs_Tmp2.MoveNext
                              Loop
                           End If
                        Else
                           Set mrs_Tmp2 = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorId(rsTmp.Fields!idOrden)
                           If mrs_Tmp2.RecordCount > 0 Then
                              mrs_Tmp2.MoveFirst
                              Do While Not mrs_Tmp2.EOF
                                 lnDebioPagar = lnDebioPagar + mrs_Tmp2.Fields!TotalPagar
                                 mrs_Tmp2.MoveNext
                              Loop
                           End If
                        End If
                        If IsNull(rsTmp.Fields!idTipoServicio) Then
                            'boleta exonerada de un  Paciente Externo (sin HC)
                            lnIdTipoServicio = 0
                            lcDtipoServicio = "Externo"
                            lnIdDepartamento = 0
                            lcDpto = "Externo"
                            lnIdEspacialidad = 0
                            lcDespecialidad = "Externo"
                            lnIdServicioPaciente = 0
                            lcDServicio = "Externo"
                        Else
                            lnIdTipoServicio = rsTmp.Fields!idTipoServicio
                            lcDtipoServicio = rsTmp.Fields!DTipoServicio
                            lnIdDepartamento = rsTmp.Fields!IdDepartamento
                            lcDpto = rsTmp.Fields!Ddpto
                            lnIdEspacialidad = rsTmp.Fields!IdEspecialidad
                            lcDespecialidad = rsTmp.Fields!DEspecialidad
                            lnIdServicioPaciente = rsTmp.Fields!IdServicioPaciente
                            lcDServicio = rsTmp.Fields!DServicio
                        End If
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           mrs_Tmp.Find "idServicio=" & lnIdServicioPaciente
                           If Not mrs_Tmp.EOF Then
                              lbPrimeraVez = False
                           End If
                        End If
                        If lbPrimeraVez = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!idTipoServicio = lnIdTipoServicio
                            mrs_Tmp.Fields!TipoServicio = lcDtipoServicio
                            mrs_Tmp.Fields!IdDepartamento = lnIdDepartamento
                            mrs_Tmp.Fields!Departamento = lcDpto
                            mrs_Tmp.Fields!IdEspecialidad = lnIdEspacialidad
                            mrs_Tmp.Fields!especialidad = lcDespecialidad
                            mrs_Tmp.Fields!idServicio = lnIdServicioPaciente
                            mrs_Tmp.Fields!Servicio = lcDServicio
                            mrs_Tmp.Fields!Cantidad = 1
                            mrs_Tmp.Fields!CostoReal = lnDebioPagar
                            mrs_Tmp.Fields!CostoExonerado = rsTmp.Fields!ImporteExonerado
                            mrs_Tmp.Fields!Pago = lnDebioPagar - rsTmp.Fields!ImporteExonerado
                        Else
                            mrs_Tmp.Fields!Cantidad = mrs_Tmp.Fields!Cantidad + 1
                            mrs_Tmp.Fields!CostoReal = mrs_Tmp.Fields!CostoReal + lnDebioPagar
                            mrs_Tmp.Fields!CostoExonerado = mrs_Tmp.Fields!CostoExonerado + rsTmp.Fields!ImporteExonerado
                            mrs_Tmp.Fields!Pago = mrs_Tmp.Fields!Pago + (lnDebioPagar - rsTmp.Fields!ImporteExonerado)
                        End If
                        mrs_Tmp.Update
                        rsTmp.MoveNext
                    Loop
                End If
                mrs_Tmp1.MoveNext
             Loop
             'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\EconExoGeneral.rpt", 1)
             ' Parametros del reporte
             Set crParamDefs = crReport.ParameterFields
             For Each crParamDef In crParamDefs
                 Select Case crParamDef.ParameterFieldName
                     Case "subTitulo"
                         crParamDef.AddCurrentValue (lc_TextoDelFiltro)
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
    Case "EConsumoXPtoCarga"
        '
        lnIdProducto = mrs_Tmp.RecordCount
        'Reporte
        mflgContinuar = True
        Set crReport = crApp.OpenReport(App.Path & "\plantillas\EconConsumoXpto.rpt", 1)
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
    Case "AHCMovimEntSal"
        
        GenerarRecordsetTemporal lc_TipoReporte
        Set mrs_Tmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaMovimientosPorDia(mda_FechaInicio, mda_FechaFin)
        If mrs_Tmp1.RecordCount > 0 Then
           mrs_Tmp1.MoveFirst
           Do While Not mrs_Tmp1.EOF
              If mrs_Tmp1.Fields!idMotivo = 9 Then
                 'Retorno al Archivo
                 If InStr(lc_TipoServicioHosp, Trim(Str(mrs_Tmp1.Fields!idTipoServicioOrigen))) > 0 Then
                    If mrs_Tmp.RecordCount > 0 Then
                       lbPrimeraVez = True
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "NroHistoria=" & mrs_Tmp1.Fields!NroHistoriaClinica
                       If Not mrs_Tmp.EOF Then
                          Do While Not mrs_Tmp.EOF
                             If mrs_Tmp.Fields!idServicio = mrs_Tmp1.Fields!IdServicioOrigen And InStr("ST", mrs_Tmp.Fields!es) > 0 Then
                                mrs_Tmp.Fields!HcYaSalio = 1
                                mrs_Tmp.Update
                                lbPrimeraVez = False
                                Exit Do
                             End If
                             mrs_Tmp.MoveNext
                          Loop
                       End If
                    End If
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!es = IIf(lbPrimeraVez = False, "SR", "R")
                    mrs_Tmp.Fields!entSal = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp1.Fields!NroHistoriaClinica)), False) & IIf(lbPrimeraVez = False, "-SR-", "-R-") & Format(mrs_Tmp1.Fields!FechaMovimiento, "hh:mm")
                    mrs_Tmp.Fields!idServicio = mrs_Tmp1.Fields!IdServicioOrigen
                    mrs_Tmp.Fields!DServicio = mrs_Tmp1.Fields!ServicioOrigen
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp1.Fields!FechaMovimiento
                    mrs_Tmp.Fields!HcYaSalio = 0
                    mrs_Tmp.Update
                 End If
              Else
                 '***************** GalenHos v.3.0 (inicio)*****************
                 'Salida del Archivo
                 If InStr(lc_TipoServicioHosp, Trim(Str(mrs_Tmp1.Fields!idTipoServicioDestino))) > 0 Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!es = IIf(mrs_Tmp1.Fields!idMotivo = 7, "T", "S")
                    mrs_Tmp.Fields!entSal = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp1.Fields!NroHistoriaClinica)), False) & "-" & IIf(mrs_Tmp1.Fields!idMotivo = 7, "T", "S") & "-" & Format(mrs_Tmp1.Fields!FechaMovimiento, "hh:mm")
                    mrs_Tmp.Fields!idServicio = mrs_Tmp1.Fields!idServicioDestino
                    mrs_Tmp.Fields!DServicio = mrs_Tmp1.Fields!ServicioDestino
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp1.Fields!FechaMovimiento
                    mrs_Tmp.Fields!HcYaSalio = 0
                    mrs_Tmp.Update
                 End If
                 '***************** GalenHos v.3.0 (fin)*****************
              End If
              mrs_Tmp1.MoveNext
           Loop
        End If
        lnIngresos = mrs_Tmp.RecordCount
        lnIdEspacialidad = 0
        lnTSalieron = 0: lnTRetornaron = 0: lnTSalieronYnoRetornaron = 0
        If mrs_Tmp.RecordCount > 0 Then
            mrs_Tmp.Sort = "idServicio,fmovimiento"
            'asigna totales
            With mrs_Tmp2
                  .Fields.Append "idServicio", adInteger
                  .Fields.Append "dServicio", adVarChar, 50, adFldIsNullable
                  .Fields.Append "EntSal", adVarChar, 30, adFldIsNullable
                  .Fields.Append "fMovimiento", adDate
                  .LockType = adLockOptimistic
                  .Open
            End With
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnSalidas = mrs_Tmp.Fields!idServicio
               lcTexto1 = mrs_Tmp.Fields!DServicio
               lnRetornaron = 0: lnSalieron = 0: lnSalieronYnoRetornaron = 0
If Right(lcTexto1, 5) = "a III" Then
lnIdEspacialidad = 0
End If
               Do While Not mrs_Tmp.EOF And lnSalidas = mrs_Tmp.Fields!idServicio

                  If (InStr(mrs_Tmp.Fields!es, "S") > 0 Or InStr(mrs_Tmp.Fields!es, "T") > 0) And mrs_Tmp.Fields!HcYaSalio = 0 Then
                     lnSalieron = lnSalieron + 1
                  End If
                  If mrs_Tmp.Fields!es = "SR" Then
                     lnRetornaron = lnRetornaron + 1
                  End If
                  If mrs_Tmp.Fields!es = "S" And mrs_Tmp.Fields!HcYaSalio = 0 Then
                     lnSalieronYnoRetornaron = lnSalieronYnoRetornaron + 1
                  End If
                  mrs_Tmp.MoveNext
                  If mrs_Tmp.EOF Then
                     Exit Do
                  End If
               Loop
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "--------------"
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "Falta Ret: " & Trim(Str(lnSalieronYnoRetornaron))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "Retornaron: " & Trim(Str(lnRetornaron))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "Total: " & Trim(Str(lnSalieron))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               lnTSalieron = lnTSalieron + lnSalieron
               lnTRetornaron = lnTRetornaron + lnRetornaron
               lnTSalieronYnoRetornaron = lnTSalieronYnoRetornaron + lnSalieronYnoRetornaron
            Loop
            If mrs_Tmp2.RecordCount > 0 Then
               mrs_Tmp2.MoveFirst
               Do While Not mrs_Tmp2.EOF
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = 0
                    mrs_Tmp.Fields!es = " "
                    mrs_Tmp.Fields!entSal = mrs_Tmp2.Fields!entSal
                    mrs_Tmp.Fields!idServicio = mrs_Tmp2.Fields!idServicio
                    mrs_Tmp.Fields!DServicio = mrs_Tmp2.Fields!DServicio
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp2.Fields!Fmovimiento
                    mrs_Tmp.Update
                    mrs_Tmp2.MoveNext
               Loop
            End If
            'elimina Historias que SALIERON, pero YA RETORNARON
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               If mrs_Tmp.Fields!HcYaSalio = 1 Then
                    mrs_Tmp.Delete
                    mrs_Tmp.Update
               End If
               mrs_Tmp.MoveNext
            Loop
            'asigna Filas
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnSalidas = mrs_Tmp.Fields!idServicio
               lnSalidasImg = 1
               Do While Not mrs_Tmp.EOF And lnSalidas = mrs_Tmp.Fields!idServicio
                  mrs_Tmp.Fields!nroFila = lnSalidasImg
                  mrs_Tmp.Update
                  lnSalidasImg = lnSalidasImg + 1
                  mrs_Tmp.MoveNext
                  If mrs_Tmp.EOF Then
                     Exit Do
                  End If
               Loop
            Loop
            lc_TextoDelFiltro = lc_TextoDelFiltro & "      Faltan Retornar: " & Trim(Str(lnTSalieronYnoRetornaron)) & "     (Retornaron: " & Trim(Str(lnTRetornaron)) & ")     (Total: " & Trim(Str(lnTSalieron)) & ")"
            'Reporte
            
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\AHCMovimEntSal.rpt", 1)
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                    Case "Titulo"
                        crParamDef.AddCurrentValue ("MOVIMIENTO DE ENTRADA Y SALIDAS DE HISTORIAS EN EL ARCHIVO CLINICO")
                    Case "subTitulo"
                        crParamDef.AddCurrentValue (lc_TextoDelFiltro)
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
        
    Case "AHCMovimFormatos"
        '***************** GalenHos v.3.0 (inicio)*****************
        GenerarRecordsetTemporal lc_TipoReporte
        Set mrs_Tmp1 = mo_ReglasArchivoClinico.MovimientosFormatoHCMovimientosPorDia(mda_FechaInicio, mda_FechaFin)
        If mrs_Tmp1.RecordCount > 0 Then
           mrs_Tmp1.MoveFirst
           Do While Not mrs_Tmp1.EOF
              If mrs_Tmp1.Fields!idMotivo = 9 Then
                 'Retorno al Archivo
                 If InStr(lc_TipoServicioHosp, Trim(Str(mrs_Tmp1.Fields!idTipoServicioOrigen))) > 0 Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!es = "R"
                    mrs_Tmp.Fields!entSal = Trim(Str(mrs_Tmp1.Fields!NroHistoriaClinica)) & "-R-" & Format(mrs_Tmp1.Fields!FechaMovimiento, "hh:mm")
                    mrs_Tmp.Fields!idServicio = mrs_Tmp1.Fields!IdServicioOrigen
                    mrs_Tmp.Fields!DServicio = mrs_Tmp1.Fields!ServicioOrigen
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp1.Fields!FechaMovimiento
                    mrs_Tmp.Update
                 End If
              Else
                 'Salida del Archivo
                 If InStr(lc_TipoServicioHosp, Trim(Str(mrs_Tmp1.Fields!idTipoServicioDestino))) > 0 Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!es = "S"
                    mrs_Tmp.Fields!entSal = Trim(Str(mrs_Tmp1.Fields!NroHistoriaClinica)) & "-S-" & Format(mrs_Tmp1.Fields!FechaMovimiento, "hh:mm")
                    mrs_Tmp.Fields!idServicio = mrs_Tmp1.Fields!idServicioDestino
                    mrs_Tmp.Fields!DServicio = mrs_Tmp1.Fields!ServicioDestino
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp1.Fields!FechaMovimiento
                    mrs_Tmp.Update
                 End If
              End If
              mrs_Tmp1.MoveNext
           Loop
        End If
        lnIngresos = mrs_Tmp.RecordCount
        lnIdEspacialidad = 0
        lnTSalieron = 0: lnTRetornaron = 0: lnTSalieronYnoRetornaron = 0
        If mrs_Tmp.RecordCount > 0 Then
            mrs_Tmp.Sort = "idServicio,fmovimiento"
            'asigna totales
            With mrs_Tmp2
                  .Fields.Append "idServicio", adInteger
                  .Fields.Append "dServicio", adVarChar, 50, adFldIsNullable
                  .Fields.Append "EntSal", adVarChar, 30, adFldIsNullable
                  .Fields.Append "fMovimiento", adDate
                  .LockType = adLockOptimistic
                  .Open
            End With
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnSalidas = mrs_Tmp.Fields!idServicio
               lcTexto1 = mrs_Tmp.Fields!DServicio
               lnSalidasImg = 0: lnIdDepartamento = 0
               Do While Not mrs_Tmp.EOF And lnSalidas = mrs_Tmp.Fields!idServicio
                  lnSalidasImg = lnSalidasImg + 1
                  If mrs_Tmp.Fields!es = "S" Then
                     lnIdDepartamento = lnIdDepartamento + 1
                     lnIdEspacialidad = lnIdEspacialidad + 1
                  End If
                  mrs_Tmp.MoveNext
                  If mrs_Tmp.EOF Then
                     Exit Do
                  End If
               Loop
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "--------------"
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "Total: " & Trim(Str(lnSalidasImg))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "R: " & Trim(Str(lnSalidasImg - lnIdDepartamento))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
               '
               mrs_Tmp2.AddNew
               mrs_Tmp2.Fields!idServicio = lnSalidas
               mrs_Tmp2.Fields!DServicio = lcTexto1
               mrs_Tmp2.Fields!entSal = "S: " & Trim(Str(lnIdDepartamento))
               mrs_Tmp2.Fields!Fmovimiento = Now
               mrs_Tmp2.Update
            Loop
            If mrs_Tmp2.RecordCount > 0 Then
               mrs_Tmp2.MoveFirst
               Do While Not mrs_Tmp2.EOF
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!nroHistoria = 0
                    mrs_Tmp.Fields!es = " "
                    mrs_Tmp.Fields!entSal = mrs_Tmp2.Fields!entSal
                    mrs_Tmp.Fields!idServicio = mrs_Tmp2.Fields!idServicio
                    mrs_Tmp.Fields!DServicio = mrs_Tmp2.Fields!DServicio
                    mrs_Tmp.Fields!Fmovimiento = mrs_Tmp2.Fields!Fmovimiento
                    mrs_Tmp.Update
                    mrs_Tmp2.MoveNext
               Loop
            End If
            'asigna Filas
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnSalidas = mrs_Tmp.Fields!idServicio
               lnSalidasImg = 1
               Do While Not mrs_Tmp.EOF And lnSalidas = mrs_Tmp.Fields!idServicio
                  mrs_Tmp.Fields!nroFila = lnSalidasImg
                  mrs_Tmp.Update
                  lnSalidasImg = lnSalidasImg + 1
                  mrs_Tmp.MoveNext
                  If mrs_Tmp.EOF Then
                     Exit Do
                  End If
               Loop
            Loop
            lc_TextoDelFiltro = lc_TextoDelFiltro & "      Total Historias: " & Trim(Str(lnIngresos)) & "     (R: " & Trim(Str(lnIngresos - lnIdEspacialidad)) & ")     (S: " & Trim(Str(lnIdEspacialidad)) & ")"
            'Reporte
            
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\AHCMovimEntSal.rpt", 1)
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                    Case "Titulo"
                        crParamDef.AddCurrentValue ("MOVIMIENTO DE ENTRADA Y SALIDAS DE FORMATOS DE HISTORIAS EN EL ARCHIVO CLINICO")
                    Case "subTitulo"
                        crParamDef.AddCurrentValue (lc_TextoDelFiltro)
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
        '***************** GalenHos v.3.0 (fin)*****************
        'SCCQ 31/08/2020 Cambio27 Inicio
            Case "AHCSinDevolver"
                'mda_FechaFin = mda_FechaInicio + ml_Dias
                Set rsReporte = mo_ReglasArchivoClinico.SeleccionarHCSinDevolver(72)
                If rsReporte.RecordCount > 0 Then
                    'Reporte
                    mflgContinuar = True
                    Set crReport = crApp.OpenReport(App.Path & "\plantillas\AHCSinDevolver.rpt", 1)
                    ' Parametros del reporte
                    Set crParamDefs = crReport.ParameterFields
                    For Each crParamDef In crParamDefs
                        Select Case crParamDef.ParameterFieldName
                            Case "subTitulo"
                                crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                           Case "lcEESS"
                               crParamDef.AddCurrentValue (lcTitEESS)
                           Case "lcEESSdireccion"
                               crParamDef.AddCurrentValue (lcTitDireccion)
                           Case "lcEESStelefono"
                               crParamDef.AddCurrentValue (lcTitTelefono)
                        End Select
                    Next
                    crReport.Database.SetDataSource rsReporte
                End If
        'SCCQ 31/08/2020 Cambio27 Fin
    End Select
    If mflgContinuar = True Then
       If mb_EnArchivoExcel = True Then
            If lcBuscaParametro.SeleccionaFilaParametro(284) = "S" Then
                Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
                Select Case lc_TipoReporte
                Case "HcXmedicoXpagina"
                     mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "HCsolicitadasXmedico", lc_TextoDelFiltro, "", Me.hwnd
                Case "EExoGeneral"
                     mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "EconExoGeneral", lc_TextoDelFiltro, "", Me.hwnd
                Case "EConsumoXPtoCarga"
                     mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "EconConsumoXpto", lc_TextoDelFiltro, "", Me.hwnd
                Case "AHCMovimEntSal"
                     mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "MOVIMIENTO DE ENTRADA Y SALIDAS DE HISTORIAS EN EL ARCHIVO CLINICO", lc_TextoDelFiltro, "", Me.hwnd
                Case "AHCMovimFormatos"
                     mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "MOVIMIENTO DE ENTRADA Y SALIDAS DE FORMATOS DE HISTORIAS EN EL ARCHIVO CLINICO", lc_TextoDelFiltro, "", Me.hwnd
                End Select
                Set mo_ReglasReportes = Nothing
             Else
                
                 'SCCQ 13/02/2020 Problema4 Inicio
                 'El siguiente codigo es para verificar si la ruta existe
                 Dim parametro_reporte As String
                 parametro_reporte = lcBuscaParametro.SeleccionaFilaParametro(269)
                 Dim strPath As String 'Variable que contiene la ruta de la carpeta donde se generan los reportes
                 Dim posicion As Integer
                 posicion = InStrRev(parametro_reporte, "\")
                 strPath = Mid(parametro_reporte, 1, posicion - 1) '"c:\Reportes" 'Ruta para genear reporte
                 If Dir(strPath, vbDirectory) = "" Then 'Si el directorio no existe
                    MkDir strPath 'Se crea la carpeta
                 End If
                 'SCCQ 13/02/2020 Problema4 Fin
                 'Codigo que genera el archivo excel
                 crReport.ExportOptions.DestinationType = crEDTDiskFile
                 crReport.ExportOptions.FormatType = crEFTExcel70
                 crReport.ExportOptions.DiskFileName = parametro_reporte '"c:\Reportes\excel.xls" 'SCCQ 13/02/2020 Problema4 Inicio/Fin
                 crReport.Export (False)
                 MsgBox "Se generó el archivo " + parametro_reporte  'SCCQ 13/02/2020 Problema4 Inicio/Fin
                 'fin del codigo que genera el archivo excel
                
             End If
        End If
        CrvReportes.ReportSource = crReport
        CrvReportes.ViewReport
        CrvReportes.Zoom 120
        
        '
        mo_reglasComunes.grabaTablaAuditoria (crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))   'debb-27/05/2015
        
    End If
    
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    Set oConexion = Nothing
    Set mo_ReglasImagenes = Nothing
    Set rsReporte = Nothing
    Set rsTmp = Nothing
    Set oDoImagMovimientoIngresos = Nothing
    Set oDoImagMovimiento = Nothing
    LimpiarVariablesDeMemoria
    Screen.MousePointer = vbDefault
    
    If ln_DestinoReporte <> sghPantalla Then
         Me.Visible = False
    End If
    Me.MousePointer = 1
    Exit Sub
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", _
            vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
   Resume
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
                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
                .Fields.Append "Interconsulta", adVarChar, 100, adFldIsNullable
                .Fields.Append "ColaTipoS", adVarChar, 100, adFldIsNullable
                .Fields.Append "turno", adVarChar, 100, adFldIsNullable
                .Fields.Append "Cola", adVarChar, 100, adFldIsNullable
                .Fields.Append "Fboleta", adVarChar, 100, adFldIsNullable
         Case "EExoGeneral"
                .Fields.Append "IdTipoServicio", adInteger
                .Fields.Append "TipoServicio", adVarChar, 50, adFldIsNullable
                .Fields.Append "IdDepartamento", adInteger
                .Fields.Append "Departamento", adVarChar, 50, adFldIsNullable
                .Fields.Append "IdEspecialidad", adInteger
                .Fields.Append "Especialidad", adVarChar, 50, adFldIsNullable
                .Fields.Append "IdServicio", adInteger
                .Fields.Append "Servicio", adVarChar, 50, adFldIsNullable
                .Fields.Append "Cantidad", adInteger
                .Fields.Append "CostoReal", adDouble
                .Fields.Append "CostoExonerado", adDouble
                .Fields.Append "Pago", adDouble
         Case "EConsumoXPtoCarga"
                .Fields.Append "NroHistoria", adInteger
                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
                .Fields.Append "idCuentaAtencion", adInteger
                .Fields.Append "idPuntoCarga", adInteger
                .Fields.Append "dPuntoCarga", adVarChar, 40, adFldIsNullable
                .Fields.Append "Consumo", adDouble
         Case "AHCMovimEntSal", "AHCMovimFormatos"
                .Fields.Append "NroHistoria", adInteger
                .Fields.Append "ES", adVarChar, 2, adFldIsNullable
                .Fields.Append "EntSal", adVarChar, 30, adFldIsNullable
                .Fields.Append "idServicio", adInteger
                .Fields.Append "dServicio", adVarChar, 50, adFldIsNullable
                .Fields.Append "NroFila", adInteger
                .Fields.Append "Fmovimiento", adDate
                .Fields.Append "HcYaSalio", adInteger
         Case "AHCauditoria"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "Accion", adVarChar, 1, adFldIsNullable
                .Fields.Append "Nusuario", adVarChar, 50, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
                .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
                .Fields.Append "Observacion1", adVarChar, 50, adFldIsNullable
                .Fields.Append "Observacion2", adVarChar, 50, adFldIsNullable
         Case "CEauditoria"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "Accion", adVarChar, 1, adFldIsNullable
                .Fields.Append "Empleado", adVarChar, 100, adFldIsNullable
                .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
         Case "Eauditoria"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "Accion", adVarChar, 1, adFldIsNullable
                .Fields.Append "Empleado", adVarChar, 100, adFldIsNullable
                .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
         Case "EmergAuditoria"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "Accion", adVarChar, 1, adFldIsNullable
                .Fields.Append "Empleado", adVarChar, 100, adFldIsNullable
                .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
         Case "Hauditoria"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "Accion", adVarChar, 1, adFldIsNullable
                .Fields.Append "Empleado", adVarChar, 100, adFldIsNullable
                .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
         End Select
         .LockType = adLockOptimistic
         .Open
    End With
End Sub



Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mrs_Tmp = Nothing
    Set mrs_Tmp1 = Nothing
    Set mrs_Tmp2 = Nothing
End Sub


'Sub ReporteHRACreaTmp()
'    If mb_EnResumen = True Then
'        With rsTmpSOAT
'                .Fields.Append "idCuentaAtencion", adInteger
'                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
'                .Fields.Append "NroHistoria", adInteger
'                .Fields.Append "Origen", adVarChar, 10, adFldIsNullable
'                .Fields.Append "Falta", adDate, 10, adFldIsNullable
'                .Fields.Append "Halta", adVarChar, 5, adFldIsNullable
'                .Fields.Append "dFinanciamiento", adVarChar, 100, adFldIsNullable
'                .Fields.Append "tFacturado", adDouble
'                .Fields.Append "tPagado", adDouble
'                .Fields.Append "tDeuda", adDouble
'                .Fields.Append "PFarmacia", adDouble
'                .Fields.Append "PRx", adDouble
'                .Fields.Append "PLab", adDouble
'                .Fields.Append "POtros", adDouble
'                .Fields.Append "DFarmacia", adDouble
'                .Fields.Append "DRx", adDouble
'                .Fields.Append "DLab", adDouble
'                .Fields.Append "DOtros", adDouble
'                .LockType = adLockOptimistic
'                .Open
'        End With
'    End If
'End Sub
'Sub ReporteHRAagregaCtas(mrs_Tmp As Recordset)
'    If mb_EnResumen = True Then
'        Dim lbEncontroRegistro As Boolean
'        lbEncontroRegistro = False
'        If rsTmpSOAT.RecordCount > 0 Then
'           rsTmpSOAT.MoveFirst
'           rsTmpSOAT.Find "idCuentaAtencion=" & mrs_Tmp.Fields!idCuentaAtencion
'           If Not rsTmpSOAT.EOF Then
'              lbEncontroRegistro = True
'           End If
'        End If
'        If lbEncontroRegistro = False Then
'           rsTmpSOAT.AddNew
'           rsTmpSOAT.Fields!idCuentaAtencion = mrs_Tmp.Fields!idCuentaAtencion
'           rsTmpSOAT.Fields!Paciente = mrs_Tmp.Fields!Paciente
'           rsTmpSOAT.Fields!nroHistoria = mrs_Tmp.Fields!nroHistoria
'           rsTmpSOAT.Fields!origen = mrs_Tmp.Fields!origen
'           rsTmpSOAT.Fields!Falta = mrs_Tmp.Fields!FechaEgreso
'           rsTmpSOAT.Fields!Halta = mrs_Tmp.Fields!horaEgreso
'           rsTmpSOAT.Fields!dFinanciamiento = mrs_Tmp.Fields!dFuenteFinanciamiento
'           rsTmpSOAT.Fields!origen = IIf(mrs_Tmp.Fields!idTipoServicio = 1, "CE", IIf(mrs_Tmp.Fields!idTipoServicio = 3, "Hosp", "Emerg"))
'        End If
'        Select Case mrs_Tmp.Fields!IdPuntoCarga
'        Case sghPuntosCargaBasicos.sghPtoCargaRayosX
'             rsTmpSOAT.Fields!DRx = mrs_Tmp.Fields!consumo
'        Case sghPuntosCargaBasicos.sghPtoCargaFarmacia
'             rsTmpSOAT.Fields!DFarmacia = mrs_Tmp.Fields!consumo
'        Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
'             rsTmpSOAT.Fields!DLab = mrs_Tmp.Fields!consumo
'        Case Else
'             rsTmpSOAT.Fields!DOtros = mrs_Tmp.Fields!consumo
'        End Select
'        rsTmpSOAT.Fields!tFacturado = rsTmpSOAT.Fields!tFacturado + mrs_Tmp.Fields!consumo
'        rsTmpSOAT.Update
'    End If
'End Sub
'
'Sub ReporteHRApagos(rsTmp10 As Recordset, lnIdCuentaAtencion As Long, moConexion As Connection, lbEsCuentaPagante As Boolean)
'    If mb_EnResumen = True Then
'        Dim rsTmpReemb2 As New Recordset, oRsTmpReemb1 As New Recordset
'        Dim lnFarmacia As Double, lnServicio As Double, lnRx As Double, lnLaboratorio As Double, lnOtros As Double, lnPucho As Double
'        Dim lnTFarmacia As Double, lnTRx As Double, lnTLaboratorio As Double, lnTOtros As Double
'        lnTFarmacia = 0: lnTRx = 0: lnTLaboratorio = 0: lnTOtros = 0
'        rsTmp10.MoveFirst
'        Do While Not rsTmp10.EOF
'           If rsTmp10.Fields!IdEstadoComprobante = 4 Then   'solo pagados
'                If lbEsCuentaPagante = True Then
'                        If rsTmp10.Fields!IdTipoOrden <> 1 Then
'                            '**** Cuenta de un Paciente PAGANTE - BOLETA DE FARMACIA
'                            lnTFarmacia = rsTmp10.Fields!Total
'                            lnServicio = 0
'                        Else
'                            '**** Cuenta de un Paciente PAGANTE - BOLETA DE SERVICIOS
'                            lnServicio = rsTmp10.Fields!Total
'                        End If
'                Else
'                         '**** Cuenta de un Paciente con SEGUROS - BOLETA DE REEMBOLSOS
'                         Set rsTmpReemb2 = mo_ReglasFacturacion.ReembolsoDetalleSeleccionaPorIdComprobantePago(rsTmp10.Fields!idComprobantePago, moConexion)
'                         If rsTmpReemb2.RecordCount > 0 Then
'                             'Reembolsos
'                             lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
'                             rsTmpReemb2.MoveFirst
'                             Do While Not rsTmpReemb2.EOF
'                                lnFarmacia = lnFarmacia + rsTmpReemb2.Fields!ReembolsoPagadoFarmacia
'                                lnServicio = lnServicio + rsTmpReemb2.Fields!ReembolsoPagadoServicio
'                                rsTmpReemb2.MoveNext
'                             Loop
'                        End If
'                        rsTmpReemb2.Close
'               End If
'               If lnServicio > 0 Then
'                   Set oRsTmpReemb1 = mo_ReglasFacturacion.ReembolsoDetalleItemSeleccionarPorIdCuenta(lnIdCuentaAtencion, moConexion)
'                   If oRsTmpReemb1.RecordCount > 0 Then
'                      oRsTmpReemb1.MoveFirst
'                      Do While Not oRsTmpReemb1.EOF
'                            If oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Then
'                               'Laboratorio
'                               lnLaboratorio = lnLaboratorio + oRsTmpReemb1.Fields!totalFinanciado
'                            ElseIf oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX Then
'                               'Rx
'                               lnRx = lnRx + oRsTmpReemb1.Fields!totalFinanciado
'                            Else
'                               lnOtros = lnOtros + oRsTmpReemb1.Fields!totalFinanciado
'                            End If
'                            oRsTmpReemb1.MoveNext
'                      Loop
'                   End If
'                   oRsTmpReemb1.Close
'                   If lnServicio <> (lnRx + lnLaboratorio + lnOtros) Then   'cuadre
'                      If lnServicio < (lnRx + lnLaboratorio + lnOtros) Then
'                         lnPucho = (lnRx + lnLaboratorio + lnOtros) - lnServicio
'                      Else
'                         lnPucho = lnServicio - (lnRx + lnLaboratorio + lnOtros)
'                      End If
'                      lnPucho = Round(lnPucho / 3, 2)
'                      lnRx = lnRx - lnPucho
'                      lnLaboratorio = lnLaboratorio - lnPucho
'                      lnOtros = lnOtros - lnPucho
'                   End If
'                   lnTFarmacia = lnTFarmacia + lnFarmacia
'                   lnTRx = lnTRx + lnRx
'                   lnTLaboratorio = lnTLaboratorio + lnLaboratorio
'                   lnTOtros = lnTOtros + lnOtros
'               End If
'           End If
'           rsTmp10.MoveNext
'        Loop
'        If lnRx < 0 Or lnLaboratorio < 0 Or lnOtros < 0 Then   'cantidades negativas
'           If lnRx < 0 Then
'              If (lnLaboratorio + lnRx) >= 0 Then
'                 lnLaboratorio = lnLaboratorio + lnRx
'              Else
'                 lnOtros = lnOtros + lnRx
'              End If
'              lnRx = 0
'           End If
'           If lnLaboratorio < 0 Then
'              If (lnRx + lnLaboratorio) >= 0 Then
'                 lnRx = lnRx + lnLaboratorio
'              Else
'                 lnOtros = lnOtros + lnLaboratorio
'              End If
'              lnLaboratorio = 0
'           End If
'           If lnOtros < 0 Then
'              If (lnRx + lnOtros) >= 0 Then
'                 lnRx = lnRx + lnOtros
'              Else
'                 lnLaboratorio = lnLaboratorio + lnOtros
'              End If
'              lnOtros = 0
'           End If
'        End If
'        rsTmpSOAT.MoveFirst
'        rsTmpSOAT.Find "idCuentaAtencion=" & lnIdCuentaAtencion
'        If Not rsTmpSOAT.EOF Then
'            rsTmpSOAT.Fields!PFarmacia = lnTFarmacia
'            rsTmpSOAT.Fields!PRx = lnTRx
'            rsTmpSOAT.Fields!PLab = lnTLaboratorio
'            rsTmpSOAT.Fields!POtros = lnTOtros
'            rsTmpSOAT.Fields!tPagado = lnTFarmacia + lnTRx + lnTLaboratorio + lnTOtros
'            rsTmpSOAT.Fields!tDeuda = rsTmpSOAT.Fields!tFacturado - (lnTFarmacia + lnTRx + lnTLaboratorio + lnTOtros)
'            rsTmpSOAT.Update
'        End If
'    End If
'    Set rsTmpReemb2 = Nothing: Set oRsTmpReemb1 = Nothing
'End Sub
