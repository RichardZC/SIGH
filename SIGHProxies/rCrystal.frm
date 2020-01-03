VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form rCrystal 
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rCrystal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin CRVIEWERLibCtl.CRViewer CrvReportes 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   480
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
'        Programa: Muestra vista previa de varios Reportes de Farmacia
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

'aqui declara los objetos que contendra al rporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Dim lcTexto1 As String:  Dim lcTexto2 As String: Dim lcTexto3 As String
Dim lbPrimeraVez As Boolean, lnSalidas As Long, LnExoner As Long: Dim lbContinua As Boolean
Dim LnEne As Long: Dim LnFeb As Long: Dim LnMar As Long: Dim LnAbr As Long: Dim LnMay As Long: Dim LnJun As Long
Dim LnJul As Long: Dim LnAgo As Long: Dim LnSet As Long: Dim LnOct As Long: Dim LnNov As Long: Dim LnDic As Long
Dim lc_TextoDelFiltro As String, lcTexto10 As String
Dim lnIdFuenteFinanciamiento As Long:  Dim dFinanciamiento As String
Dim lc_TipoReporte As String, lcCodigo As String, lcNombre As String
Dim lnIdAlmacen As Long
Dim lnOrdenadoPor As Long: Dim lnIdProducto As Long
Dim mrs_Tmp As New Recordset
Dim mrs_tmp1 As New Recordset
Dim mrs_Tmp2 As New Recordset
Dim mrs_Tmp3 As New Recordset
Dim rsReporte As New ADODB.Recordset
Dim rsReporteAgrupado As New Recordset
Dim rsTmp As New Recordset
Dim rsTmp1 As New Recordset
Dim rsTmp111 As New Recordset
Dim mrs_Tmp99 As New Recordset
Dim rsErrores As New Recordset
Dim rsDebug As New Recordset
Dim mo_DoFarmMovimientoVentas As New DoFarmMovimientoVentas
Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
Dim oBuscaMovimientos As New farmMovimientoDetalle
Dim oDoCatProductosHosp As New DoFinanciamientoCatalogoBien

'AGREGADO X Mariano 07112014
Dim ml_mes  As Long
Dim mb_Rreportes As String
Dim mb_SolooConsolido As Boolean

Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_HoraInicio As String
Dim ml_HoraFin As String
Dim mb_ConsiderarSinMovimientos As Boolean
Dim mb_SeMuestraLotes As Boolean
Dim mb_StockMinimoMayorAcantidad As Boolean
Dim ml_idUsuario As Long
Dim ml_idProducto  As Long
Dim lnIdAlmacenOrigen As Long
Dim lnIdAlmacenDestino As Long
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
Dim ml_idFuenteFinanciamiento As Long
Dim mb_SoloPagados As Boolean
Dim mb_ConsideraOSH  As Boolean
Dim lnIngresos As Long: Dim LnDevolucionesP As Long: Dim TotIngresos  As Long
Dim LnVentas As Long: Dim lnSis As Long: Dim lnSoat As Long
Dim LnConvenio As Long: Dim lnCreditoH As Long: Dim lnDefensaN As Long
Dim LnOsDevol As Long:: Dim LnOsVencim As Long: Dim LnOsMerma As Long
Dim LnExonerac As Long:: Dim LnIntervencionS As Long
Dim LnOtrasS As Long: Dim TotSalidas As Long, LnDevolMerma As Long: Dim LnVentaInst As Long
Dim lnPrecio As Double: Dim ldFechaVencimiento As Date
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision    'debb-03/11/2015
Dim oConexion As New ADODB.Connection
Dim ldFechaInicioMovim As Date, ldFechaHistoricoXmes As Date, lcUltDiaMes As String
Dim lnIdAlmacenRep As Long, lnFor As Integer
Dim mb_ConsiderarPAdesdeServidor As Boolean
Dim mb_ConsiderarReembolsos As Boolean
Dim mb_EsUnaDonacion As Boolean
Dim ml_idTipoSalidaBienInsumo As Long
Dim mb_VtaYestrategicoSeparado As Boolean
Dim lc_TipoServicioHosp As String
Dim lc_OdbcICI As String
Dim lc_CodigoSismed As String
Dim mb_EsDonaciones As Boolean
Dim ml_IdTipoFinanciamiento As Long, lnTotalRegistros As Long
Dim ml_Observaciones As String
Dim mb_MuestraTipoSoporteSISMED As Boolean
Dim mb_SoloBoletas As Boolean
Dim ml_Proveedor As String
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String
Dim mb_NOconsiderarTipoConceptos As Boolean
Dim lc_CodigoItem As String
Dim lb_ConsideraItemsDePaquetes As Boolean
Dim mRs_Record As New Recordset
Dim lb_NOconsiderarSALDOcero As Boolean
Dim lb_ConsiderarSaldoInicialDelHistorico As Boolean
Dim lb_SeGrabaICImensual As Boolean
Dim lb_EsUnIciHistorico As Boolean
Dim lb_SeCorrigioDato As Boolean
Property Let SeCorrigioDato(lValue As Boolean)
    lb_SeCorrigioDato = lValue
End Property

Property Let EsUnIciHistorico(lValue As Boolean)
    lb_EsUnIciHistorico = lValue
End Property

Property Let SeGrabaICImensual(lValue As Boolean)
    lb_SeGrabaICImensual = lValue
End Property
Property Let ConsiderarSaldoInicialDelHistorico(lValue As Boolean)
    lb_ConsiderarSaldoInicialDelHistorico = lValue
End Property


Property Let NOconsiderarSALDOcero(lValue As Boolean)
    lb_NOconsiderarSALDOcero = lValue
End Property


Property Set oRsRecord(mRsValue As Recordset)
    Set mRs_Record = mRsValue
End Property

Property Let ConsideraItemsDePaquetes(lValue As Boolean)
    lb_ConsideraItemsDePaquetes = lValue
End Property


Property Let CodigoItem(lValue As String)
    lc_CodigoItem = lValue
End Property
Property Let NOconsiderarTipoConceptos(lValue As String)
    mb_NOconsiderarTipoConceptos = lValue
End Property
Property Let Proveedor(lValue As String)
    ml_Proveedor = lValue
End Property

Property Let SoloBoletas(lValue As Boolean)
    mb_SoloBoletas = lValue
End Property

Property Let MuestraTipoSoporteSISMED(lValue As Boolean)
    mb_MuestraTipoSoporteSISMED = lValue
End Property


Property Let Observaciones(lValue As String)
    ml_Observaciones = lValue
End Property


Property Let IdTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property


Property Let EsDonaciones(lValue As Boolean)
    mb_EsDonaciones = lValue
End Property


Property Let CodigoSismed(lValue As String)
    lc_CodigoSismed = lValue
End Property


Property Let OdbcICI(lValue As String)
    lc_OdbcICI = lValue
End Property


Property Let TipoServicioHosp(lValue As String)
    lc_TipoServicioHosp = lValue
End Property

Property Let VtaYestrategicoSeparado(lValue As Boolean)
    mb_VtaYestrategicoSeparado = lValue
End Property
Property Let idTipoSalidaBienInsumo(lValue As Long)
    ml_idTipoSalidaBienInsumo = lValue
End Property

Property Let EsUnaDonacion(lValue As Boolean)
    mb_EsUnaDonacion = lValue
End Property

Property Let ConsiderarReembolsos(lValue As Boolean)
    mb_ConsiderarReembolsos = lValue
End Property
Property Let ConsiderarPAdesdeServidor(lValue As Boolean)
    mb_ConsiderarPAdesdeServidor = lValue
End Property
Property Let ConsideraOSH(lValue As Boolean)
    mb_ConsideraOSH = lValue
End Property

'MARIANO 07112014
Property Let Mes(lValue As Long)
    ml_mes = lValue
End Property
'MARIANO 07112014
Property Let Rreportes(lValue As String)
    mb_Rreportes = lValue
End Property
Property Let SoloConsolidado(lValue As Boolean)
    mb_SolooConsolido = lValue
End Property

Property Let SoloPagados(lValue As Boolean)
    mb_SoloPagados = lValue
End Property
Property Let idFuenteFinanciamiento(lValue As Long)
    ml_idFuenteFinanciamiento = lValue
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


Property Let IdAlmacenDestino(iValue As Long)
   lnIdAlmacenDestino = iValue
End Property
Property Let IdAlmacenOrigen(iValue As Long)
   lnIdAlmacenOrigen = iValue
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

Property Let IdAlmacen(iValue As Long)
   lnIdAlmacen = iValue
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



Private Sub CrvReportes_DblClicked(ByVal x As Long, ByVal Y As Long, EventInfo As Variant, UseDefault As Boolean)
    Dim oCRVEventInfo As New CRVEventInfo
    Dim lcMovNumero As String
    Dim oFields As CRFields
    Dim oField As CRField
    Set oCRVEventInfo = EventInfo
    
    Select Case oCRVEventInfo.index
    Case 0   'ICI
         If lb_EsUnIciHistorico = True Then
             Dim orICIxITEM As New rICIxITEM
             orICIxITEM.codigo = oCRVEventInfo.Text
             orICIxITEM.AnioMes = Format(mda_FechaInicio, "yyyy") & Format(mda_FechaInicio, "mm")
             orICIxITEM.FarmaciaICI = lc_CodigoSismed
             orICIxITEM.Show 1
             Set orICIxITEM = Nothing
             Me.Visible = False
         End If
    End Select
    
'    If oCRVEventInfo.Type = 100 And (oCRVEventInfo.index = 4 Or oCRVEventInfo.index = 2) And Len(Trim(oCRVEventInfo.Text)) = 9 Then
'       '(Reporte KARDEX).................4
'       '(Reporte MOVIMIENTOS DE E/S).....2
'       lcMovNumero = oCRVEventInfo.Text
'       Set oFields = oCRVEventInfo.GetFields
'       Set oField = oFields.Item(oCRVEventInfo.index)
'
'       Select Case oField.Value
'       Case "S"
'       Case "E"
'       End Select
'    End If
    Set oCRVEventInfo = Nothing
End Sub



Private Sub Form_Activate()
    If lb_SeCorrigioDato = True Then    'debb1212
       Me.Visible = False
    End If
    If Len(lc_TextoDelFiltro) > 250 Then
       lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
    End If
    
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long, lnDivision As Integer: Dim lnDevol As Double
    Dim Lnab As Long: Dim lnReingresos As Long: Dim LnDistribucion As Long
    Dim LnTransferencia As Long: Dim LnDevolVencido As Long
    Dim lnIdTipoFinanciamientoAtenciones As Long: Dim lnBruto As Double
    Dim lnTotal As Double: Dim lcMovNumero As String: Dim lnPagoNeto As Double
    Dim lnConsultorios As Long: Dim lnHospital    As Long: Dim lnEmergencia   As Long: Dim lnClinica   As Long: Dim lnParticular As Long
    Dim lnAdelantos As Double, lnConsumoFarmacia As Double
    Dim lnIdFuenteFinanciamiento1 As Long
    Dim lnPendientePorPagarEnFarmacia As Double, lbEncontroExoneracion As Boolean
    Dim lcAnioMes As String, lnErrorEnOdbc As Integer, lbContinuar As Boolean, lnConsumoFarmacia1 As Double
    Dim lcSerieB As String, lcDocumentoB As String, lnRedondeoB As Double, lbTienePagoAcuenta As Boolean
    Dim lnTotalBol As Double, lnTotalExo As Double, lnTotalAde As Double, lbEsNuevoDocumento As Boolean
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
    lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    Select Case lc_TipoReporte
    Case "rSaldosPorAlmacen"
            'debb-13/04/2016 (inicio)
            If lnIdAlmacen > 0 Then
                If mb_SeMuestraLotes = True Then
                   If mb_StockMinimoMayorAcantidad = True Then
                      Set rsReporte = mo_ReglasFarmacia.FarmDevuelveSaldosConSinLotesPorAlmacen(lnIdAlmacen, lnOrdenadoPor, 0, 1)
                   Else
                      Set rsReporte = mo_ReglasFarmacia.FarmDevuelveSaldosConSinLotesPorAlmacen(lnIdAlmacen, lnOrdenadoPor, 0, 0)
                   End If
                   If ml_idTipoSalidaBienInsumo > 0 Then
                      If lc_CodigoItem = "1" Then    'muestra saldos mayores a CERO
                         rsReporte.Filter = "cantidad<>0 and cantidadLote<>0 and IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo & " and IdTipoSalidaBienInsumoSaldo=" & ml_idTipoSalidaBienInsumo
                      Else
                         rsReporte.Filter = "IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo & " and IdTipoSalidaBienInsumoSaldo=" & ml_idTipoSalidaBienInsumo
                      End If
                   Else
                      If lc_CodigoItem = "1" Then        'muestra saldos mayores a CERO
                         rsReporte.Filter = "cantidad<>0 and cantidadLote<>0"
                      End If
                   End If
                Else
                   If mb_StockMinimoMayorAcantidad = True Then
                      Set rsReporte = mo_ReglasFarmacia.FarmDevuelveSaldosConSinLotesPorAlmacen(lnIdAlmacen, lnOrdenadoPor, 1, 1)
                   Else
                      Set rsReporte = mo_ReglasFarmacia.FarmDevuelveSaldosConSinLotesPorAlmacen(lnIdAlmacen, lnOrdenadoPor, 1, 0)
                   End If
                   If ml_idTipoSalidaBienInsumo > 0 Then
                      If lc_CodigoItem = "1" Then    'muestra saldos mayores a CERO
                         rsReporte.Filter = "cantidad<>0 and  IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
                      Else
                         rsReporte.Filter = "IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
                      End If
                   Else
                      If lc_CodigoItem = "1" Then     'muestra saldos mayores a CERO
                         rsReporte.Filter = "cantidad<>0"
                      End If
                   End If
                End If
            Else
                Set rsReporte = mo_ReglasFarmacia.SaldosSegunOrden(lnOrdenadoPor)
                If lc_CodigoItem = "1" Then         'muestra saldos mayores a CERO
                   rsReporte.Filter = "cantidad<>0"
                End If
            End If
            'debb-13/04/2016 (fin)
            If ml_idTipoSalidaBienInsumo = 0 And mb_SeMuestraLotes = False And rsReporte.RecordCount > 0 Then
               CargaSaldosAgrupados rsReporteAgrupado, rsReporte, lnOrdenadoPor
            End If
            'Reporte
            mflgContinuar = True
            If mb_SeMuestraLotes = True Then
               If lnOrdenadoPor = 1 Then
                  Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmSaldosDetNombre.rpt", 1)
               Else
                  Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmSaldosDet.rpt", 1)
               End If
            Else
               Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmSaldos.rpt", 1)
            End If
            ' Parametros del reporte
            Set crParamDefs = crReport.ParameterFields
            For Each crParamDef In crParamDefs
                Select Case crParamDef.ParameterFieldName
                   Case "IdAlmacen"
                        crParamDef.AddCurrentValue (lnIdAlmacen)
                    Case "Orden"
                        crParamDef.AddCurrentValue (lnOrdenadoPor)
                    Case "Filtro"
                        crParamDef.AddCurrentValue ("")
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
            If ml_idTipoSalidaBienInsumo = 0 And mb_SeMuestraLotes = False Then
               crReport.Database.SetDataSource rsReporteAgrupado
            Else
               crReport.Database.SetDataSource rsReporte
            End If
    Case "IciMensual"
            If Right(ml_HoraFin, 2) = "10" Then
sighentidades.ParaAuditoria = "59"
                GenerarRecordsetTemporalICI
sighentidades.ParaAuditoria = "p temp"
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!codigo = "00808"
                mrs_Tmp.Fields!nombre = "amox"
                mrs_Tmp.Fields!precio = 1
                mrs_Tmp.Fields!saldoI = 100
                mrs_Tmp.Fields!ingresos = 200
                mrs_Tmp.Fields!DevolucionesP = 0   'LnDevolucionesP
                mrs_Tmp.Fields!TotIngresos = 200
                mrs_Tmp.Fields!Ventas = 1
                mrs_Tmp.Fields!sis = 1
                mrs_Tmp.Fields!soat = 1
                mrs_Tmp.Fields!convenio = 1
                mrs_Tmp.Fields!creditoH = 1
                mrs_Tmp.Fields!defensaN = 1
                mrs_Tmp.Fields!OsDevol = 1
                mrs_Tmp.Fields!OsVencim = 1
                mrs_Tmp.Fields!OsMerma = 1
                mrs_Tmp.Fields!Exonerac = 1
                mrs_Tmp.Fields!IntervencionS = 1
                mrs_Tmp.Fields!otrasS = 1
                mrs_Tmp.Fields!TotSalidas = 1
                mrs_Tmp.Fields!fechaVencimiento = Date
                mrs_Tmp.Fields!tipo = "I"
                mrs_Tmp.Update
sighentidades.ParaAuditoria = "p update"
            Else
                Set mrs_Tmp = mRs_Record.Clone
            End If
sighentidades.ParaAuditoria = "clone"
            crReport.Database.SetDataSource mrs_Tmp
sighentidades.ParaAuditoria = "setdataso"
            If mrs_Tmp.RecordCount > 0 Then
sighentidades.ParaAuditoria = "sort"
                'Impresion
'                If lnOrdenadoPor = 0 Then
'                   mrs_Tmp.Sort = "codigo"
'                Else
'                   mrs_Tmp.Sort = "nombre"
'                End If
sighentidades.ParaAuditoria = "paso sort " & IIf(mb_EsDonaciones = True, "true", "false")
               'Reporte
                mflgContinuar = True
                If mb_EsDonaciones = True Then
sighentidades.ParaAuditoria = Trim(str(mrs_Tmp.RecordCount)) & "  rutad " & App.Path
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmICId.rpt", 1)
                Else
sighentidades.ParaAuditoria = Trim(str(mrs_Tmp.RecordCount)) & "  ruta: " & App.Path
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmICI.rpt", 1)
                   
                End If
sighentidades.ParaAuditoria = "paso plantilla"
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "pRecetas"
                            crParamDef.AddCurrentValue (lcTexto3)
                         Case "lcEESS"
                             crParamDef.AddCurrentValue (lcTitEESS)
                         Case "lcEESSdireccion"
                             crParamDef.AddCurrentValue (lcTitDireccion)
                         Case "lcEESStelefono"
                             crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
sighentidades.ParaAuditoria = "parametros"
                crReport.Database.SetDataSource mrs_Tmp
            Else
                MsgBox "no hay datos", vbInformation, "ICI"
            End If
    Case "rICI"
            GrabaParametro206 "antes de Procesar"
            ProcesarDatosICI
            GrabaParametro206 "despues de procesar"
'            Set crParamDefs = crReport.ParameterFields
'            For Each crParamDef In crParamDefs
'                Select Case crParamDef.ParameterFieldName
'                    Case "subTitulo"
'                        crParamDef.AddCurrentValue (lc_TextoDelFiltro)
'                        GrabaParametro206 "despues de subtitulo"
'                    Case "pRecetas"
'                        crParamDef.AddCurrentValue (lcTexto3)
'                        GrabaParametro206 "despues de precetas"
'                     Case "lcEESS"
'                         crParamDef.AddCurrentValue (lcTitEESS)
'                         GrabaParametro206 "despues de lcEESS"
'                     Case "lcEESSdireccion"
'                         crParamDef.AddCurrentValue (lcTitDireccion)
'                         GrabaParametro206 "despues de lcEESSdireccion"
'                     Case "lcEESStelefono"
'                         crParamDef.AddCurrentValue (lcTitTelefono)
'                         GrabaParametro206 "despues de lcessTelefono"
'                End Select
'            Next
'            GrabaParametro206 "paso parametros"
'            crReport.Database.SetDataSource mrs_Tmp
            If mrs_Tmp.RecordCount > 0 Then
                'Impresion
                If lnOrdenadoPor = 0 Then
                   mrs_Tmp.Sort = "codigo"
                Else
                   mrs_Tmp.Sort = "nombre"
                End If
               'Reporte
                mflgContinuar = True
                If mb_EsDonaciones = True Then
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmICId.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmICI.rpt", 1)
                End If
                GrabaParametro206 "paso app.path\rpt"
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                            GrabaParametro206 "subTitulo2"
                        Case "pRecetas"
                            crParamDef.AddCurrentValue (lcTexto3)
                            GrabaParametro206 "precetas_:2"
                         Case "lcEESS"
                             crParamDef.AddCurrentValue (lcTitEESS)
                             GrabaParametro206 "lceess_2"
                         Case "lcEESSdireccion"
                             crParamDef.AddCurrentValue (lcTitDireccion)
                             GrabaParametro206 "lceessdirecion2"
                         Case "lcEESStelefono"
                             crParamDef.AddCurrentValue (lcTitTelefono)
                             GrabaParametro206 "lceesstelefono2"
                    End Select
                Next
                GrabaParametro206 "paso parametros_2"
                crReport.Database.SetDataSource mrs_Tmp
                GrabaParametro206 "paso setdataSource"
            Else
                MsgBox "no hay datos", vbInformation, "ICI"
            End If
    Case "rPdiario"
            'ParteDiario
            GenerarRecordsetTemporalICI
            If mb_VtaYestrategicoSeparado = True Then
               ParteDiarioSeparadandoMovimientosVtasYestrategicos
            Else
               ParteDiario
            End If
            If mrs_Tmp.RecordCount > 0 Then
                'Impresion
                If lnOrdenadoPor = 0 Then
                   mrs_Tmp.Sort = "codigo"
                Else
                   mrs_Tmp.Sort = "nombre"
                End If
               'Reporte
                mflgContinuar = True
'                'Nro Devoluciones
                Set rsTmp = mo_ReglasFarmacia.FarmMovimientoFiltrarPorFechasYtipoConcepto(oConexion, mda_FechaInicio, mda_FechaFin, 21)
                lnSalidas = rsTmp.RecordCount
                'Nro Recetas totales y anuladas
                Set rsTmp = mo_ReglasFarmacia.FarmMovimientoVentasFiltrarPorFechas(oConexion, mda_FechaInicio, mda_FechaFin)
                lnSaldoInicial = 0: lnIdProducto = 0
                If rsTmp.RecordCount > 0 Then
                   rsTmp.MoveFirst
                   Do While Not rsTmp.EOF
                      If rsTmp.Fields!idEstadoMovimiento = 0 Then
                         lnSaldoInicial = lnSaldoInicial + 1
                      Else
                         lnIdProducto = lnIdProducto + 1
                      End If
                      rsTmp.MoveNext
                   Loop
                End If
                lcTexto3 = "Existen " + Trim(str(lnIdProducto)) + " recetas registradas," + Trim(str(lnSalidas)) + " Devoluciones y " + Trim(str(lnSaldoInicial)) + " anuladas"
                '
                Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmICIsinRecalculo.rpt", 1)
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "pRecetas"
                            crParamDef.AddCurrentValue (lcTexto3)
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
    Case "rIDI"
            oConexion.Open sighentidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            'Filtra los Datos
            Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(CDate("01/01/1990"), mda_FechaFin, 0, "")
            lnTotalRegistros = rsReporte.RecordCount
            
            If lnTotalRegistros > 0 Then
                Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                GenerarRecordsetTemporalIDI
                ProcesaDatosIDI oConexion
                mo_ReglasFarmacia.ActualizaImporteDeCabeceraMovimientos mda_FechaInicio, mda_FechaFin  'ojo
               'Reporte
                mflgContinuar = True
                If mb_EsDonaciones = True Then
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmIDId.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmIDI.rpt", 1)
                End If
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
            oConexion.Close
    Case "rRecetasXusuario"
            lnConsumoFarmacia1 = 0
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(mda_FechaInicio, mda_FechaFin, lnIdAlmacen, "S")
            lnTotalRegistros = rsReporte.RecordCount
            If lnTotalRegistros = 0 Then
                MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
                Screen.MousePointer = vbDefault
                Exit Sub
                
            Else
                Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                GenerarRecordsetTemporalRecetaXusuario
                rsReporte.MoveFirst
                lbPrimeraVez = True
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    lnPrecio = rsReporte.Fields!precio
                    '*******Saldo Inicial********
                    lnSaldoInicial = 0
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaInicio
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    '****** Movimientos en el Rango de Fechas***********
                    lnIngresos = 0: LnDevolucionesP = 0: TotIngresos = 0
                    LnVentas = 0: lnSis = 0: lnSoat = 0: LnConvenio = 0: lnCreditoH = 0: lnDefensaN = 0
                    LnOsDevol = 0: LnOsVencim = 0: LnOsMerma = 0: LnExonerac = 0: LnIntervencionS = 0
                    LnOtrasS = 0: TotSalidas = 0
                    If Not rsReporte.EOF Then
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                           lbContinuar = False
                           If ml_idUsuario > 0 Then
                                If rsReporte.Fields!idUsuario = ml_idUsuario Then
                                   lbContinuar = True
                                End If
                           Else
                                lbContinuar = True
                           End If
                           If lbContinuar = True Then
                                If ml_IdTipoFinanciamiento > 0 Then
                                    Set mrs_Tmp3 = mo_ReglasFarmacia.FarmMovimientoVentasFiltrarMovnumeroIdTipoFinanc(rsReporte.Fields!movNumero, rsReporte!MovTipo, ml_IdTipoFinanciamiento, oConexion)
                                    If mrs_Tmp3.RecordCount > 0 Then
                                       lbContinuar = True
                                    Else
                                       lbContinuar = False
                                    End If
                                    mrs_Tmp3.Close
                                Else
                                    lbContinuar = True
                                End If
                                If lbContinuar = True And mb_SoloBoletas = True Then
                                    Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarXMovimiento("S", rsReporte!movNumero, oConexion)
                                    If mrs_Tmp3.RecordCount > 0 Then
                                       If mrs_Tmp3!idPreVenta > 0 Then
                                       Else
                                          lbContinuar = False
                                       End If
                                    End If
                                    mrs_Tmp3.Close
                                End If
                                
                           End If
                           If lbContinuar = True Then
                                lbEsNuevoDocumento = False
                                If lbPrimeraVez = True Then
                                   lbPrimeraVez = False
                                   mrs_tmp1.AddNew
                                   mrs_tmp1.Fields!movNumero = rsReporte.Fields!DocumentoNumero
                                   mrs_tmp1.Update
                                   lbEsNuevoDocumento = True
                                Else
                                   mrs_tmp1.MoveFirst
                                   mrs_tmp1.Find "movNumero = '" & rsReporte.Fields!DocumentoNumero & "'"
                                   If mrs_tmp1.EOF Then
                                         mrs_tmp1.AddNew
                                         mrs_tmp1.Fields!movNumero = rsReporte.Fields!DocumentoNumero
                                         mrs_tmp1.Update
                                         lbEsNuevoDocumento = True
                                   End If
                                End If
                                'Redondeo
                                If mb_SoloBoletas = True And lbEsNuevoDocumento = True Then
                                    lcSerieB = Left(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") - 1)
                                    lcDocumentoB = Mid(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") + 1)
                                    Set rsTmp111 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(lcSerieB, lcDocumentoB)
                                    If rsTmp111.RecordCount > 0 Then
                                       rsTmp111.MoveFirst
                                       Do While Not rsTmp111.EOF
                                          If Format(rsTmp111.Fields!fechaCobranza, sighentidades.DevuelveFechaSoloFormato_DMY) = Format(rsReporte!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) Then
                                             lnTotalBol = rsTmp111!Total
                                             lnTotalExo = rsTmp111!exoneraciones
                                             lnTotalAde = rsTmp111!Adelantos
                                            
                                             lbTienePagoAcuenta = mo_ReglasFacturacion.ChequeaSiEsPagosAcuenta(rsTmp111!IdComprobantePago, _
                                                                                        oConexion, 0, lnRedondeoB, _
                                                                                        rsTmp111!IdTipoOrden, lnTotalExo, _
                                                                                        lnTotalAde, rsTmp111!IdEstadoComprobante, _
                                                                                        lnTotalBol)
                                             lnConsumoFarmacia1 = lnConsumoFarmacia1 + lnRedondeoB
                                             Exit Do
                                          End If
                                          rsTmp111.MoveNext
                                       Loop
                                    End If
                                    rsTmp111.Close
                                End If
                                TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                           End If
                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                           rsReporte.MoveNext
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
                    If TotSalidas > 0 Then
                        '
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!codigo = lcCodigo
                        mrs_Tmp.Fields!nombre = lcNombre
                        mrs_Tmp.Fields!TotSalidas = TotSalidas
                        mrs_Tmp.Fields!precio = lnPrecio
                        mrs_Tmp.Update
                    End If
                    If Not rsReporte.EOF Then
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                           rsReporte.MoveNext
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                If mrs_Tmp.RecordCount = 0 Then
                     mflgContinuar = False
                Else
                     lcTexto1 = "": LnVentas = 0
                     mrs_tmp1.Sort = "movNumero"
                     mrs_tmp1.MoveFirst
                     Do While Not mrs_tmp1.EOF
                        lcTexto1 = lcTexto1 + mrs_tmp1.Fields!movNumero + ", "
                        LnVentas = LnVentas + 1
                        mrs_tmp1.MoveNext
                        If LnVentas > 20 Then
                           Exit Do
                        End If
                     Loop
                     '
                     lcTexto2 = "": LnVentas = 0
                     Do While Not mrs_tmp1.EOF
                        lcTexto2 = lcTexto2 + mrs_tmp1.Fields!movNumero + ", "
                        LnVentas = LnVentas + 1
                        mrs_tmp1.MoveNext
                        If LnVentas > 20 Then
                           Exit Do
                        End If
                     Loop
                     '
                     lcTexto3 = "": LnVentas = 0
                     Do While Not mrs_tmp1.EOF
                        lcTexto3 = lcTexto3 + mrs_tmp1.Fields!movNumero + ", "
                        LnVentas = LnVentas + 1
                        mrs_tmp1.MoveNext
                        If LnVentas > 20 Then
                           Exit Do
                        End If
                     Loop
                    'Reporte
                    mrs_Tmp.Sort = "nombre"
                     mflgContinuar = True
                     Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmRecetasXusuario.rpt", 1)
                     ' Parametros del reporte
                     Set crParamDefs = crReport.ParameterFields
                     For Each crParamDef In crParamDefs
                         Select Case crParamDef.ParameterFieldName
                             Case "Receta1"
                                 crParamDef.AddCurrentValue (lcTexto1)
                             Case "Receta2"
                                 crParamDef.AddCurrentValue (lcTexto2)
                             Case "Receta3"
                                 crParamDef.AddCurrentValue (lcTexto3)
                             Case "subTitulo"
                                 crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                             'debb-setiembre2014****inicio
                             Case "Receta4"
                                 crParamDef.AddCurrentValue (lnConsumoFarmacia1)
                             'debb-setiembre2014****fin
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
           End If
           oConexion.Close
    Case "rKardex"
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosDeProducto(ml_idProducto, mda_FechaFin)
        If ml_idTipoSalidaBienInsumo > 0 Then
           rsReporte.Filter = "IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
        End If
        lnTotalRegistros = rsReporte.RecordCount
        
        If lnTotalRegistros = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
            GenerarRecordsetTemporalKARDEX
            lnSaldoInicial = 0
            'Saldo Inicial
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF And rsReporte.Fields!fechaCreacion < mda_FechaInicio
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                    lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                  End If
               Else
                  If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                  End If
               End If
               Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            lnSaldofinal = lnSaldoInicial
            mrs_Tmp.AddNew
            mrs_Tmp.Fields!movNumero = "<<Saldo>>"
            mrs_Tmp.Fields!ingresos = lnSaldoInicial
            mrs_Tmp.Fields!saldo = lnSaldoInicial
            mrs_Tmp.Update
            '
            mrs_Tmp99.AddNew
            mrs_Tmp99.Fields!Concepto = "Saldo Inicial"
            mrs_Tmp99.Fields!saldo = lnSaldoInicial
            mrs_Tmp99.Update
            '
            Do While Not rsReporte.EOF
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                    lnSaldofinal = lnSaldofinal - rsReporte.Fields!Cantidad
                    lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                    '
                    mrs_Tmp99.MoveFirst
                    mrs_Tmp99.Find "concepto='" & rsReporte.Fields!Concepto & "'"
                    If mrs_Tmp99.EOF Then
                       mrs_Tmp99.AddNew
                       mrs_Tmp99.Fields!Concepto = rsReporte.Fields!Concepto
                       mrs_Tmp99.Fields!salidas = rsReporte.Fields!Cantidad
                    Else
                       mrs_Tmp99.Fields!salidas = mrs_Tmp99.Fields!salidas + rsReporte.Fields!Cantidad
                    End If
                    mrs_Tmp99.Update
                    
                    'debb-03/11/2015 (inicio)
                    lcTexto3 = Trim(rsReporte.Fields!Concepto)
                    Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarXMovimiento("S", rsReporte!movNumero, oConexion)
                    If mrs_Tmp3.RecordCount > 0 Then
                       LnDic = mrs_Tmp3!idCuentaAtencion
                       mrs_Tmp3.Close
                       Set mrs_Tmp3 = mo_ReglasAdmision.AtencionesFiltraDatosCabecera(LnDic, oConexion)
                       If mrs_Tmp3.RecordCount > 0 Then
                          lcTexto3 = Left(lcTexto3, 18) & " <Actual: " & Trim(mrs_Tmp3!dFuenteFinanciamiento) & ">"
                       End If
                    End If
                    mrs_Tmp3.Close
                    'debb-03/11/2015 (fin)
                    
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                    mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
                    mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                    mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                    mrs_Tmp.Fields!salidas = rsReporte.Fields!Cantidad
                    mrs_Tmp.Fields!saldo = lnSaldofinal
                    mrs_Tmp.Fields!abreviatura = rsReporte.Fields!abreviatura
                    mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                    mrs_Tmp.Fields!Concepto = Left(lcTexto3, 100)                                     'debb-03/11/2015
                   
                    'debb-17/08/2015 (inicio)
                    If Left(rsReporte!Concepto, 5) = "VENTA" And InStr(rsReporte!DocumentoNumero, "-") > 0 Then
                       lcSerieB = Trim(Left(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") - 1))
                       lcDocumentoB = Trim(Mid(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") + 1, 100))
                       Set mrs_Tmp3 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(lcSerieB, lcDocumentoB)
                       If mrs_Tmp3.RecordCount > 0 Then
                          If Not IsNull(mrs_Tmp3!razonSocial) Then
                             mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino & " " & Trim(mrs_Tmp3!razonSocial), 100)
                          Else
                             mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino, 100)
                          End If
                       Else
                          mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino, 100)
                       End If
                       mrs_Tmp3.Close
                    Else
                       mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino & " " & rsReporte.Fields!Datpaciente, 100)
                    End If
                    'debb-17/08/2015 (fin)
                    mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                    mrs_Tmp.Fields!fechaVencimiento = rsReporte.Fields!fechaVencimiento
                    mrs_Tmp.Update
                  End If
               Else
                  If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                        lnSaldofinal = lnSaldofinal + rsReporte.Fields!Cantidad
                        lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                        '
                        
                        lcTexto3 = ""
                        Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoNotaIngresoSeleccionarXmovimiento(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                        If mrs_Tmp3.RecordCount > 0 Then
                           If Not IsNull(mrs_Tmp3!abreviatura) Then
                              lcTexto3 = Trim(mrs_Tmp3!abreviatura)
                           End If
                           If Not IsNull(mrs_Tmp3!oRigenNumero) Then
                              lcTexto3 = lcTexto3 & " " & Trim(mrs_Tmp3!oRigenNumero)
                           End If
                           If lcTexto3 <> "" Then
                              lcTexto3 = " (" & lcTexto3 & ")"
                           End If
                        End If
                        '
                        mrs_Tmp99.MoveFirst
                        mrs_Tmp99.Find "concepto='" & rsReporte.Fields!Concepto & "'"
                        If mrs_Tmp99.EOF Then
                           mrs_Tmp99.AddNew
                           mrs_Tmp99.Fields!Concepto = rsReporte.Fields!Concepto
                           mrs_Tmp99.Fields!ingresos = rsReporte.Fields!Cantidad
                        Else
                           mrs_Tmp99.Fields!ingresos = mrs_Tmp99.Fields!ingresos + rsReporte.Fields!Cantidad
                        End If
                        mrs_Tmp99.Update
                        '
                        mrs_Tmp3.Close
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
                        mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                        mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                        mrs_Tmp.Fields!ingresos = rsReporte.Fields!Cantidad
                        mrs_Tmp.Fields!saldo = lnSaldofinal
                        mrs_Tmp.Fields!abreviatura = rsReporte.Fields!abreviatura
                        mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                        mrs_Tmp.Fields!Concepto = Trim(rsReporte.Fields!Concepto)
                        mrs_Tmp.Fields!fOrigen = Left(Trim(rsReporte.Fields!fOrigen) & lcTexto3, 100)
                        mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                        mrs_Tmp.Fields!fechaVencimiento = rsReporte.Fields!fechaVencimiento
                        
                        mrs_Tmp.Update
                  End If
               End If
               Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            If mrs_Tmp.RecordCount = 0 Then
                 mflgContinuar = False
            Else
                'Reporte
                 mflgContinuar = True
                 Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmKARDEX.rpt", 1)
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
        End If
        oConexion.Close
    Case "rMovimientoES"
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set rsReporte = mo_ReglasFarmacia.FarmDevuelveMovimientos(lnIdAlmacen, ml_MovTipo, mda_FechaInicio, mda_FechaFin)
        lcTexto1 = ""
        
        If ml_IdConcepto > 0 Then
           lcTexto1 = lcTexto1 & " IdTipoConcepto=" & ml_IdConcepto & " and "
        End If
        If ml_IdEstado <> 2 Then
           lcTexto1 = lcTexto1 & " IdEstadoMovimiento=" & ml_IdEstado & " and "
        End If
        If lnIdAlmacenOrigen > 0 Then
           lcTexto1 = lcTexto1 & " IdAlmacenOrigen=" & lnIdAlmacenOrigen & " and "
        End If
        If lnIdAlmacenDestino > 0 Then
           lcTexto1 = lcTexto1 & " IdAlmacenDestino=" & lnIdAlmacenDestino & " and "
        End If
        If ml_idUsuario > 0 Then
           lcTexto1 = lcTexto1 & " IdUsuario=" & ml_idUsuario & " and "
        End If
        If lcTexto1 <> "" Then
           lcTexto1 = Left(lcTexto1, Len(lcTexto1) - 5)
           rsReporte.Filter = lcTexto1
        End If
        If rsReporte.RecordCount = 0 Then
            mflgContinuar = False
        Else
            GenerarRecordsetTemporalKARDEX
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lbContinuar = True
               lcTexto1 = ""
               If rsReporte.Fields!MovTipo = "S" Then
                     Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoVentasFiltrarMovnumero(rsReporte.Fields!movNumero)
                     If mrs_Tmp3.RecordCount > 0 Then
                        lcTexto1 = " (Pac: " & Trim(mrs_Tmp3.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp3.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp3.Fields!PrimerNombre) & ")"
                     Else
                        lcTexto1 = ""
                     End If
                     mrs_Tmp3.Close
                End If
                If lbContinuar = True Then
                    lcTexto3 = ""
                    If rsReporte.Fields!MovTipo = "E" Then
                        Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoNotaIngresoSeleccionarXmovimiento(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                        If mrs_Tmp3.RecordCount > 0 Then
                           If Not IsNull(mrs_Tmp3!abreviatura) Then
                              lcTexto3 = Trim(mrs_Tmp3!abreviatura)
                           End If
                           If Not IsNull(mrs_Tmp3!oRigenNumero) Then
                              lcTexto3 = lcTexto3 & " " & Trim(mrs_Tmp3!oRigenNumero)
                           End If
                           If lcTexto3 <> "" Then
                              lcTexto3 = " (" & lcTexto3 & ")"
                           End If
                        End If
                    End If
                    '
                    
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                    mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
                    mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                    mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                    mrs_Tmp.Fields!abreviatura = rsReporte.Fields!abreviatura
                    mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                    mrs_Tmp.Fields!Concepto = rsReporte.Fields!Concepto
                    mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fOrigen & lcTexto3, 100)
                    mrs_Tmp.Fields!fDestino = Trim(rsReporte.Fields!fDestino) & lcTexto1
                    mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
                    mrs_Tmp.Fields!Total = rsReporte.Fields!Total
                    mrs_Tmp.Update
                End If
                rsReporte.MoveNext
            Loop
            'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmMovimientoES.rpt", 1)
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
        oConexion.Close
    Case "rConsumoPromAnual"
        mda_FechaInicio = CDate("01/01/" + Trim(str(ml_IdAnio) + " 00:00:01"))
        mda_FechaFin = CDate("31/12/" + Trim(str(ml_IdAnio)) + " 23:59:59")
        GenerarRecordsetTemporalConsPromAnual
        Set mrs_tmp1 = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idtipoSuministro='01'")
        rConsumoPromAnual_ProcesaTemporal
        If mrs_Tmp.RecordCount = 0 Then
             mflgContinuar = False
        Else
             'Stock de Almacenes
             Set mrs_tmp1 = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idtipoLocales='A' and idTipoSuministro='01' and idEstado=1")
             lnIdAlmacen = mrs_tmp1.Fields!IdAlmacen
             mrs_Tmp.MoveFirst
             Do While Not mrs_Tmp.EOF
                lnDivision = 0
                If LnEne > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnFeb > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnMar > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnAbr > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnMay > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnJun > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnJul > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnAgo > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnSet > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnOct > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnNov > 0 Then
                    lnDivision = lnDivision + 1
                End If
                If LnDic > 0 Then
                    lnDivision = lnDivision + 1
                End If
                
                Set mrs_tmp1 = mo_ReglasFarmacia.FarmDevuelveSaldosSinLotesSegunAlmacen(lnIdAlmacen, 0, Trim(mrs_Tmp.Fields!codigo))
                If lnDivision <= 0 Then
                   lnPagoNeto = 0
                Else
                   lnPagoNeto = Round(mrs_Tmp.Fields!Total / lnDivision, 2)
                End If
                lnSalidas = IIf(mrs_tmp1.RecordCount > 0, mrs_tmp1.Fields!saldo, 0)
                If lnPagoNeto <= 0 Then
                   lnBruto = 0
                Else
                   lnBruto = Round(lnSalidas / lnPagoNeto, 2)
                End If
                mrs_Tmp.Fields!totalAlm = lnSalidas
                mrs_Tmp.Fields!promedio = lnPagoNeto
                mrs_Tmp.Fields!mesesExistencia = lnBruto
                mrs_Tmp.Fields!Estado = IIf(lnBruto > 11, "Stock Critico", IIf(lnBruto > 5, "Sobre Stock", IIf(lnBruto >= 2, "Normal Stock", IIf(lnBruto >= 1, "Sub Stock", "Desabastecimiento"))))
                mrs_Tmp.Update
                mrs_Tmp.MoveNext
             Loop
             'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmConsumoPromAnual.rpt", 1)
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
    Case "rConsumoPorCuenta"
    
    Case "rMontosPorPlan"
    Case "rProductoPorVencer"
    Case "rRecetasXservicio"
        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoVentasFiltraTipoServicioHosp(mda_FechaInicio, mda_FechaFin, lnIdAlmacen)
    Case "NiNs"
    Case "rAuditoriaFarmacia"
    Case "Inventario"
        Set rsReporte = mo_ReglasFarmacia.FarmInventarioSeleccionarXdocumento(ml_Documento, lnIdAlmacenDestino)
        If rsReporte.RecordCount = 0 Then
             mflgContinuar = False
        Else
             'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\farmInventario.rpt", 1)
             ' Parametros del reporte
             Set crParamDefs = crReport.ParameterFields
             For Each crParamDef In crParamDefs
                 Select Case crParamDef.ParameterFieldName
                     Case "subTitulo"
                         crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                     Case "lcAlmacenDestino"
                         crParamDef.AddCurrentValue (ml_Almacen)
                     Case "LcNumeroNI"
                         crParamDef.AddCurrentValue (ml_Documento)
                     Case "lcAlmacenOrigen"
                         crParamDef.AddCurrentValue (ml_AlmacenO)
                     Case "lcFechaNI"
                         crParamDef.AddCurrentValue (ml_HoraInicio)
                     Case "lcDocumento"
                         crParamDef.AddCurrentValue ("Inventario N° " & ml_Documento)
                     Case "Total"
                         crParamDef.AddCurrentValue (ml_Importe)
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
    Case "rConsumoXservicio"
        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoVentasFiltrarFechasAlmacIdTipoServ1(mda_FechaInicio, mda_FechaFin, _
                               lc_TipoServicioHosp, lnIdAlmacen)
        If rsReporte.RecordCount = 0 Then
            mflgContinuar = False
        Else
            'Reporte
            mflgContinuar = True
            Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmConsumoXservicio.rpt", 1)
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
            crReport.Database.SetDataSource rsReporte
        End If
    End Select
    GrabaParametro206 "paso case"
    CrvReportes.Top = 0
    If mflgContinuar = True Then
       If mb_EnArchivoExcel = True Then
            If lcBuscaParametro.SeleccionaFilaParametro(284) = "S" Then
                 'Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
                 Select Case lc_TipoReporte
                    Case "rSaldosPorAlmacen"
                        If ml_idTipoSalidaBienInsumo = 0 And mb_SeMuestraLotes = False Then
                           mo_ReglasReportes.ExportarRecordSetAexcel rsReporteAgrupado, "Saldos x Almacen", lc_TextoDelFiltro, "", Me.hwnd
                        Else
                           mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Saldos x Almacen", lc_TextoDelFiltro, "", Me.hwnd
                        End If
                    Case "rICI"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Formato ICI", lc_TextoDelFiltro, "", Me.hwnd
                        
                    Case "rPdiario"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Parte Diario", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rIDI"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Formato IDI", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rRecetasXusuario"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Recetas por usuario del Sistema", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rKardex"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Kardex", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rMovimientoES"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Movimientos de Entrada y Salida", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rConsumoPromAnual"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Consumo Promedio Anual", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rConsumoPorCuenta"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Consumo de Pacientes por N° Cuenta", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rMontosPorPlan"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Montos según Plan", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rProductoPorVencer"
                        mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Productos por Vencer", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rRecetasXservicio"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Recetas por Servicio", lc_TextoDelFiltro, "", Me.hwnd
                    Case "NiNs"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Nota de Ingreso, Nota de Salida", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rAuditoriaFarmacia"
                        mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Auditoria", lc_TextoDelFiltro, "", Me.hwnd
                    Case "Inventario"
                        mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Inventario", lc_TextoDelFiltro, "", Me.hwnd
                    Case "rConsumoXservicio"
                        mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Consumo por Servicios", lc_TextoDelFiltro, "", Me.hwnd
                End Select
            Else
                Select Case lc_TipoReporte
                Case "Inventario"
                    mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Inventario", lc_TextoDelFiltro, "", Me.hwnd
                Case "rKardex"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp99, "Kardex", lc_TextoDelFiltro, "", Me.hwnd
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Kardex", lc_TextoDelFiltro, "", Me.hwnd
                Case Else

                    crReport.ExportOptions.DestinationType = crEDTDiskFile
                    crReport.ExportOptions.FormatType = crEFTExcel70
                    crReport.ExportOptions.DiskFileName = lcBuscaParametro.SeleccionaFilaParametro(269)
                    crReport.Export (False)
                    MsgBox "Se generó el archivo " & lcBuscaParametro.SeleccionaFilaParametro(269)
                    If lc_TipoReporte = "rICI" And Val(lc_CodigoItem) > 0 Then
                       mo_ReglasReportes.ExportarRecordSetAexcel rsDebug, "Detalle de Movimientos", lc_TextoDelFiltro, "", Me.hwnd
                    End If
                End Select
            End If
        End If
        GrabaParametro206 "antes reportSource"
        CrvReportes.ReportSource = crReport
        GrabaParametro206 "antes viewReport"
        CrvReportes.ViewReport
        GrabaParametro206 "antes Zoom"
        CrvReportes.Zoom 120
        
        '
        GrabaParametro206 "antes de auditoria"
        mo_reglasComunes.grabaTablaAuditoria (crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))   'debb-27/05/2015
        GrabaParametro206 "paso auditoria"
    End If
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    'LimpiarVariablesDeMemoria
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
sighentidades.ParaAuditoria = "error: " & Err.Description

GrabaParametro206 "error " & Err.Description
    If lnErrorEnOdbc = 0 Then
       Resume Next
    ElseIf Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", _
            vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
    Resume

End Sub

Sub GrabaParametro206(lcTexto As String)
'sighentidades.ParaAuditoria = lcTexto
'    Dim oConexion1 As New Connection
'    Dim oRsTmp99 As New Recordset
'    oConexion1.CommandTimeout = 900
'    oConexion1.CursorLocation = adUseClient
'    oConexion1.Open sighentidades.CadenaConexion
'    oRsTmp99.Open "update parametros set valorTexto='" & lcTexto & "' where idParametro=374", oConexion1, adOpenKeyset, adLockOptimistic
'    oConexion1.Close
'    'oRsTmp99.Close
'    Set oConexion1 = Nothing
'    Set oRsTmp99 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
    LimpiarVariablesDeMemoria
End Sub
Private Sub Form_Resize()
    CrvReportes.Top = 500
    CrvReportes.Left = 0
    CrvReportes.Height = ScaleHeight
    CrvReportes.Width = ScaleWidth
    Me.ProgressBar1.Left = ScaleWidth - Me.ProgressBar1.Width - 1000
    'Me.ProgressBar1.Width = (Me.ProgressBar1.Width - 1500)
End Sub

Sub GenerarRecordsetTemporalICI()
    With mrs_Tmp
          
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 250, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "saldoI", adInteger, 4, adFldIsNullable
          .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "DevolucionesP", adInteger, 4, adFldIsNullable
          .Fields.Append "TotIngresos", adInteger, 4, adFldIsNullable
          .Fields.Append "ventas", adInteger, 4, adFldIsNullable
          .Fields.Append "sis", adInteger, 4, adFldIsNullable
          .Fields.Append "soat", adInteger, 4, adFldIsNullable
          .Fields.Append "convenio", adInteger, 4, adFldIsNullable
          .Fields.Append "creditoH", adInteger, 4, adFldIsNullable
          .Fields.Append "defensaN", adInteger, 4, adFldIsNullable
          .Fields.Append "OsDevol", adInteger, 4, adFldIsNullable
          .Fields.Append "OsVencim", adInteger, 4, adFldIsNullable
          .Fields.Append "OsMerma", adInteger, 4, adFldIsNullable
          .Fields.Append "Exonerac", adInteger, 4, adFldIsNullable
          .Fields.Append "IntervencionS", adInteger, 4, adFldIsNullable
          .Fields.Append "otrasS", adInteger, 4, adFldIsNullable
          .Fields.Append "TotSalidas", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
          .Fields.Append "tipo", adVarChar, 15, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub GenerarRecordsetTemporalIDI()
    With mrs_Tmp
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "saldoI", adInteger, 4, adFldIsNullable
          .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "DevolucionesP", adInteger, 4, adFldIsNullable
          .Fields.Append "ab", adInteger, 4, adFldIsNullable
          .Fields.Append "reingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "TotIngresos", adInteger, 4, adFldIsNullable
          .Fields.Append "Distribucion", adInteger, 4, adFldIsNullable
          .Fields.Append "Transferencia", adInteger, 4, adFldIsNullable
          .Fields.Append "DevolVencido", adInteger, 4, adFldIsNullable
          .Fields.Append "DevolMerma", adInteger, 4, adFldIsNullable
          .Fields.Append "VentaInst", adInteger, 4, adFldIsNullable
          .Fields.Append "Exoner", adInteger, 4, adFldIsNullable
          .Fields.Append "otrasS", adInteger, 4, adFldIsNullable
          .Fields.Append "TotSalidas", adInteger, 4, adFldIsNullable
          .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub GenerarRecordsetTemporalRecetaXusuario()
    With mrs_Tmp
          .Fields.Append "precio", adDouble
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "TotSalidas", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    With mrs_tmp1
          .Fields.Append "MovNumero", adVarChar, 20, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub GenerarRecordsetTemporalKARDEX()
    With mrs_Tmp
          .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
          .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
          .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
          .Fields.Append "MovNumero", adVarChar, 10, adFldIsNullable
          .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "salidas", adInteger, 4, adFldIsNullable
          .Fields.Append "saldo", adInteger, 4, adFldIsNullable
          .Fields.Append "Abreviatura", adVarChar, 10, adFldIsNullable
          .Fields.Append "DocumentoNumero", adVarChar, 20, adFldIsNullable
          .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
          .Fields.Append "fOrigen", adVarChar, 100, adFldIsNullable
          .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
          .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
          .Fields.Append "fDestino", adVarChar, 100, adFldIsNullable
          .Fields.Append "Estado", adVarChar, 30, adFldIsNullable
          .Fields.Append "Total", adDouble
          
          .LockType = adLockOptimistic
          .Open
    End With

    With mrs_Tmp99
          .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
          .Fields.Append "salidas", adInteger, 4, adFldIsNullable
          .Fields.Append "saldo", adInteger, 4, adFldIsNullable
          .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub GenerarRecordsetTemporalConsPromAnual()
    With mrs_Tmp
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "Ene", adInteger, 4, adFldIsNullable
          .Fields.Append "Feb", adInteger, 4, adFldIsNullable
          .Fields.Append "Mar", adInteger, 4, adFldIsNullable
          .Fields.Append "Abr", adInteger, 4, adFldIsNullable
          .Fields.Append "May", adInteger, 4, adFldIsNullable
          .Fields.Append "Jun", adInteger, 4, adFldIsNullable
          .Fields.Append "Jul", adInteger, 4, adFldIsNullable
          .Fields.Append "Ago", adInteger, 4, adFldIsNullable
          .Fields.Append "Set", adInteger, 4, adFldIsNullable
          .Fields.Append "Oct", adInteger, 4, adFldIsNullable
          .Fields.Append "Nov", adInteger, 4, adFldIsNullable
          .Fields.Append "Dic", adInteger, 4, adFldIsNullable
          .Fields.Append "total", adInteger, 4, adFldIsNullable
          .Fields.Append "totalAlm", adInteger, 4, adFldIsNullable
          .Fields.Append "Promedio", adDouble
          .Fields.Append "mesesExistencia", adDouble
          .Fields.Append "Estado", adVarChar, 50, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub GenerarRecordsetTemporalConsumoCUENTA()
    With mrs_Tmp
          .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
          .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
          .Fields.Append "MovNumero", adVarChar, 15, adFldIsNullable
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "cantidad", adInteger, 4, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "Estado", adVarChar, 20, adFldIsNullable
          .Fields.Append "dAlmacen", adVarChar, 100, adFldIsNullable
          .Fields.Append "dFinanciamiento", adVarChar, 50, adFldIsNullable
          .Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
          .Fields.Append "Tipo", adVarChar, 30, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub GenerarRecordsetTemporalMontoXplan()
    With mrs_Tmp
          .Fields.Append "idPlan", adInteger
          .Fields.Append "Plan", adVarChar, 150, adFldIsNullable
          .Fields.Append "mBruto", adDouble
          .Fields.Append "mDevoluciones", adDouble
          .Fields.Append "mPagoNeto", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub GenerarRecordsetTemporalRecetasXusuario()
    With mrs_Tmp
          .Fields.Append "id", adInteger, 4, adFldIsNullable
          .Fields.Append "dservicio", adVarChar, 50, adFldIsNullable
          .Fields.Append "MovNumero", adVarChar, 20, adFldIsNullable
          .Fields.Append "RecetaInterna", adInteger, 4, adFldIsNullable
          .Fields.Append "RecetaExterna", adInteger, 4, adFldIsNullable
          .Fields.Append "SinReceta", adInteger, 4, adFldIsNullable
          .Fields.Append "Total", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mrs_Tmp = Nothing
    Set mrs_tmp1 = Nothing
    Set mrs_Tmp2 = Nothing
    Set rsReporte = Nothing
    Set rsTmp = Nothing
    Set rsTmp1 = Nothing
    Set rsErrores = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasCaja = Nothing
    Set mo_ReglasReportes = Nothing
    Set oConexion = Nothing
    Set mo_DoFarmMovimientoVentas = Nothing
    Set oDoCatProductosHosp = Nothing
    Set oFarmMovimientoDetalle = Nothing
    Set oBuscaMovimientos = Nothing
    
End Sub



Sub ProcesaAuditoria()
             
             GenerarRecordsetTemporalKARDEX
             mrs_tmp1.MoveFirst
             Do While Not mrs_tmp1.EOF
                If UCase(Left(mrs_tmp1.Fields!Tabla, 14)) = "FARMMOVIMIENTO" Or UCase(Trim(mrs_tmp1.Fields!Tabla)) = "FARMPREVENTA" Then
                    lcTexto3 = IIf(mrs_tmp1.Fields!accion = "A", "Agregó", IIf(mrs_tmp1.Fields!accion = "M", "Modificó", "Anuló"))
                    lbPrimeraVez = True
                    If UCase(Left(mrs_tmp1.Fields!Tabla, 14)) = "FARMMOVIMIENTO" Then
                        lcTexto1 = Mid(mrs_tmp1.Fields!Tabla, 16, 1)
                        lcTexto2 = Right(Trim(mrs_tmp1.Fields!Tabla), 9)
                        dFinanciamiento = ""
                        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoFiltraPorMovNumeroTipo(lcTexto2, lcTexto1)
                        If rsReporte.RecordCount > 0 Then
                            If lcTexto1 = "S" Then
                               If lnIdAlmacen <> rsReporte.Fields!IdAlmacenOrigen Then
                                  lbPrimeraVez = False
                               End If
                            Else
                               If lnIdAlmacen <> rsReporte.Fields!IdAlmacenDestino Then
                                  lbPrimeraVez = False
                               End If
                            End If
                            If rsReporte.Fields!idTipoConcepto = 10 Then  'Contado
                                mo_DoFarmMovimientoVentas.movNumero = lcTexto2
                                mo_DoFarmMovimientoVentas.MovTipo = lcTexto1
                                If mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorId(mo_DoFarmMovimientoVentas) Then
                                   dFinanciamiento = "(Preventa: " & mo_DoFarmMovimientoVentas.idPreVenta & ")"
                                End If
                            End If
                            If lbPrimeraVez = True Then
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!fechaCreacion = Format(mrs_tmp1.Fields!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY)
                                mrs_Tmp.Fields!HoraCreacion = Format(mrs_tmp1.Fields!fechaHora, sighentidades.DevuelveHoraSoloFormato_HM)
                                mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                                mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                                mrs_Tmp.Fields!abreviatura = rsReporte.Fields!abreviatura
                                mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                                mrs_Tmp.Fields!Concepto = rsReporte.Fields!Concepto
                                If lcTexto1 = "S" Then
                                   mrs_Tmp.Fields!fOrigen = Trim(rsReporte.Fields!fDestino) & "    " & dFinanciamiento
                                Else
                                   mrs_Tmp.Fields!fOrigen = rsReporte.Fields!fOrigen
                                End If
                                mrs_Tmp.Fields!fDestino = Trim(mrs_tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_tmp1.Fields!Nombres) & "   (Pc: " & Trim(mrs_tmp1.Fields!nombrePc) & ")"
                                mrs_Tmp.Fields!Lote = lcTexto3
                                mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
                                mrs_Tmp.Fields!Total = rsReporte.Fields!Total
                                mrs_Tmp.Update
                            End If
                        End If
                    End If
                    If UCase(Trim(mrs_tmp1.Fields!Tabla)) = "FARMPREVENTA" Then
                        Set rsReporte = mo_ReglasFarmacia.FarmPreVentaFiltraPorIdPreVenta(mrs_tmp1.Fields!idRegistro)
                        If rsReporte.RecordCount > 0 Then
                            If lnIdAlmacen <> rsReporte.Fields!IdAlmacen Then
                               lbPrimeraVez = False
                            End If
                            If lbPrimeraVez = True Then
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!fechaCreacion = Format(mrs_tmp1.Fields!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY)
                                mrs_Tmp.Fields!HoraCreacion = Format(mrs_tmp1.Fields!fechaHora, sighentidades.DevuelveHoraSoloFormato_HM)
                                mrs_Tmp.Fields!movNumero = Trim(str(rsReporte.Fields!idPreVenta))
                                mrs_Tmp.Fields!abreviatura = ""
                                mrs_Tmp.Fields!DocumentoNumero = "PreVenta"
                                mrs_Tmp.Fields!Concepto = ""
                                mrs_Tmp.Fields!fOrigen = ""
                                mrs_Tmp.Fields!fDestino = Trim(mrs_tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_tmp1.Fields!Nombres) & "   (Pc: " & Trim(mrs_tmp1.Fields!nombrePc) & ")"
                                mrs_Tmp.Fields!Lote = lcTexto3
                                mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
                                mrs_Tmp.Fields!Total = rsReporte.Fields!Total
                                mrs_Tmp.Update
                            End If
                        End If
                    End If
                End If
                mrs_tmp1.MoveNext
             Loop

End Sub

Sub rConsumoPromAnual_ProcesaTemporal()
        mrs_tmp1.MoveFirst
        Do While Not mrs_tmp1.EOF
            lnIdAlmacen = mrs_tmp1.Fields!IdAlmacen
            Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(mda_FechaInicio, mda_FechaFin, lnIdAlmacen, "S")
            If mb_NOconsiderarTipoConceptos = True Then
               rsReporte.Filter = "idTipoConcepto<>4 and idTipoConcepto<>5 and idTipoConcepto<>6 and  " & _
                                 "idTipoConcepto<>7 and idTipoConcepto<>19 and idTipoConcepto<>20"
            End If
            If rsReporte.RecordCount > 0 Then
                
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                   LnEne = 0: LnFeb = 0: LnMar = 0: LnAbr = 0: LnMay = 0: LnJun = 0
                   LnJul = 0: LnAgo = 0: LnSet = 0: LnOct = 0: LnNov = 0: LnDic = 0
                   lnIdProducto = rsReporte.Fields!idProducto
                   lcCodigo = rsReporte.Fields!codigo
                   lcNombre = rsReporte.Fields!nombre
                   Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                        Select Case Month(rsReporte.Fields!fechaCreacion)
                        Case 1
                            LnEne = LnEne + rsReporte.Fields!Cantidad
                        Case 2
                            LnFeb = LnFeb + rsReporte.Fields!Cantidad
                        Case 3
                            LnMar = LnMar + rsReporte.Fields!Cantidad
                        Case 4
                            LnAbr = LnAbr + rsReporte.Fields!Cantidad
                        Case 5
                            LnMay = LnMay + rsReporte.Fields!Cantidad
                        Case 6
                            LnJun = LnJun + rsReporte.Fields!Cantidad
                        Case 7
                            LnJul = LnJul + rsReporte.Fields!Cantidad
                        Case 8
                            LnAgo = LnAgo + rsReporte.Fields!Cantidad
                        Case 9
                            LnSet = LnSet + rsReporte.Fields!Cantidad
                        Case 10
                            LnOct = LnOct + rsReporte.Fields!Cantidad
                        Case 11
                            LnNov = LnNov + rsReporte.Fields!Cantidad
                        Case 12
                            LnDic = LnDic + rsReporte.Fields!Cantidad
                        End Select
                        rsReporte.MoveNext
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                   Loop
                   lnSalidas = LnEne + LnFeb + LnMar + LnAbr + LnMay + LnJun + LnJul + LnAgo + LnSet + LnOct + LnNov + LnDic
                   lbPrimeraVez = True
                   If mrs_Tmp.RecordCount > 0 Then
                      mrs_Tmp.MoveFirst
                      mrs_Tmp.Find "codigo='" & lcCodigo & "'"
                      If Not mrs_Tmp.EOF Then
                         lbPrimeraVez = False
                      End If
                   End If
                   If lbPrimeraVez = True Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!codigo = lcCodigo
                        mrs_Tmp.Fields!nombre = lcNombre
                        mrs_Tmp.Fields!ene = LnEne
                        mrs_Tmp.Fields!feb = LnFeb
                        mrs_Tmp.Fields!mar = LnMar
                        mrs_Tmp.Fields!abr = LnAbr
                        mrs_Tmp.Fields!may = LnMay
                        mrs_Tmp.Fields!jun = LnJun
                        mrs_Tmp.Fields!jul = LnJul
                        mrs_Tmp.Fields!ago = LnAgo
                        mrs_Tmp.Fields!Set = LnSet
                        mrs_Tmp.Fields!Oct = LnOct
                        mrs_Tmp.Fields!nov = LnNov
                        mrs_Tmp.Fields!dic = LnDic
                        mrs_Tmp.Fields!Total = lnSalidas
                   Else
                        mrs_Tmp.Fields!ene = LnEne + mrs_Tmp.Fields!ene
                        mrs_Tmp.Fields!feb = LnFeb + mrs_Tmp.Fields!feb
                        mrs_Tmp.Fields!mar = LnMar + mrs_Tmp.Fields!mar
                        mrs_Tmp.Fields!abr = LnAbr + mrs_Tmp.Fields!abr
                        mrs_Tmp.Fields!may = LnMay + mrs_Tmp.Fields!may
                        mrs_Tmp.Fields!jun = LnJun + mrs_Tmp.Fields!jun
                        mrs_Tmp.Fields!jul = LnJul + mrs_Tmp.Fields!jul
                        mrs_Tmp.Fields!ago = LnAgo + mrs_Tmp.Fields!ago
                        mrs_Tmp.Fields!Set = LnSet + mrs_Tmp.Fields!Set
                        mrs_Tmp.Fields!Oct = LnOct + mrs_Tmp.Fields!Oct
                        mrs_Tmp.Fields!nov = LnNov + mrs_Tmp.Fields!nov
                        mrs_Tmp.Fields!dic = LnDic + mrs_Tmp.Fields!dic
                        mrs_Tmp.Fields!Total = lnSalidas + mrs_Tmp.Fields!Total
                   End If
                   mrs_Tmp.Update
                Loop
           End If
           mrs_tmp1.MoveNext
        Loop
        mrs_tmp1.Close
End Sub



Sub ParteDiario()
        Dim lnIdProducto As Long, lnSaldoInicial As Long
        Dim lnRegistro As Long
        Dim lnRegTope As Long
        Dim rsTmp11 As New Recordset
        Dim rsTmp12 As New Recordset
        Dim rsTmp13 As New Recordset
        Dim rsTmp14 As New Recordset
        Dim rsTmp15 As New Recordset
        Dim lnidTipoConceptoFarmacia As Long
        Dim lcSql As String
        '
        On Error GoTo ErrParteDia
        '
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oFarmMovimientoDetalle.Conexion = oConexion
        
        'Proceso
        lcUltDiaMes = Trim(str(sighentidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))))
        ldFechaHistoricoXmes = CDate("01" & Format(mda_FechaInicio, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
        ldFechaHistoricoXmes = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaHistoricoXmes)
        ldFechaInicioMovim = DateAdd("n", 1, ldFechaHistoricoXmes)
        'Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(CDate("01/01/1990"), mda_FechaFin, 0, "")
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDI(ldFechaInicioMovim, mda_FechaFin, 0, "")
        lnTotalRegistros = rsReporte.RecordCount
        
        If lnTotalRegistros > 0 Then
            Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
            '
            lnRegistro = 1
            lnRegTope = 28320
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
If Val(rsReporte.Fields!codigo) = 625 Then
lcCodigo = ""
End If
            
                lnIdProducto = rsReporte.Fields!idProducto
                lcCodigo = rsReporte.Fields!codigo
                lcNombre = rsReporte.Fields!nombre
                '*******Saldo Inicial****************************************
                lnSaldoInicial = 0
                'saldos-barre historico mensual
                For lnFor = 1 To Len(lc_AlmacenesParaICI)
                    If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                       lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                       lnFor = Len(lc_AlmacenesParaICI)
                    Else
                        lcTexto1 = ""
                        Do While True
                           If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                              Exit Do
                           Else
                              lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                              lnFor = lnFor + 1
                           End If
                        Loop
                        lnIdAlmacenRep = Val(lcTexto1)
                    End If
                    If lnIdAlmacenRep > 1 Then
                        Set rsErrores = mo_ReglasFarmacia.FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(lnIdProducto, lnIdAlmacenRep, ldFechaHistoricoXmes)
                        Do While Not rsErrores.EOF
                            lnSaldoInicial = lnSaldoInicial + rsErrores.Fields!saldo
                            rsErrores.MoveNext
                        Loop
                        rsErrores.Close
                    End If
                Next
                'saldos-barre movimiento
                Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaInicio
                   If rsReporte.Fields!MovTipo = "S" Then
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                        lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                      End If
                   Else
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                         lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                      End If
                   End If
                   rsReporte.MoveNext
                   lnRegistro = lnRegistro + 1
If lnRegistro > lnRegTope Then
lcTexto10 = "="
End If
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                '****** Movimientos en el Rango de Fechas***********************************
                lnIngresos = 0: LnDevolucionesP = 0: TotIngresos = 0
                LnVentas = 0: lnSis = 0: lnSoat = 0: LnConvenio = 0: lnCreditoH = 0: lnDefensaN = 0
                LnOsDevol = 0: LnOsVencim = 0: LnOsMerma = 0: LnExonerac = 0: LnIntervencionS = 0
                LnOtrasS = 0: TotSalidas = 0
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       lbPrimeraVez = True
                       If rsReporte.Fields!MovTipo = "S" Then
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                                If mb_ConsiderarRecalculo = True Then
                                    '********* con recalculo
                                    lcTexto1 = rsReporte.Fields!MovTipo
                                    lcTexto2 = rsReporte.Fields!movNumero
                                    'Busca si tiene Pagos
                                    'debb 02/02/2011
                                    Set rsTmp11 = mo_ReglasFarmacia.FacturacionBienesPagoSeleccionarPorMovNumeroProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp11.RecordCount > 0 Then
                                       If rsTmp11.Fields!IdComprobantePago > 0 And rsTmp11.Fields!idEstadoFacturacion = 4 Then
                                          lbPrimeraVez = False
                                          rsTmp11.MoveFirst
                                          Do While Not rsTmp11.EOF
                                                Select Case rsReporte.Fields!idTipoConcepto
                                                Case 10       'Ventas
                                                     LnVentas = LnVentas + rsTmp11.Fields!CantidadPagar
                                                Case 13       'Sis
                                                     lnSis = lnSis + rsTmp11.Fields!CantidadPagar
                                                Case 14       'Soat
                                                     lnSoat = lnSoat + rsTmp11.Fields!CantidadPagar
                                                Case 23, 26      'Convenios, Credito Personal
                                                     LnConvenio = LnConvenio + rsTmp11.Fields!CantidadPagar
                                                Case 17       'Credito Hospitalario
                                                     lnCreditoH = lnCreditoH + rsTmp11.Fields!CantidadPagar
                                               ' Case 22       'Defensa nacional
                                               '      lnDefensaN = lnDefensaN + rsTmp11.Fields!cantidadPagar
                                               ' Case 7       'Otras salidas Devolucion
                                               '      LnOsDevol = LnOsDevol + rsTmp11.Fields!cantidadPagar
                                               ' Case 5       'Otras salidas Vencimiento
                                               '      LnOsVencim = LnOsVencim + rsTmp11.Fields!cantidadPagar
                                               ' Case 6       'Otras salidas Merma
                                               '      LnOsMerma = LnOsMerma + rsTmp11.Fields!cantidadPagar
                                               ' Case 15       'Exoneraciones
                                               '      LnExonerac = LnExonerac + rsTmp11.Fields!cantidadPagar
                                                Case 16       'Intervencion Sanitaria
                                                     LnIntervencionS = LnIntervencionS + rsTmp11.Fields!CantidadPagar
                                                Case Else
                                                     LnOtrasS = LnOtrasS + rsTmp11.Fields!CantidadPagar
                                                End Select
                                                TotSalidas = TotSalidas + rsTmp11.Fields!CantidadPagar
                                                rsTmp11.MoveNext
                                          Loop
                                       End If
                                    End If
                                    rsTmp11.Close
                                    'Busca si tiene algun seguro o exoneracion (Plan)
                                    'debb 02/02/2011
                                    Set rsTmp12 = mo_ReglasFarmacia.FacturacionBienesFinancSeleccionarPorProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp12.RecordCount > 0 Then
                                       lbPrimeraVez = False
                                        rsTmp12.MoveFirst
                                        Do While Not rsTmp12.EOF
                                            Select Case mo_ReglasFacturacion.FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia(oConexion, rsTmp12.Fields!idFuenteFinanciamiento)
                                            Case 10       'Ventas
                                                 LnVentas = LnVentas + rsTmp12.Fields!CantidadFinanciada
                                            Case 13       'Sis
                                                 lnSis = lnSis + rsTmp12.Fields!CantidadFinanciada
                                            Case 14      'Soat
                                                 lnSoat = lnSoat + rsTmp12.Fields!CantidadFinanciada
                                            Case 16       'Intervencion Sanitaria
                                                 LnIntervencionS = LnIntervencionS + rsTmp12.Fields!CantidadFinanciada
                                            Case 17       'Credito Hospitalario
                                                 lnCreditoH = lnCreditoH + rsTmp12.Fields!CantidadFinanciada
                                            Case 23, 26       'Convenios, Credito Personal
                                                 LnConvenio = LnConvenio + rsTmp12.Fields!CantidadFinanciada
'                                            Case 0      'Exoneraciones
'                                                 If rsTmp12.Fields!cantidadFinanciada > 0 Then
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                 Else
'                                                    '**no se sabe la CANTIDAD EXONERADA solo el IMPORTE EXONERADO
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                    LnExonerac = LnExonerac - rsTmp12.Fields!cantidadFinanciada
'                                                    TotSalidas = TotSalidas - rsTmp12.Fields!cantidadFinanciada
'                                                 End If
                                            Case Else
                                                 LnOtrasS = LnOtrasS + rsTmp12.Fields!CantidadFinanciada
                                            End Select
                                            TotSalidas = TotSalidas + rsTmp12.Fields!CantidadFinanciada
                                            rsTmp12.MoveNext
                                        Loop
                                    End If
                                    rsTmp12.Close
                                    '
                                    If lbPrimeraVez = False Then
                                        Do While Not rsReporte.EOF And lcTexto1 = rsReporte.Fields!MovTipo And lcTexto2 = rsReporte.Fields!movNumero And lnIdProducto = rsReporte.Fields!idProducto
                                           rsReporte.MoveNext
                                           lnRegistro = lnRegistro + 1

                                           If rsReporte.EOF Then
                                              Exit Do
                                           End If
                                        Loop
                                    End If
                                End If
                                If lbPrimeraVez = True Then
                                    '******** sin recalculo
                                    Select Case rsReporte.Fields!idTipoConcepto
                                    Case 10       'Ventas
                                         LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                    Case 13       'Sis
                                         lnSis = lnSis + rsReporte.Fields!Cantidad
                                    Case 14       'Soat
                                         lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                    Case 23, 26      'Convenios, Credito Personal
                                         LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                    Case 17       'Credito Hospitalario
                                         lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
'                                    Case 22       'Defensa nacional
'                                         lnDefensaN = lnDefensaN + rsReporte.Fields!cantidad
'                                    Case 7       'Otras salidas Devolucion
'                                         LnOsDevol = LnOsDevol + rsReporte.Fields!cantidad
'                                    Case 5       'Otras salidas Vencimiento
'                                         LnOsVencim = LnOsVencim + rsReporte.Fields!cantidad
'                                    Case 6       'Otras salidas Merma
'                                         LnOsMerma = LnOsMerma + rsReporte.Fields!cantidad
'                                    Case 15       'Exoneraciones
'                                         LnExonerac = LnExonerac + rsReporte.Fields!cantidad
                                    Case 16       'Intervencion Sanitaria
                                         LnIntervencionS = LnIntervencionS + rsReporte.Fields!Cantidad
                                    Case Else
                                         LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                    End Select
                                    If rsReporte.Fields!IdAlmacenDestino = 10 And rsReporte.Fields!idTipoConcepto = 4 And mb_ConsideraOSH = False Then
                                       'destino=otros servicios hospital, tipoConcepto=distribucion
                                    Else
                                        TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                    End If
                                End If
                          End If
                       Else
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                                Select Case rsReporte.Fields!idTipoConcepto
                                Case 19        'Inventario
                                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                Case 21        'Devolucion de Pacientes
                                     LnDevolucionesP = LnDevolucionesP + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                Case Else        'Ingresos
                                     lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                End Select
                                
                           End If
                       End If
                       If lbPrimeraVez = True Then
                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                           rsReporte.MoveNext
                           lnRegistro = lnRegistro + 1

                       End If
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                       rsReporte.MoveNext
                       lnRegistro = lnRegistro + 1
If lnRegistro > lnRegTope Then
lcTexto10 = "="
End If
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                '
                lnPrecio = 0
                Set rsTmp13 = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                If rsTmp13.RecordCount > 0 Then
                   lnPrecio = rsTmp13.Fields!PrecioUnitario
                End If
                rsTmp13.Close
                '
                If InStr(lc_AlmacenesParaICI, "/") > 0 Then
                   lnIdAlmacen = Val(Left(lc_AlmacenesParaICI, InStr(lc_AlmacenesParaICI, "/") - 1)) 'toma la primera farmacia para las FECHAS DE VENCIMIENTO
                Else
                   lnIdAlmacen = Val(lc_AlmacenesParaICI)
                End If
                '
                lbContinua = True
                If mb_ConsiderarSinMovimientos = False Then
                   If TotIngresos = 0 And TotSalidas = 0 Then
                      lbContinua = False
                   End If
                End If
                '
                If lbContinua Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!codigo = lcCodigo
                    mrs_Tmp.Fields!nombre = lcNombre
                    mrs_Tmp.Fields!precio = lnPrecio
                    mrs_Tmp.Fields!saldoI = lnSaldoInicial
                    mrs_Tmp.Fields!ingresos = lnIngresos
                    mrs_Tmp.Fields!DevolucionesP = LnDevolucionesP
                    mrs_Tmp.Fields!TotIngresos = TotIngresos
                    mrs_Tmp.Fields!Ventas = LnVentas
                    mrs_Tmp.Fields!sis = lnSis
                    mrs_Tmp.Fields!soat = lnSoat
                    mrs_Tmp.Fields!convenio = LnConvenio
                    mrs_Tmp.Fields!creditoH = lnCreditoH
                    mrs_Tmp.Fields!defensaN = lnDefensaN
                    mrs_Tmp.Fields!OsDevol = LnOsDevol
                    mrs_Tmp.Fields!OsVencim = LnOsVencim
                    mrs_Tmp.Fields!OsMerma = LnOsMerma
                    mrs_Tmp.Fields!Exonerac = LnExonerac
                    mrs_Tmp.Fields!IntervencionS = LnIntervencionS
                    mrs_Tmp.Fields!otrasS = LnOtrasS
                    mrs_Tmp.Fields!TotSalidas = TotSalidas
                    mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                    mrs_Tmp.Update
                End If
                'Graba Datos en Temporal
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
       End If
       'oConexion.Close
       rsReporte.Close
        Set rsTmp11 = Nothing
        Set rsTmp12 = Nothing
        Set rsTmp13 = Nothing
        Set rsTmp14 = Nothing
        Set rsTmp15 = Nothing
        Exit Sub
ErrParteDia:
     MsgBox Err.Description
     Resume
End Sub

Sub ParteDiarioBuscaPagosYseguros(lnIdProducto101 As Long, lcMovTipo101 As String, lcMovNumero101 As String, lnIdTipoConcepto101 As Long, oConexion As Connection)
        Dim rsTmp11 As New Recordset
        Dim rsTmp12 As New Recordset
        'Busca si tiene Pagos
        Set rsTmp11 = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto(lcMovNumero101, lcMovTipo101, lnIdProducto101)
        If rsTmp11.RecordCount > 0 Then
           If rsTmp11.Fields!IdComprobantePago > 0 And rsTmp11.Fields!idEstadoFacturacion = 4 Then
              lbPrimeraVez = False
              rsTmp11.MoveFirst
              Do While Not rsTmp11.EOF
                    Select Case lnIdTipoConcepto101
                    Case 10       'Ventas
                         LnVentas = LnVentas + rsTmp11.Fields!CantidadPagar
                    Case 13       'Sis
                         lnSis = lnSis + rsTmp11.Fields!CantidadPagar
                    Case 14       'Soat
                         lnSoat = lnSoat + rsTmp11.Fields!CantidadPagar
                    Case 23       'Convenios
                         LnConvenio = LnConvenio + rsTmp11.Fields!CantidadPagar
                    Case 17       'Credito Hospitalario
                         lnCreditoH = lnCreditoH + rsTmp11.Fields!CantidadPagar
                    Case 22       'Defensa nacional
                         lnDefensaN = lnDefensaN + rsTmp11.Fields!CantidadPagar
                    Case 7       'Otras salidas Devolucion
                         LnOsDevol = LnOsDevol + rsTmp11.Fields!CantidadPagar
                    Case 5       'Otras salidas Vencimiento
                         LnOsVencim = LnOsVencim + rsTmp11.Fields!CantidadPagar
                    Case 6       'Otras salidas Merma
                         LnOsMerma = LnOsMerma + rsTmp11.Fields!CantidadPagar
                    Case 15       'Exoneraciones
                         LnExonerac = LnExonerac + rsTmp11.Fields!CantidadPagar
                    Case 16       'Intervencion Sanitaria
                         LnIntervencionS = LnIntervencionS + rsTmp11.Fields!CantidadPagar
                    Case Else
                         LnOtrasS = LnOtrasS + rsTmp11.Fields!CantidadPagar
                    End Select
                    TotSalidas = TotSalidas + rsTmp11.Fields!CantidadPagar
                    rsTmp11.MoveNext
              Loop
           End If
        End If
        rsTmp11.Close
        'Busca si tiene algun seguro o exoneracion (Plan)
        Set rsTmp12 = mo_ReglasFarmacia.oFacturacionBienesFinanciamientoSeleccionarPorIdProducto(lcMovNumero101, lcMovTipo101, lnIdProducto101)
        If rsTmp12.RecordCount > 0 Then
           lbPrimeraVez = False
            rsTmp12.MoveFirst
            Do While Not rsTmp12.EOF
                Select Case mo_ReglasFacturacion.FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(rsTmp12.Fields!idFuenteFinanciamiento, oConexion)
                Case 13       'Sis
                     lnSis = lnSis + rsTmp12.Fields!CantidadFinanciada
                Case 14      'Soat
                     lnSoat = lnSoat + rsTmp12.Fields!CantidadFinanciada
                Case 23      'Convenios
                     LnConvenio = LnConvenio + rsTmp12.Fields!CantidadFinanciada
                Case 0      'Exoneraciones
                     LnExonerac = LnExonerac + rsTmp12.Fields!CantidadFinanciada
                Case Else
                     LnOtrasS = LnOtrasS + rsTmp12.Fields!CantidadFinanciada
                End Select
                TotSalidas = TotSalidas + rsTmp12.Fields!CantidadFinanciada
                rsTmp12.MoveNext
            Loop
        End If
        rsTmp12.Close
        '
        Set rsTmp11 = Nothing
        Set rsTmp12 = Nothing
End Sub

Function ParteDiarioFechaVencimiento(lnIdAlmacen99 As Long, lcCodigo99 As String) As Date
        Dim rsTmp14 As New Recordset
        ParteDiarioFechaVencimiento = Date
        Set rsTmp14 = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen99, 0, lcCodigo99)
        If rsTmp14.RecordCount > 0 Then
           ParteDiarioFechaVencimiento = rsTmp14.Fields!fechaVencimiento
        End If
        rsTmp14.Close
        Set rsTmp14 = Nothing
End Function

Function ParteDiarioPrecio(lcCodigo100 As String, oConexion As Connection) As Double
        Dim rsTmp13 As New Recordset
        ParteDiarioPrecio = 0
        Set rsTmp13 = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigo100, 1, oConexion)
        If rsTmp13.RecordCount > 0 Then
           ParteDiarioPrecio = rsTmp13.Fields!PrecioUnitario
        End If
        rsTmp13.Close
        Set rsTmp13 = Nothing
End Function





'***Agrupa Saldos De Items Con Tipo Salida Diferente
'***Solo para Reporte ACUMULADOS
Sub CargaSaldosAgrupados(oRsReporteAgrupado As Recordset, oRsReporte As Recordset, oLnOrdenadoPor As Long)
               Dim lnIdAlmacen1 As Long, lnIdProducto As Long, lnPrecio As Double, lcCodigo  As String
               Dim lcNombre As String, lnCantidad As Long, lnStockMinimo As Long, lcLote As String
               Dim ldFechaVencimiento As Date, lnCantidadLote As Long
               With rsReporteAgrupado
                    .Fields.Append "idAlmacen", adInteger
                    .Fields.Append "idProducto", adInteger
                    .Fields.Append "precio", adDouble
                    .Fields.Append "codigo", adVarChar, 10, adFldIsNullable
                    .Fields.Append "nombre", adVarChar, 150, adFldIsNullable
                    .Fields.Append "cantidad", adInteger
                    .Fields.Append "stockMinimo", adInteger
                    .Fields.Append "lote", adVarChar, 20, adFldIsNullable
                    .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
                    .Fields.Append "cantidadLote", adInteger
                    .LockType = adLockOptimistic
                    .Open
               End With
               If oLnOrdenadoPor = 0 Then
                  'Por codigo
                  rsReporte.MoveFirst
                  Do While Not rsReporte.EOF
If Trim(rsReporte.Fields!codigo) = "00808" Then
lcCodigo = ""
End If
                     lnIdAlmacen1 = rsReporte.Fields!IdAlmacen
                     lnIdProducto = rsReporte.Fields!idProducto
                     lnPrecio = rsReporte.Fields!precio
                     lcCodigo = rsReporte.Fields!codigo
                     lcNombre = rsReporte.Fields!nombre
                     lnCantidad = 0
                     lnStockMinimo = IIf(IsNull(rsReporte.Fields!stockMinimo), 0, rsReporte.Fields!stockMinimo)
                     Do While Not rsReporte.EOF And lcCodigo = rsReporte.Fields!codigo
                        lnCantidad = lnCantidad + rsReporte.Fields!Cantidad
                        rsReporte.MoveNext
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                     Loop
                     rsReporteAgrupado.AddNew
                     rsReporteAgrupado.Fields!IdAlmacen = lnIdAlmacen1
                     rsReporteAgrupado.Fields!idProducto = lnIdProducto
                     rsReporteAgrupado.Fields!precio = lnPrecio
                     rsReporteAgrupado.Fields!codigo = lcCodigo
                     rsReporteAgrupado.Fields!nombre = lcNombre
                     rsReporteAgrupado.Fields!Cantidad = lnCantidad
                     rsReporteAgrupado.Fields!stockMinimo = lnStockMinimo
                     rsReporteAgrupado.Update
                  Loop
               Else
                  'Por Nombre
                  rsReporte.MoveFirst
                  Do While Not rsReporte.EOF
                     lnIdAlmacen1 = rsReporte.Fields!IdAlmacen
                     lnIdProducto = rsReporte.Fields!idProducto
                     lnPrecio = rsReporte.Fields!precio
                     lcCodigo = rsReporte.Fields!codigo
                     lcNombre = rsReporte.Fields!nombre
                     lnCantidad = 0
                     lnStockMinimo = IIf(IsNull(rsReporte.Fields!stockMinimo), 0, rsReporte.Fields!stockMinimo)
                     Do While Not rsReporte.EOF And lcNombre = rsReporte.Fields!nombre
                        lnCantidad = lnCantidad + rsReporte.Fields!Cantidad
                        rsReporte.MoveNext
                        If rsReporte.EOF Then
                           Exit Do
                        End If
                     Loop
                     rsReporteAgrupado.AddNew
                     rsReporteAgrupado.Fields!IdAlmacen = lnIdAlmacen1
                     rsReporteAgrupado.Fields!idProducto = lnIdProducto
                     rsReporteAgrupado.Fields!precio = lnPrecio
                     rsReporteAgrupado.Fields!codigo = lcCodigo
                     rsReporteAgrupado.Fields!nombre = lcNombre
                     rsReporteAgrupado.Fields!Cantidad = lnCantidad
                     rsReporteAgrupado.Fields!stockMinimo = lnStockMinimo
                     rsReporteAgrupado.Update
                  Loop
               End If
End Sub


Sub ParteDiarioSeparadandoMovimientosVtasYestrategicos()
        Dim lnIdProducto As Long, lnSaldoInicial As Long
        Dim lnRegistro As Long
        Dim lnRegTope As Long
        Dim rsTmp11 As New Recordset
        Dim rsTmp12 As New Recordset
        Dim rsTmp13 As New Recordset
        Dim rsTmp14 As New Recordset
        Dim rsTmp15 As New Recordset
        Dim lnidTipoConceptoFarmacia As Long
        Dim lnidTipoSalidaBienInsumo As Long
        '
        On Error GoTo ErrParteDia
        '
        oConexion.CursorLocation = adUseClient
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oFarmMovimientoDetalle.Conexion = oConexion
        
        'Proceso
        lcUltDiaMes = Trim(str(sighentidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))))
        ldFechaHistoricoXmes = CDate("01" & Format(mda_FechaInicio, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263)) - 1
        ldFechaHistoricoXmes = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaHistoricoXmes)
        ldFechaInicioMovim = DateAdd("n", 1, ldFechaHistoricoXmes)
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDIPorTproducto(ldFechaInicioMovim, mda_FechaFin, 0, "")
        lnTotalRegistros = rsReporte.RecordCount
        
        If lnTotalRegistros > 0 Then
           Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
            'GenerarRecordsetTemporalICI
            '
            lnRegistro = 1
            lnRegTope = 28320
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnIdProducto = rsReporte.Fields!idProducto
If Val(rsReporte.Fields!codigo) = 5335 Then
lcCodigo = ""
End If
                lcCodigo = rsReporte.Fields!codigo
                lcNombre = rsReporte.Fields!nombre
                lnidTipoSalidaBienInsumo = rsReporte.Fields!idTipoSalidaBienInsumo
                '*******Saldo Inicial****************************************
                lnSaldoInicial = 0
                'saldos-barre historico mensual
                For lnFor = 1 To Len(lc_AlmacenesParaICI)
                    If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                       lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                       lnFor = Len(lc_AlmacenesParaICI)
                    Else
                        lcTexto1 = ""
                        Do While True
                           If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                              Exit Do
                           Else
                              lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                              lnFor = lnFor + 1
                           End If
                        Loop
                        lnIdAlmacenRep = Val(lcTexto1)
                    End If
                    If lnIdAlmacenRep > 1 Then
                        Set rsErrores = mo_ReglasFarmacia.FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(lnIdProducto, lnIdAlmacenRep, ldFechaHistoricoXmes)
                        If mb_VtaYestrategicoSeparado = True Then
                           rsErrores.Filter = "IdTipoSalidaBienInsumo=" & lnidTipoSalidaBienInsumo
                        End If
                        Do While Not rsErrores.EOF
                            lnSaldoInicial = lnSaldoInicial + rsErrores.Fields!saldo
                            rsErrores.MoveNext
                        Loop
                        rsErrores.Close
                    End If
                Next
                'saldos-barre movimiento
                Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And lnidTipoSalidaBienInsumo = rsReporte.Fields!idTipoSalidaBienInsumo And rsReporte.Fields!fechaCreacion <= mda_FechaInicio
                   If rsReporte.Fields!MovTipo = "S" Then
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                        lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                      End If
                   Else
                      If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                         lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                      End If
                   End If
                   Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                   rsReporte.MoveNext
                   lnRegistro = lnRegistro + 1
If lnRegistro > lnRegTope Then
lcTexto10 = "="
End If
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                '****** Movimientos en el Rango de Fechas***********************************
                lnIngresos = 0: LnDevolucionesP = 0: TotIngresos = 0
                LnVentas = 0: lnSis = 0: lnSoat = 0: LnConvenio = 0: lnCreditoH = 0: lnDefensaN = 0
                LnOsDevol = 0: LnOsVencim = 0: LnOsMerma = 0: LnExonerac = 0: LnIntervencionS = 0
                LnOtrasS = 0: TotSalidas = 0
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And lnidTipoSalidaBienInsumo = rsReporte.Fields!idTipoSalidaBienInsumo And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       lbPrimeraVez = True
                       If rsReporte.Fields!MovTipo = "S" Then
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                                If mb_ConsiderarRecalculo = True Then
                                    '********* con recalculo
                                    lcTexto1 = rsReporte.Fields!MovTipo
                                    lcTexto2 = rsReporte.Fields!movNumero
                                    'Busca si tiene Pagos
                                    Set rsTmp11 = mo_ReglasFarmacia.FacturacionBienesPagoSeleccionarPorMovNumeroProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp11.RecordCount > 0 Then
                                       If rsTmp11.Fields!IdComprobantePago > 0 And rsTmp11.Fields!idEstadoFacturacion = 4 Then
                                          lbPrimeraVez = False
                                          rsTmp11.MoveFirst
                                          Do While Not rsTmp11.EOF
                                                Select Case rsReporte.Fields!idTipoConcepto
                                                Case 10       'Ventas
                                                     LnVentas = LnVentas + rsTmp11.Fields!CantidadPagar
                                                Case 13       'Sis
                                                     lnSis = lnSis + rsTmp11.Fields!CantidadPagar
                                                Case 14       'Soat
                                                     lnSoat = lnSoat + rsTmp11.Fields!CantidadPagar
                                                Case 23, 26      'Convenios, Credito Personal
                                                     LnConvenio = LnConvenio + rsTmp11.Fields!CantidadPagar
                                                Case 17       'Credito Hospitalario
                                                     lnCreditoH = lnCreditoH + rsTmp11.Fields!CantidadPagar
'                                                Case 22       'Defensa nacional
'                                                     lnDefensaN = lnDefensaN + rsTmp11.Fields!cantidadPagar
'                                                Case 7       'Otras salidas Devolucion
'                                                     LnOsDevol = LnOsDevol + rsTmp11.Fields!cantidadPagar
'                                                Case 5       'Otras salidas Vencimiento
'                                                     LnOsVencim = LnOsVencim + rsTmp11.Fields!cantidadPagar
'                                                Case 6       'Otras salidas Merma
'                                                     LnOsMerma = LnOsMerma + rsTmp11.Fields!cantidadPagar
'                                                Case 15       'Exoneraciones
'                                                     LnExonerac = LnExonerac + rsTmp11.Fields!cantidadPagar
                                                Case 16       'Intervencion Sanitaria
                                                     LnIntervencionS = LnIntervencionS + rsTmp11.Fields!CantidadPagar
                                                Case Else
                                                     LnOtrasS = LnOtrasS + rsTmp11.Fields!CantidadPagar
                                                End Select
                                                TotSalidas = TotSalidas + rsTmp11.Fields!CantidadPagar
                                                rsTmp11.MoveNext
                                          Loop
                                       End If
                                    End If
                                    rsTmp11.Close
                                    'Busca si tiene algun seguro o exoneracion (Plan)
                                    Set rsTmp12 = mo_ReglasFarmacia.FacturacionBienesFinancSeleccionarPorProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                    If rsTmp12.RecordCount > 0 Then
                                       lbPrimeraVez = False
                                        rsTmp12.MoveFirst
                                        Do While Not rsTmp12.EOF
                                            Select Case mo_ReglasFacturacion.FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia(oConexion, rsTmp12.Fields!idFuenteFinanciamiento)
                                            Case 10       'Ventas
                                                 LnVentas = LnVentas + rsTmp12.Fields!CantidadFinanciada
                                            Case 13       'Sis
                                                 lnSis = lnSis + rsTmp12.Fields!CantidadFinanciada
                                            Case 14      'Soat
                                                 lnSoat = lnSoat + rsTmp12.Fields!CantidadFinanciada
                                            Case 16       'Intervencion Sanitaria
                                                 LnIntervencionS = LnIntervencionS + rsTmp12.Fields!CantidadFinanciada
                                            Case 17       'Credito Hospitalario
                                                 lnCreditoH = lnCreditoH + rsTmp12.Fields!CantidadFinanciada
                                            Case 23, 26      'Convenios, Credito Personal
                                                 LnConvenio = LnConvenio + rsTmp12.Fields!CantidadFinanciada
'                                            Case 0      'Exoneraciones
'                                                 If rsTmp12.Fields!cantidadFinanciada > 0 Then
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                 Else
'                                                    '**no se sabe la CANTIDAD EXONERADA solo el IMPORTE EXONERADO
'                                                    LnExonerac = LnExonerac + rsTmp12.Fields!cantidadFinanciada
'                                                    LnExonerac = LnExonerac - rsTmp12.Fields!cantidadFinanciada
'                                                    TotSalidas = TotSalidas - rsTmp12.Fields!cantidadFinanciada
'                                                 End If
                                            Case Else
                                                 LnOtrasS = LnOtrasS + rsTmp12.Fields!CantidadFinanciada
                                            End Select
                                            TotSalidas = TotSalidas + rsTmp12.Fields!CantidadFinanciada
                                            rsTmp12.MoveNext
                                        Loop
                                    End If
                                    rsTmp12.Close
                                    '
                                    If lbPrimeraVez = False Then
                                        Do While Not rsReporte.EOF And lcTexto1 = rsReporte.Fields!MovTipo And lcTexto2 = rsReporte.Fields!movNumero And lnIdProducto = rsReporte.Fields!idProducto And lnidTipoSalidaBienInsumo = rsReporte.Fields!idTipoSalidaBienInsumo
                                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                                           rsReporte.MoveNext
                                           lnRegistro = lnRegistro + 1

                                           If rsReporte.EOF Then
                                              Exit Do
                                           End If
                                        Loop
                                    End If
                                End If
                                If lbPrimeraVez = True Then
                                    '******** sin recalculo
                                    Select Case rsReporte.Fields!idTipoConcepto
                                    Case 10       'Ventas
                                         LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                    Case 13       'Sis
                                         lnSis = lnSis + rsReporte.Fields!Cantidad
                                    Case 14       'Soat
                                         lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                    Case 23, 26      'Convenios, Credito Personal
                                         LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                    Case 17       'Credito Hospitalario
                                         lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
'                                    Case 22       'Defensa nacional
'                                         lnDefensaN = lnDefensaN + rsReporte.Fields!cantidad
'                                    Case 7       'Otras salidas Devolucion
'                                         LnOsDevol = LnOsDevol + rsReporte.Fields!cantidad
'                                    Case 5       'Otras salidas Vencimiento
'                                         LnOsVencim = LnOsVencim + rsReporte.Fields!cantidad
'                                    Case 6       'Otras salidas Merma
'                                         LnOsMerma = LnOsMerma + rsReporte.Fields!cantidad
'                                    Case 15       'Exoneraciones
'                                         LnExonerac = LnExonerac + rsReporte.Fields!cantidad
                                    Case 16       'Intervencion Sanitaria
                                         LnIntervencionS = LnIntervencionS + rsReporte.Fields!Cantidad
                                    Case Else
                                         LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                    End Select
                                    If rsReporte.Fields!IdAlmacenDestino = 10 And rsReporte.Fields!idTipoConcepto = 4 And mb_ConsideraOSH = False Then
                                       'destino=otros servicios hospital, tipoConcepto=distribucion
                                    Else
                                        TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                    End If
                                End If
                          End If
                       Else
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                                Select Case rsReporte.Fields!idTipoConcepto
                                Case 19        'Inventario
                                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                Case 21        'Devolucion de Pacientes
                                     LnDevolucionesP = LnDevolucionesP + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                Case Else        'Ingresos
                                     lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                                     TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                End Select
                                
                           End If
                       End If
                       If lbPrimeraVez = True Then
                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                           rsReporte.MoveNext
                           lnRegistro = lnRegistro + 1

                       End If
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And lnidTipoSalidaBienInsumo = rsReporte.Fields!idTipoSalidaBienInsumo And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                       Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                       rsReporte.MoveNext
                       lnRegistro = lnRegistro + 1
If lnRegistro > lnRegTope Then
lcTexto10 = "="
End If
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                '
                lnPrecio = 0
                Set rsTmp13 = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                If rsTmp13.RecordCount > 0 Then
                   lnPrecio = rsTmp13.Fields!PrecioUnitario
                End If
                rsTmp13.Close
                '
                If InStr(lc_AlmacenesParaICI, "/") > 0 Then
                   lnIdAlmacen = Val(Left(lc_AlmacenesParaICI, InStr(lc_AlmacenesParaICI, "/") - 1)) 'toma la primera farmacia para las FECHAS DE VENCIMIENTO
                Else
                   lnIdAlmacen = Val(lc_AlmacenesParaICI)
                End If
                'ParteDiarioFechaVencimiento lnIdAlmacen, lcCodigo
                ldFechaVencimiento = Date
                Set rsTmp14 = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacen, 0, lcCodigo)
                If rsTmp14.RecordCount > 0 Then
                   ldFechaVencimiento = rsTmp14.Fields!fechaVencimiento
                End If
                rsTmp14.Close
                '
                lbContinua = True
                If mb_ConsiderarSinMovimientos = False Then
                   If TotIngresos = 0 And TotSalidas = 0 Then
                      lbContinua = False
                   End If
                End If
                '
                If lbContinua Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!codigo = lcCodigo
                    mrs_Tmp.Fields!nombre = Trim(lcNombre) & "  (" & sighentidades.ElijeSiEsEstrategicoDevuelveNombre(lnidTipoSalidaBienInsumo) & ")"
                    mrs_Tmp.Fields!precio = lnPrecio
                    mrs_Tmp.Fields!saldoI = lnSaldoInicial
                    mrs_Tmp.Fields!ingresos = lnIngresos
                    mrs_Tmp.Fields!DevolucionesP = LnDevolucionesP
                    mrs_Tmp.Fields!TotIngresos = TotIngresos
                    mrs_Tmp.Fields!Ventas = LnVentas
                    mrs_Tmp.Fields!sis = lnSis
                    mrs_Tmp.Fields!soat = lnSoat
                    mrs_Tmp.Fields!convenio = LnConvenio
                    mrs_Tmp.Fields!creditoH = lnCreditoH
                    mrs_Tmp.Fields!defensaN = lnDefensaN
                    mrs_Tmp.Fields!OsDevol = LnOsDevol
                    mrs_Tmp.Fields!OsVencim = LnOsVencim
                    mrs_Tmp.Fields!OsMerma = LnOsMerma
                    mrs_Tmp.Fields!Exonerac = LnExonerac
                    mrs_Tmp.Fields!IntervencionS = LnIntervencionS
                    mrs_Tmp.Fields!otrasS = LnOtrasS
                    mrs_Tmp.Fields!TotSalidas = TotSalidas
                    mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                    mrs_Tmp.Update
                End If
                'Graba Datos en Temporal
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
       End If
       'oConexion.Close
        Set rsTmp11 = Nothing
        Set rsTmp12 = Nothing
        Set rsTmp13 = Nothing
        Set rsTmp14 = Nothing
        Set rsTmp15 = Nothing
        Exit Sub
ErrParteDia:
     MsgBox Err.Description
     Resume
End Sub






Sub ProcesaDatosIDI(oConexion As Connection)
                On Error GoTo ErrIDI
                Dim Lnab As Long: Dim lnReingresos As Long: Dim LnDistribucion As Long
                Dim LnTransferencia As Long: Dim LnDevolVencido As Long
                Dim oConexionFox As New Connection, lnSaldoInicial As Long
                Dim oRsFox As New Recordset, oRsFox1 As New Recordset, oRsFox2 As New Recordset
                Dim rsTmp9 As New Recordset
                Dim lcDisa As String, lcEstablecimiento As String, lcAnioMes As String
                Dim lnPrecioItem As Double, ldFechaHistoricoXmes As Date
                Dim lnDonacionesOtrIng As Long, lnDonacionesIng As Long, lnDonacionesSaldoI As Long
                Dim lnDonacionesSal As Long, lnDonacionesOtrSal As Long, ldDonacionFechaVctoUlt As Date
                Dim lcLoteXitem As String, lnSaldoFinalD As Long, lnSaldofinal As Long
                
                '
                oConexionFox.CommandTimeout = 300
                oConexionFox.Open "DSN=his"
                oConexionFox.CursorLocation = adUseClient
                
                ldFechaHistoricoXmes = CDate("01" & Format(mda_FechaInicio, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
                ldFechaHistoricoXmes = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaHistoricoXmes)
                
                lcDisa = Right("000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(239)), 3)
                lcEstablecimiento = lcBuscaParametro.SeleccionaFilaParametro(208)
                lcAnioMes = Format(mda_FechaInicio, "yyyy") & Format(mda_FechaInicio, "mm")
                lcCodigo = Format(mda_FechaFin, "yyyy") & Format(mda_FechaFin, "mm")
                mo_ReglasFarmacia.PreparaTablasDBF oRsFox, oRsFox1, oRsFox2, lc_CodigoSismed, lcAnioMes, oConexionFox, _
                                                   lcCodigo, True
                '
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
If Val(rsReporte.Fields!codigo) = 91 Then
lcCodigo = ""
End If
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    '*******Saldo Inicial********
                    lnSaldoInicial = 0
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion < mda_FechaInicio
                       If rsReporte.Fields!MovTipo = "S" Then
                          If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                            lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                          End If
                       Else
                          If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                             lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                          End If
                       End If
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    lnDonacionesSaldoI = 0
                    If mb_EsDonaciones = True Then
                       lnDonacionesSaldoI = lnSaldoInicial
                    End If

                    '****** Movimientos en el Rango de Fechas***********
                    lnDonacionesOtrIng = 0: lnDonacionesIng = 0: lnDonacionesSal = 0: lnDonacionesOtrSal = 0
                    lnIngresos = 0: LnDevolucionesP = 0: Lnab = 0: lnReingresos = 0: TotIngresos = 0
                    LnDistribucion = 0: LnTransferencia = 0: LnDevolVencido = 0: LnDevolMerma = 0: LnVentaInst = 0: LnExoner = 0: LnOtrasS = 0: TotSalidas = 0
                    If Not rsReporte.EOF Then
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                       If rsReporte.Fields!MovTipo = "S" Then
                          If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                                Select Case rsReporte.Fields!idTipoConcepto
                                Case 3           'Salidas x Donaciones
                                     lnDonacionesSal = lnDonacionesSal + rsReporte.Fields!Cantidad
                                Case 4       'Distribucion
                                     LnDistribucion = LnDistribucion + rsReporte.Fields!Cantidad
                                Case 8, 9       'Transferencias
                                     LnTransferencia = LnTransferencia + rsReporte.Fields!Cantidad
                                Case 5        '
                                     LnDevolVencido = LnDevolVencido + rsReporte.Fields!Cantidad
                                Case 6        'merma (deterioro)
                                     LnDevolMerma = LnDevolMerma + rsReporte.Fields!Cantidad
                                Case 12        'venta institucional
                                     LnVentaInst = LnVentaInst + rsReporte.Fields!Cantidad
                                Case 15        '
                                     LnExoner = LnExoner + rsReporte.Fields!Cantidad
                                Case Else      'Ajuste inventario, Defensa Nacional
                                     If mb_EsDonaciones = True Then
                                        lnDonacionesOtrSal = lnDonacionesOtrSal + rsTmp.Fields!CantidadPagar
                                     Else
                                        LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                     End If
                                End Select
                                TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                          End If
                       Else
                          If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                                Select Case rsReporte.Fields!idTipoConcepto
                                Case 1, 2       'Compra,Encargo
                                     lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                                     Lnab = Lnab + rsReporte.Fields!Cantidad
                                Case 3           'Ingresos x Donaciones
                                     lnDonacionesIng = lnDonacionesIng + rsReporte.Fields!Cantidad
                                Case 19        'Inventario
                                     If mb_EsDonaciones = True Then
                                        lnDonacionesOtrIng = lnDonacionesOtrIng + rsReporte.Fields!Cantidad
                                     Else
                                        lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                     End If
                                Case 21        'Devolucion de Pacientes
                                     LnDevolucionesP = LnDevolucionesP + rsReporte.Fields!Cantidad
                                     Lnab = Lnab + rsReporte.Fields!Cantidad
                                Case Else        'Devol.por Venc/deter/sobrestock/Ajuste Inventario
                                     If mb_EsDonaciones = True Then
                                        lnDonacionesOtrIng = lnDonacionesOtrIng + rsReporte.Fields!Cantidad
                                     Else
                                        lnReingresos = lnReingresos + rsReporte.Fields!Cantidad
                                     End If
                                End Select
                                TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                           End If
                       End If
                       Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    End If
                    '
                    lnPrecio = 0
                    If mb_EsDonaciones = True Then
                        TotSalidas = 0: TotIngresos = 0: lnSaldoInicial = 0
                        Set rsTmp = mo_reglasComunes.CatalogoBienesInsumosFiltrarDonacionesXcodigo(lcCodigo)
                    Else
                        Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigo, 1, oConexion)
                    End If
                    If rsTmp.RecordCount > 0 Then
                       lnPrecio = rsTmp.Fields!PrecioUnitario
                    End If
                    rsTmp.Close
                    '
                    ldFechaVencimiento = Date
                    lcLoteXitem = ""
                    Set rsTmp = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen, 0, lcCodigo, oConexion)
                    If rsTmp.RecordCount > 0 Then
                       ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                       lcLoteXitem = rsTmp.Fields!Lote
                    End If
                    rsTmp.Close
                    ldDonacionFechaVctoUlt = Date
                    
                    If mb_EsDonaciones = True Then
                       ldDonacionFechaVctoUlt = ldFechaVencimiento
                    End If
                    '
                    If (lnDonacionesSaldoI + TotIngresos) > 0 Or (lnDonacionesSaldoI + lnDonacionesOtrIng + lnDonacionesIng) > 0 Or TotSalidas > 0 Then
                         mrs_Tmp.AddNew
                         mrs_Tmp.Fields!codigo = lcCodigo
                         mrs_Tmp.Fields!nombre = lcNombre
                         mrs_Tmp.Fields!precio = lnPrecio
                         If mb_EsDonaciones = True Then
                             mrs_Tmp.Fields!saldoI = lnDonacionesSaldoI
                             mrs_Tmp.Fields!ingresos = lnDonacionesOtrIng + lnDonacionesIng
                             mrs_Tmp.Fields!DevolucionesP = lnDonacionesSal
                             mrs_Tmp.Fields!otrasS = lnDonacionesOtrSal
                             mrs_Tmp.Fields!TotSalidas = lnDonacionesSal + lnDonacionesOtrSal
                             mrs_Tmp.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                         Else
                             mrs_Tmp.Fields!saldoI = lnSaldoInicial
                             mrs_Tmp.Fields!ingresos = lnIngresos
                             mrs_Tmp.Fields!DevolucionesP = LnDevolucionesP
                             mrs_Tmp.Fields!ab = Lnab
                             mrs_Tmp.Fields!reingresos = lnReingresos
                             mrs_Tmp.Fields!TotIngresos = TotIngresos
                             mrs_Tmp.Fields!Distribucion = LnDistribucion
                             mrs_Tmp.Fields!Transferencia = LnTransferencia
                             mrs_Tmp.Fields!DevolVencido = LnDevolVencido
                             mrs_Tmp.Fields!DevolMerma = LnDevolMerma
                             mrs_Tmp.Fields!ventaInst = LnVentaInst
                             mrs_Tmp.Fields!Exoner = LnExoner
                             mrs_Tmp.Fields!otrasS = LnOtrasS
                             mrs_Tmp.Fields!TotSalidas = TotSalidas
                             mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                         End If
                         mrs_Tmp.Update
                         '***************************grabar el IDI-detalle**************************************
                         lnSaldofinal = lnSaldoInicial + lnIngresos + lnReingresos + LnDevolucionesP - TotSalidas
                         lnSaldoFinalD = (lnDonacionesSaldoI + lnDonacionesOtrIng + lnDonacionesIng) - (lnDonacionesSal + lnDonacionesOtrSal)
                         lnPrecioItem = lnPrecio
                         oRsFox.AddNew
                         oRsFox.Fields!CODIGO_EJE = lcDisa
                         oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                         oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                         oRsFox.Fields!annoMes = lcAnioMes
                         oRsFox.Fields!codigo_med = Left(lcCodigo, 7)
                         oRsFox.Fields!saldo = lnSaldoInicial
                         oRsFox.Fields!precio = lnPrecioItem
                         oRsFox.Fields!ingre = lnIngresos
                         oRsFox.Fields!reingre = lnReingresos + LnDevolucionesP
                         oRsFox.Fields!VENTA = 0
                         oRsFox.Fields!sis = 0
                         oRsFox.Fields!intersan = 0
                         oRsFox.Fields!fac_perd = 0                        'falta
                         oRsFox.Fields!DEFNAC = 0
                         oRsFox.Fields!exo = LnExoner
                         oRsFox.Fields!soat = 0
                         oRsFox.Fields!credHosp = 0
                         oRsFox.Fields!otr_conv = 0
                         oRsFox.Fields!DEVOL = 0
                         oRsFox.Fields!vencido = 0
                         oRsFox.Fields!merma = 0
                         oRsFox.Fields!distri = LnDistribucion
                         oRsFox.Fields!transf = LnTransferencia
                         oRsFox.Fields!ventaInst = LnVentaInst
                         oRsFox.Fields!DEV_VEN = LnDevolVencido
                         oRsFox.Fields!DEV_MERMA = LnDevolMerma
                         oRsFox.Fields!otras_sal = LnOtrasS
                         oRsFox.Fields!STOCK_FIN = lnSaldofinal
                         oRsFox.Fields!stock_fin1 = lnSaldofinal
                         oRsFox.Fields!REQ = 0
                         oRsFox.Fields!Total = TotSalidas
                         If mb_EsDonaciones = False Then
                            oRsFox.Fields!FEC_EXP = ldFechaVencimiento
                         Else
                            oRsFox.Fields!saldo = lnDonacionesSaldoI
                            oRsFox.Fields!ingre = lnDonacionesOtrIng + lnDonacionesIng
                            oRsFox.Fields!distri = lnDonacionesSal
                            oRsFox.Fields!otras_sal = lnDonacionesOtrSal
                            oRsFox.Fields!Total = lnDonacionesSal + lnDonacionesOtrSal
                            oRsFox.Fields!STOCK_FIN = lnSaldoFinalD
                            oRsFox.Fields!stock_fin1 = lnSaldoFinalD
                            oRsFox.Fields!FEC_EXP = ldDonacionFechaVctoUlt
                         End If
                         oRsFox.Fields!do_saldo = 0
                         oRsFox.Fields!do_ingre = 0
                         oRsFox.Fields!do_con = 0
                         oRsFox.Fields!do_otr = 0
                         oRsFox.Fields!do_tot = 0
                         oRsFox.Fields!do_stk = 0
'                         oRsFox.Fields!do_saldo = lnDonacionesSaldoI
'                         oRsFox.Fields!do_ingre = lnDonacionesOtrIng + lnDonacionesIng
'                         oRsFox.Fields!do_con = lnDonacionesSal
'                         oRsFox.Fields!do_otr = lnDonacionesOtrSal
'                         oRsFox.Fields!do_tot = lnDonacionesSal + lnDonacionesOtrSal
'                         oRsFox.Fields!do_stk = lnSaldoFinalD
'                         If mb_EsDonaciones = True Then
'                            oRsFox.Fields!do_fecExp = ldDonacionFechaVctoUlt
'                         End If
                        
                        ' oRsFox.Fields!fecha = Date
                         oRsFox.Fields!Usuario = " "
                         oRsFox.Fields!indiProc = " "
                         oRsFox.Fields!SIT = "1"
                         oRsFox.Fields!indiSiga = " "
                         oRsFox.Fields!dstkCero = 0
                         oRsFox.Fields!mptoRepo = 0
                         oRsFox.Update
                         'FormDetL
                         oRsFox1.AddNew
                         oRsFox1.Fields!CODIGO_EJE = lcDisa
                         oRsFox1.Fields!CODIGO_PRE = lc_CodigoSismed
                         oRsFox1.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                         oRsFox1.Fields!annoMes = lcAnioMes
                         oRsFox1.Fields!codigo_med = Left(lcCodigo, 7)
                         oRsFox1.Fields!Lote = lcLoteXitem
                         oRsFox1.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                         oRsFox1.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                         oRsFox1.Fields!SIT = "1"
                         oRsFox1.Update
                         'FormDetM
                         oRsFox2.AddNew
                         oRsFox2.Fields!CODIGO_EJE = lcDisa
                         oRsFox2.Fields!CODIGO_PRE = lc_CodigoSismed
                         oRsFox2.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                         oRsFox2.Fields!annoMes = lcAnioMes
                         oRsFox2.Fields!codigo_med = Left(lcCodigo, 7)
                         oRsFox2.Fields!Lote = lcLoteXitem
                         oRsFox2.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                         oRsFox2.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                         oRsFox2.Fields!SIT = "1"
                         oRsFox2.Update
                         '***************************grabar el IDI-detalle**************************************
                    End If
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                '**************Carga Saldos de Items sin Movimientos****************
                If mb_ConsiderarSinMovimientos = True Then
                   Set rsTmp9 = mo_ReglasFarmacia.farmSaldoMensualFiltrarFechaYalmacen(ldFechaHistoricoXmes, lnIdAlmacen)
                   If rsTmp9.RecordCount > 0 Then
                      rsTmp9.MoveFirst
                      Do While Not rsTmp9.EOF
                         lnIdAlmacen = rsTmp9.Fields!IdAlmacen
                         lcCodigo = rsTmp9.Fields!codigo
                         lcNombre = rsTmp9.Fields!nombre
                         lnSaldoInicial = 0
                         Do While Not rsTmp9.EOF And lcCodigo = rsTmp9.Fields!codigo
                            lnSaldoInicial = lnSaldoInicial + rsTmp9.Fields!saldo
                            rsTmp9.MoveNext
                            If rsTmp9.EOF Then
                               Exit Do
                            End If
                         Loop
                         If mrs_Tmp.RecordCount > 0 Then
                            mrs_Tmp.MoveFirst
                            mrs_Tmp.Find "codigo='" & lcCodigo & "'"
                            If mrs_Tmp.EOF Then
                                lnSaldofinal = lnSaldoInicial
                                lnDonacionesSaldoI = lnSaldoInicial
                                lnSaldoFinalD = lnSaldoInicial
                                '
                                lnPrecio = 0
                                If rsTmp.State = 1 Then rsTmp.Close
                                If mb_EsDonaciones = True Then
                                    Set rsTmp = mo_reglasComunes.CatalogoBienesInsumosFiltrarDonacionesXcodigo(lcCodigo)
                                Else
                                    Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                                End If
                                If rsTmp.RecordCount > 0 Then
                                   lnPrecio = rsTmp.Fields!PrecioUnitario
                                End If
                                rsTmp.Close
                                '
                                ldFechaVencimiento = Date
                                lcLoteXitem = ""
                                Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacen, 0, lcCodigo)
                                If rsTmp.RecordCount > 0 Then
                                   ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                                   lcLoteXitem = rsTmp.Fields!Lote
                                End If
                                rsTmp.Close
                                ldDonacionFechaVctoUlt = Date
                                If mb_EsDonaciones = True Then
                                   ldDonacionFechaVctoUlt = ldFechaVencimiento
                                End If
                                '
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!codigo = lcCodigo
                                mrs_Tmp.Fields!nombre = lcNombre
                                mrs_Tmp.Fields!precio = lnPrecio
                                If mb_EsDonaciones = True Then
                                    mrs_Tmp.Fields!saldoI = lnDonacionesSaldoI
                                    mrs_Tmp.Fields!ingresos = 0
                                    mrs_Tmp.Fields!DevolucionesP = 0
                                    mrs_Tmp.Fields!otrasS = 0
                                    mrs_Tmp.Fields!TotSalidas = 0
                                    mrs_Tmp.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                                Else
                                    mrs_Tmp.Fields!saldoI = lnSaldoInicial
                                    mrs_Tmp.Fields!ingresos = 0
                                    mrs_Tmp.Fields!DevolucionesP = 0
                                    mrs_Tmp.Fields!ab = 0
                                    mrs_Tmp.Fields!reingresos = 0
                                    mrs_Tmp.Fields!TotIngresos = 0
                                    mrs_Tmp.Fields!Distribucion = 0
                                    mrs_Tmp.Fields!Transferencia = 0
                                    mrs_Tmp.Fields!DevolVencido = 0
                                    mrs_Tmp.Fields!DevolMerma = 0
                                    mrs_Tmp.Fields!ventaInst = 0
                                    mrs_Tmp.Fields!Exoner = 0
                                    mrs_Tmp.Fields!otrasS = 0
                                    mrs_Tmp.Fields!TotSalidas = 0
                                    mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                                End If
                                mrs_Tmp.Update
                                '***************************grabar el ICI-detalle**************************************
                                lnSaldofinal = lnSaldoInicial
                                lnSaldoFinalD = lnDonacionesSaldoI
                                lnPrecioItem = lnPrecio
                                oRsFox.AddNew
                                oRsFox.Fields!CODIGO_EJE = lcDisa
                                oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox.Fields!annoMes = lcAnioMes
                                oRsFox.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox.Fields!saldo = lnSaldoInicial
                                oRsFox.Fields!precio = lnPrecioItem
                                oRsFox.Fields!ingre = 0
                                oRsFox.Fields!reingre = 0
                                oRsFox.Fields!VENTA = 0
                                oRsFox.Fields!sis = 0
                                oRsFox.Fields!intersan = 0
                                oRsFox.Fields!fac_perd = 0                        'falta
                                oRsFox.Fields!DEFNAC = 0
                                oRsFox.Fields!exo = 0
                                oRsFox.Fields!soat = 0
                                oRsFox.Fields!credHosp = 0
                                oRsFox.Fields!otr_conv = 0
                                oRsFox.Fields!DEVOL = 0
                                oRsFox.Fields!vencido = 0
                                oRsFox.Fields!merma = 0
                                oRsFox.Fields!distri = 0
                                oRsFox.Fields!transf = 0
                                oRsFox.Fields!ventaInst = 0
                                oRsFox.Fields!DEV_VEN = 0
                                oRsFox.Fields!DEV_MERMA = 0
                                oRsFox.Fields!otras_sal = 0
                                oRsFox.Fields!STOCK_FIN = lnSaldofinal
                                oRsFox.Fields!stock_fin1 = lnSaldofinal
                                oRsFox.Fields!REQ = 0
                                oRsFox.Fields!Total = 0
                                If mb_EsDonaciones = False Then
                                   oRsFox.Fields!FEC_EXP = ldFechaVencimiento
                                Else
                                   oRsFox.Fields!saldo = lnDonacionesSaldoI
                                   oRsFox.Fields!ingre = 0
                                   oRsFox.Fields!distri = 0
                                   oRsFox.Fields!otras_sal = 0
                                   oRsFox.Fields!Total = 0
                                   oRsFox.Fields!STOCK_FIN = lnSaldoFinalD
                                   oRsFox.Fields!stock_fin1 = lnSaldoFinalD
                                   oRsFox.Fields!FEC_EXP = ldDonacionFechaVctoUlt
                                End If
                                oRsFox.Fields!do_saldo = 0
                                oRsFox.Fields!do_ingre = 0
                                oRsFox.Fields!do_con = 0
                                oRsFox.Fields!do_otr = 0
                                oRsFox.Fields!do_tot = 0
                                oRsFox.Fields!do_stk = 0
        '                         oRsFox.Fields!do_saldo = lnDonacionesSaldoI
        '                         oRsFox.Fields!do_ingre = lnDonacionesOtrIng + lnDonacionesIng
        '                         oRsFox.Fields!do_con = lnDonacionesSal
        '                         oRsFox.Fields!do_otr = lnDonacionesOtrSal
        '                         oRsFox.Fields!do_tot = lnDonacionesSal + lnDonacionesOtrSal
        '                         oRsFox.Fields!do_stk = lnSaldoFinalD
        '                         If mb_EsDonaciones = True Then
        '                            oRsFox.Fields!do_fecExp = ldDonacionFechaVctoUlt
        '                         End If
                                ' oRsFox.Fields!fecha = Date
                                oRsFox.Fields!Usuario = " "
                                oRsFox.Fields!indiProc = " "
                                oRsFox.Fields!SIT = "1"
                                oRsFox.Fields!indiSiga = " "
                                oRsFox.Fields!dstkCero = 0
                                oRsFox.Fields!mptoRepo = 0
                                oRsFox.Update
                                'FormDetL
                                oRsFox1.AddNew
                                oRsFox1.Fields!CODIGO_EJE = lcDisa
                                oRsFox1.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox1.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox1.Fields!annoMes = lcAnioMes
                                oRsFox1.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox1.Fields!Lote = lcLoteXitem
                                oRsFox1.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                                oRsFox1.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                                oRsFox1.Fields!SIT = "1"
                                oRsFox1.Update
                                'FormDetM
                                oRsFox2.AddNew
                                oRsFox2.Fields!CODIGO_EJE = lcDisa
                                oRsFox2.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox2.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox2.Fields!annoMes = lcAnioMes
                                oRsFox2.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox2.Fields!Lote = lcLoteXitem
                                oRsFox2.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                                oRsFox2.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                                oRsFox2.Fields!SIT = "1"
                                oRsFox2.Update
                            End If
                         End If
                      Loop
                   End If
                   rsTmp9.Close
                End If
                
                
                
                
                
                
                
                '***************************grabar el IDI- cabecera*************************************
                mo_ReglasFarmacia.dbfFormatoSeleccionarTodos oRsFox, oConexionFox
                oRsFox.AddNew
                oRsFox.Fields!CODIGO_EJE = lcDisa
                oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                oRsFox.Fields!annoMes = lcAnioMes
                oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                oRsFox.Fields!tipo_pre = "A"
                oRsFox.Fields!rec_vtas = 0
                oRsFox.Fields!rec_sis = 0
                oRsFox.Fields!rec_ints = 0
                oRsFox.Fields!rec_dn = 0
                oRsFox.Fields!rec_Exo = 0
                oRsFox.Fields!rec_soat = 0
                oRsFox.Fields!rec_crehos = 0
                oRsFox.Fields!rec_otrcon = 0
                oRsFox.Fields!indiProc = "A"
                oRsFox.Fields!fecha = Date
                'oRsFox.Fields!fechaUlt = Date
                oRsFox.Fields!vers = "V2.0 04102011"
                oRsFox.Fields!SIT = "1"
                oRsFox.Fields!fdesde = Format(mda_FechaInicio, "dd/mm/yyyy")
                oRsFox.Fields!fhasta = Format(mda_FechaFin, "dd/mm/yyyy")
                oRsFox.Fields!ctrlcal = "P"
                'oRsFox.Fields!catalogo = Date               'vacio
                oRsFox.Fields!codpto = lcEstablecimiento
                oRsFox.Fields!tip_ins = "E"
                oRsFox.Update
                '***************************grabar el IDI- cabecera*************************************
                If lnOrdenadoPor = 0 Then
                   mrs_Tmp.Sort = "codigo"
                Else
                   mrs_Tmp.Sort = "nombre"
                End If
                Exit Sub
ErrIDI:
    MsgBox Err.Description
    Resume
End Sub



'debb-14/09/2015
Sub ProcesarDatosICI()

            
            On Error GoTo ErrICI
            Dim oConexionFox As New Connection, lnSaldoInicial As Long
            Dim oRsFox As New Recordset, oRsFox1 As New Recordset, oRsFox2 As New Recordset
            Dim rsTmp9 As New Recordset, RsTmp989 As New Recordset, RsTmp988 As New Recordset
            Dim lcDisa As String, lcEstablecimiento As String, lcAnioMes As String
            Dim lnPrecioItem As Double, lnCantidadDevolucion As Long, lbNoEntro As Boolean, lbEsPrimerLoteDelItem As Boolean
            Dim lnDonacionesOtrIng As Long, lnDonacionesIng As Long, lnDonacionesSaldoI As Long
            Dim lnDonacionesSal As Long, lnDonacionesOtrSal As Long, ldDonacionFechaVctoUlt As Date
            Dim lcLoteXitem As String, lnSaldoFinalD As Long, lnSaldofinal As Long, lnTipoConceptoFarm As Long
            Dim lcTexto2Fmovimiento As Date, lcTipoMI As String
            'debb-10/12/2018
           ' lb_SeGrabaICImensual = sighentidades.VerificaSiRangoEsDeUnMesCompleto(mda_FechaInicio, mda_FechaFin, lc_CodigoItem)
              
            '
            With rsDebug
                .Fields.Append "codigo", adVarChar, 7, adFldIsNullable
                .Fields.Append "movtipo", adVarChar, 1, adFldIsNullable
                .Fields.Append "movnumero", adVarChar, 9, adFldIsNullable
                .Fields.Append "TipoConcepto", adInteger
                .Fields.Append "cantidad", adInteger
                .LockType = adLockOptimistic
                .Open
                
            End With
            '
            lcCodigo = "": lcTexto2 = "": lcTexto2Fmovimiento = Now
            '
            lcDisa = Right("000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(239)), 3)
            lcEstablecimiento = lcBuscaParametro.SeleccionaFilaParametro(208)
            '
            oConexionFox.CommandTimeout = 300
            oConexionFox.Open "DSN=his"
            oConexionFox.CursorLocation = adUseClient
            '
            lcAnioMes = Format(mda_FechaInicio, "yyyy") & Format(mda_FechaInicio, "mm")
            lcCodigo = Format(mda_FechaFin, "yyyy") & Format(mda_FechaFin, "mm")
            mo_ReglasFarmacia.PreparaTablasDBF oRsFox, oRsFox1, oRsFox2, lc_CodigoSismed, lcAnioMes, oConexionFox, lcCodigo, True
            '
            oConexion.Open sighentidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            '
            Set oFarmMovimientoDetalle.Conexion = oConexion
            'Proceso
            lcUltDiaMes = Trim(str(sighentidades.DevuelveUltimoDiaDelMes(Month(mda_FechaInicio), Year(mda_FechaInicio))))
            ldFechaHistoricoXmes = CDate("01" & Format(mda_FechaInicio, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
            ldFechaHistoricoXmes = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaHistoricoXmes)
            ldFechaInicioMovim = DateAdd("n", 1, ldFechaHistoricoXmes)
            Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosParaICIeIDIPorTproducto(ldFechaInicioMovim, mda_FechaFin, 0, "")
            'Set rsReporte = oBuscaMovimientos.farmDevuelveMovimientosParaICI(ldFechaInicioMovim, mda_FechaFin, 0, "")
            If Val(lc_CodigoItem) > 0 Then
               rsReporte.Filter = " codigo='" & lc_CodigoItem & "'"
            End If
            lnTotalRegistros = rsReporte.RecordCount
GrabaParametro206 "procesa ICI totalRegistros"
            If lnTotalRegistros > 0 Then
                Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                GenerarRecordsetTemporalICI
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
If Val(rsReporte.Fields!codigo) = 4 Then
lcCodigo = ""
End If
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    lcTexto1 = ""
                    lcTexto2 = ""
                    '***********************************Saldo Inicial (inicio)**************************************************
                    If lb_ConsiderarSaldoInicialDelHistorico = True Then     'debb-10/12/2018
                        lnSaldoInicial = 0
                        Set rsErrores = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes(lcCodigo, lc_CodigoSismed, ldFechaHistoricoXmes, oConexion)
                        If rsErrores.RecordCount > 0 Then
                           If Not IsNull(rsErrores!STOCK_FIN) Then
                              lnSaldoInicial = rsErrores!STOCK_FIN
                           End If
                        End If
                        rsErrores.Close
                    Else
                        lnSaldoInicial = 0
                        'saldos-barre tabla historico mensual (final del mes anterior)
                        For lnFor = 1 To Len(lc_AlmacenesParaICI)
                            If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                               lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                               lnFor = Len(lc_AlmacenesParaICI)
                            Else
                                lcTexto1 = ""
                                Do While True
                                   If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                                      Exit Do
                                   Else
                                      lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                                      lnFor = lnFor + 1
                                   End If
                                Loop
                                lnIdAlmacenRep = Val(lcTexto1)
                            End If
                            If lnIdAlmacenRep > 3 Then
                                Set rsErrores = mo_ReglasFarmacia.FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(lnIdProducto, lnIdAlmacenRep, ldFechaHistoricoXmes, oConexion)
                                Do While Not rsErrores.EOF
                                    lnSaldoInicial = lnSaldoInicial + rsErrores.Fields!saldo
                                    rsErrores.MoveNext
                                Loop
                                rsErrores.Close
                            End If
                        Next
                    End If
                    '**************************************** saldo inicial (fin)************************************************
                    'saldos-barre movimiento
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaInicio
                       If rsReporte.Fields!MovTipo = "S" Then
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                            lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                          End If
                       Else
                          If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                             lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                          End If
                       End If
                       Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    DebugAgregar lcCodigo, "E", "saldoIn", 0, lnSaldoInicial
                    lnDonacionesSaldoI = 0
                    If mb_EsDonaciones = True Then
                       lnDonacionesSaldoI = lnSaldoInicial
                    End If
'                    'huaral nov 15
'                    If Trim(lc_CodigoSismed) = "07637F01" Then
'                       If lcAnioMes = "201511" Then
'                          If Val(lcCodigo) = 3213 Then
'                             lnSaldoInicial = lnSaldoInicial + 1
'                          ElseIf Val(lcCodigo) = 11368 Then
'                             lnSaldoInicial = lnSaldoInicial + 1
'                          ElseIf Val(lcCodigo) = 23522 Then
'                             lnSaldoInicial = -10
'                          End If
'                       End If
'                    End If
                    '****** Movimientos en el Rango de Fechas****************************************
                    lnDonacionesOtrIng = 0: lnDonacionesIng = 0: lnDonacionesSal = 0: lnDonacionesOtrSal = 0
                    lnIngresos = 0: LnDevolucionesP = 0: TotIngresos = 0
                    LnVentas = 0: lnSis = 0: lnSoat = 0: LnConvenio = 0: lnCreditoH = 0: lnDefensaN = 0
                    LnOsDevol = 0: LnOsVencim = 0: LnOsMerma = 0: LnExonerac = 0: LnIntervencionS = 0
                    LnOtrasS = 0: TotSalidas = 0
                    If Not rsReporte.EOF Then
                        
                        'lcAnioMes = Format(rsReporte.Fields!FechaCreacion, "yyyy") & Format(rsReporte.Fields!FechaCreacion, "mm")
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
If rsReporte.Fields!movNumero = "150051214" Then
   lbPrimeraVez = True
End If
                           lnPrecio = rsReporte!precio
                           lbPrimeraVez = True
                           If rsReporte.Fields!MovTipo = "S" Then
                              If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenOrigen)) & "/") > 0 Then
                                    If mb_ConsiderarRecalculo = True Then
                                        '********* con recalculo
                                        lbEsPrimerLoteDelItem = IIf(lcTexto2 = rsReporte.Fields!movNumero And lcTexto2Fmovimiento = rsReporte!fechaCreacion, False, True)
                                        
                                        '
                                        lcTexto1 = rsReporte.Fields!MovTipo
                                        lcTexto2 = rsReporte.Fields!movNumero
                                        lcTexto2Fmovimiento = rsReporte!fechaCreacion
                                        'Busca si tiene Pagos
                                        lbNoEntro = True
                                        Set rsTmp = mo_ReglasFarmacia.FacturacionBienesPagoSeleccionarPorMovNumeroProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                        If rsTmp.RecordCount > 0 Then
                                           If rsTmp.Fields!IdComprobantePago > 0 And rsTmp.Fields!idEstadoFacturacion = 4 Then
                                              lbNoEntro = False
                                              lbPrimeraVez = False
                                              If lbEsPrimerLoteDelItem = True Then
                                                    rsTmp.MoveFirst
                                                    Do While Not rsTmp.EOF
                                                          Select Case rsReporte.Fields!idTipoConcepto
                                                          Case 10       'Ventas
                                                               LnVentas = LnVentas + rsTmp.Fields!CantidadPagar
                                                          Case 13       'Sis
                                                               lnSis = lnSis + rsTmp.Fields!CantidadPagar
                                                          Case 14       'Soat
                                                               lnSoat = lnSoat + rsTmp.Fields!CantidadPagar
                                                          Case 23, 26       'Convenios, Credito Personal
                                                               LnConvenio = LnConvenio + rsTmp.Fields!CantidadPagar
                                                          Case 17       'Credito Hospitalario
                                                               lnCreditoH = lnCreditoH + rsTmp.Fields!CantidadPagar
                                                          Case 22       'Defensa nacional
                                                               lnDefensaN = lnDefensaN + rsTmp.Fields!CantidadPagar
                                                          Case 7       'Devolucion x sobrestock
                                                               LnOsDevol = LnOsDevol + rsTmp.Fields!CantidadPagar
                                                          Case 5       'DEVOLUCION X VENCIMIENTO
                                                               LnOsVencim = LnOsVencim + rsTmp.Fields!CantidadPagar
                                                          Case 6       'DEVOLUCION X DETERIORO
                                                               LnOsMerma = LnOsMerma + rsTmp.Fields!CantidadPagar
                                                          Case 15       'Exoneraciones
                                                               LnExonerac = LnExonerac + rsTmp.Fields!CantidadPagar
                                                          Case 16       'Intervencion Sanitaria
                                                               LnIntervencionS = LnIntervencionS + rsTmp.Fields!CantidadPagar
                                                          Case Else     'Ajuste inventario
                                                               If mb_EsDonaciones = True Then
                                                                  lnDonacionesOtrSal = lnDonacionesOtrSal + rsTmp.Fields!CantidadPagar
                                                               Else
                                                                  LnOtrasS = LnOtrasS + rsTmp.Fields!CantidadPagar
                                                               End If
                                                          End Select
                                                          DebugAgregar lcCodigo, lcTexto1, lcTexto2, rsReporte!idTipoConcepto, rsTmp!CantidadPagar


                                                          'LnVentas = LnVentas + rsTmp.Fields!cantidadPagar
                                                          TotSalidas = TotSalidas + rsTmp.Fields!CantidadPagar
                                                          rsTmp.MoveNext
                                                    Loop
                                              End If
                                           End If
                                        End If
                                        'Busca si tiene algun seguro o exoneracion (Plan)
                                        If lbNoEntro = True Then
                                            lbPrimeraVez = False
                                            Set rsTmp = mo_ReglasFarmacia.FacturacionBienesFinancSeleccionarPorProducto(oConexion, rsReporte.Fields!movNumero, rsReporte.Fields!MovTipo, rsReporte.Fields!idProducto)
                                            If rsTmp.RecordCount = 0 Then
                                               'If lbNoEntro = True Then
                                                        Select Case rsReporte.Fields!idTipoConcepto
                                                        Case 10       'Ventas
                                                             LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                                        Case 13       'Sis
                                                             lnSis = lnSis + rsReporte.Fields!Cantidad
                                                        Case 14       'Soat
                                                             lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                                        Case 23, 26       'Convenios, Credito Personal
                                                             LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                                        Case 17       'Credito Hospitalario
                                                             lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
                                                        Case 22       'Defensa nacional
                                                             lnDefensaN = lnDefensaN + rsReporte.Fields!Cantidad
                                                        Case 7       'Devolucion x sobrestock
                                                             LnOsDevol = LnOsDevol + rsReporte.Fields!Cantidad
                                                        Case 5       'DEVOLUCION X VENCIMIENTO
                                                             LnOsVencim = LnOsVencim + rsReporte.Fields!Cantidad
                                                        Case 6       'DEVOLUCION X DETERIORO
                                                             LnOsMerma = LnOsMerma + rsReporte.Fields!Cantidad
                                                        Case 15       'Exoneraciones
                                                             LnExonerac = LnExonerac + rsReporte.Fields!Cantidad
                                                        Case 16       'Intervencion Sanitaria
                                                             LnIntervencionS = LnIntervencionS + rsReporte.Fields!Cantidad
                                                        Case Else     'Ajuste inventario
                                                             If mb_EsDonaciones = True Then
                                                                lnDonacionesOtrSal = lnDonacionesOtrSal + rsReporte.Fields!Cantidad
                                                             Else
                                                                LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                                             End If
                                                        End Select
                                                        DebugAgregar lcCodigo, lcTexto1, lcTexto2, rsReporte!idTipoConcepto, rsReporte!Cantidad
                                                        'LnVentas = LnVentas + rsReporte.Fields!cantidad
                                                        TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                        
                                               'End If
                                            Else
                                              
                                                rsTmp.MoveFirst
                                                Do While Not rsTmp.EOF
                                                    lnTipoConceptoFarm = mo_ReglasFacturacion.FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia(oConexion, rsTmp.Fields!idFuenteFinanciamiento)
                                                    lnCantidadDevolucion = 0

                                                    Select Case lnTipoConceptoFarm
                                                    Case 10       'Ventas
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsReporte.Fields!Cantidad
                                                         LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                                         TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                    Case 13       'Sis
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsReporte.Fields!Cantidad
                                                         lnSis = lnSis + rsReporte.Fields!Cantidad
                                                         TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                    Case 14      'Soat
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsReporte.Fields!Cantidad
                                                         lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                                         TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                    Case 16       'Intervencion Sanitaria
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsTmp.Fields!CantidadFinanciada
                                                         LnIntervencionS = LnIntervencionS + rsTmp.Fields!CantidadFinanciada
                                                         TotSalidas = TotSalidas + rsTmp.Fields!CantidadFinanciada
                                                    Case 17       'Credito Hospitalario
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsReporte.Fields!Cantidad
                                                         lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
                                                         TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                    Case 23, 26       'Convenios, Credito Personal
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsReporte.Fields!Cantidad
                                                         LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                                         TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                                    Case 0      'Exoneraciones
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsTmp.Fields!CantidadFinanciada
                                                         '**no se sabe la CANTIDAD EXONERADA solo el IMPORTE EXONERADO
                                                         If rsTmp.Fields!CantidadFinanciada > 0 Then
                                                            LnExonerac = LnExonerac + rsTmp.Fields!CantidadFinanciada
                                                         Else
                                                            LnExonerac = LnExonerac + rsTmp.Fields!CantidadFinanciada
                                                            LnExonerac = LnExonerac - rsTmp.Fields!CantidadFinanciada
                                                            TotSalidas = TotSalidas - rsTmp.Fields!CantidadFinanciada
                                                         End If
                                                         TotSalidas = TotSalidas + rsTmp.Fields!CantidadFinanciada
                                                    Case Else
                                                         DebugAgregar lcCodigo, lcTexto1, lcTexto2, lnTipoConceptoFarm, rsTmp.Fields!CantidadFinanciada
                                                         LnOtrasS = LnOtrasS + rsTmp.Fields!CantidadFinanciada
                                                         TotSalidas = TotSalidas + rsTmp.Fields!CantidadFinanciada
                                                    End Select
                                                    
                                                    rsTmp.MoveNext
                                                Loop
                                            End If
                                        End If
                                        If lbPrimeraVez = False Then
                                            'Do While Not rsReporte.EOF And lcTexto1 = rsReporte.Fields!MovTipo And lcTexto2 = rsReporte.Fields!movNumero And lnIdProducto = rsReporte.Fields!idProducto
                                               rsReporte.MoveNext
                                             '  If rsReporte.EOF Then
                                             '     Exit Do
                                              ' End If
                                            'Loop
                                        End If
                                    End If
                                    If lbPrimeraVez = True Then
                                        '******** sin recalculo
                                        Select Case rsReporte.Fields!idTipoConcepto
                                        Case 3           'Salidas x Donaciones
                                             lnDonacionesSal = lnDonacionesSal + rsReporte.Fields!Cantidad
                                        Case 10       'Ventas
                                             LnVentas = LnVentas + rsReporte.Fields!Cantidad
                                        Case 13       'Sis
                                             lnSis = lnSis + rsReporte.Fields!Cantidad
                                        Case 14       'Soat
                                             lnSoat = lnSoat + rsReporte.Fields!Cantidad
                                        Case 23, 26      'Convenios, Credito Personal
                                             LnConvenio = LnConvenio + rsReporte.Fields!Cantidad
                                        Case 17       'Credito Hospitalario
                                             lnCreditoH = lnCreditoH + rsReporte.Fields!Cantidad
                                        Case 22       'Defensa nacional
                                             lnDefensaN = lnDefensaN + rsReporte.Fields!Cantidad
                                        Case 7       'Otras salidas Devolucion
                                             LnOsDevol = LnOsDevol + rsReporte.Fields!Cantidad
                                        Case 5       'Otras salidas Vencimiento
                                             LnOsVencim = LnOsVencim + rsReporte.Fields!Cantidad
                                        Case 6       'Otras salidas Merma
                                             LnOsMerma = LnOsMerma + rsReporte.Fields!Cantidad
                                        Case 15       'Exoneraciones
                                             LnExonerac = LnExonerac + rsReporte.Fields!Cantidad
                                        Case 16       'Intervencion Sanitaria
                                             LnIntervencionS = LnIntervencionS + rsReporte.Fields!Cantidad
                                        Case Else
                                             If mb_EsDonaciones = True Then
                                                lnDonacionesOtrSal = lnDonacionesOtrSal + rsTmp.Fields!CantidadPagar
                                             Else
                                                 LnOtrasS = LnOtrasS + rsReporte.Fields!Cantidad
                                             End If
                                        End Select
                                        DebugAgregar lcCodigo, lcTexto1, lcTexto2, rsReporte.Fields!idTipoConcepto, rsReporte.Fields!Cantidad
                                        If rsReporte.Fields!IdAlmacenDestino = 10 And rsReporte.Fields!idTipoConcepto = 4 And mb_ConsideraOSH = False Then
                                           'destino=otros servicios hospital, tipoConcepto=distribucion
                                        Else
                                            TotSalidas = TotSalidas + rsReporte.Fields!Cantidad
                                        End If
                                    End If
                              End If
                           Else
                              If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsReporte.Fields!IdAlmacenDestino)) & "/") > 0 Then
                                    Select Case rsReporte.Fields!idTipoConcepto
                                    Case 3           'Ingresos x Donaciones
                                         lnDonacionesIng = lnDonacionesIng + rsReporte.Fields!Cantidad
                                         TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                    Case 19        'Inventario
                                         If mb_EsDonaciones = True Then
                                            lnDonacionesOtrIng = lnDonacionesOtrIng + rsReporte.Fields!Cantidad
                                         Else
                                            lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                         End If
                                    Case 21        'Devolucion de Pacientes
                                         'Actualizado 15102014
                                         LnDevolucionesP = LnDevolucionesP + rsReporte.Fields!Cantidad
                                         TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                    Case Else        'Ingresos, ajustes de inventario (Sismed/Donaciones)
                                         If mb_EsDonaciones = True Then
                                            lnDonacionesOtrIng = lnDonacionesOtrIng + rsReporte.Fields!Cantidad
                                         Else
                                            lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                                         End If
                                         TotIngresos = TotIngresos + rsReporte.Fields!Cantidad
                                    End Select
                                    DebugAgregar lcCodigo, rsReporte.Fields!MovTipo, rsReporte.Fields!movNumero, rsReporte.Fields!idTipoConcepto, rsReporte.Fields!Cantidad
                               End If
                           End If
                           If lbPrimeraVez = True Then
                               rsReporte.MoveNext
                           End If
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
                    If Not rsReporte.EOF Then
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fechaCreacion <= mda_FechaFin
                           Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                           rsReporte.MoveNext
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
                    '
                    'lnPrecio = 0
                    If mb_EsDonaciones = True Then
                        TotSalidas = 0: TotIngresos = 0: lnSaldoInicial = 0
                        Set rsTmp = mo_reglasComunes.CatalogoBienesInsumosFiltrarDonacionesXcodigo(lcCodigo)
                    Else
                        Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                    End If
                    If rsTmp.RecordCount > 0 And lnPrecio = 0 Then
                       lnPrecio = rsTmp.Fields!PrecioUnitario
                       If mb_EsDonaciones = False Then
                          lcTipoMI = IIf(rsTmp!TipoProducto = 1, "Insumo", "Medicamento")
                       End If
                    End If
                    lcTipoMI = " "
                    If rsTmp.RecordCount > 0 And mb_EsDonaciones = False Then
                       If Not IsNull(rsTmp!TipoProducto) Then
                          lcTipoMI = IIf(rsTmp!TipoProducto = 1, "Insumo", "Medicamento")
                       End If
                    End If
                    rsTmp.Close
                    '
'                    If InStr(lc_AlmacenesParaICI, "/") > 0 Then
'                       lnIdAlmacen = Val(Left(lc_AlmacenesParaICI, InStr(lc_AlmacenesParaICI, "/") - 1)) 'toma la primera farmacia para las FECHAS DE VENCIMIENTO
'                    Else
'                       lnIdAlmacen = Val(lc_AlmacenesParaICI)
'                    End If
'                    ldFechaVencimiento = Date
'                    lcLoteXitem = ""
'                    Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacen, 0, lcCodigo)
'                    If rsTmp.RecordCount > 0 Then
'                       ldFechaVencimiento = rsTmp.Fields!FechaVencimiento
'                       lcLoteXitem = rsTmp.Fields!lote
'                    End If
'                    rsTmp.Close
                    ldFechaVencimiento = Date
                    lcLoteXitem = ""
                    If InStr(lc_AlmacenesParaICI, "/") > 0 Then
                        ldFechaVencimiento = CDate("31/12/2029")
                       
                        For lnFor = 1 To Len(lc_AlmacenesParaICI)
                            If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                               lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                               lnFor = Len(lc_AlmacenesParaICI)
                            Else
                                lcTexto1 = ""
                                Do While True
                                   If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                                      Exit Do
                                   Else
                                      lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                                      lnFor = lnFor + 1
                                   End If
                                Loop
                                lnIdAlmacenRep = Val(lcTexto1)
                            End If
                            
                            Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacenRep, 0, lcCodigo)
                            If rsTmp.RecordCount > 0 Then
                               If rsTmp.Fields!fechaVencimiento < ldFechaVencimiento Then
                                    ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                                    lcLoteXitem = rsTmp.Fields!Lote
                               End If
                            End If
                            rsTmp.Close
                            
                        Next
                       
                       
                    Else
                        lnIdAlmacen = Val(lc_AlmacenesParaICI)
                        Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacen, 0, lcCodigo)
                        If rsTmp.RecordCount > 0 Then
                           ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                           lcLoteXitem = rsTmp.Fields!Lote
                        End If
                        rsTmp.Close
                    End If
                    ldDonacionFechaVctoUlt = Date
                    If mb_EsDonaciones = True Then
                       ldDonacionFechaVctoUlt = ldFechaVencimiento
                    End If
                    '
                    lbContinua = True
                    If mb_ConsiderarSinMovimientos = False Then
                       If TotIngresos = 0 And TotSalidas = 0 And ((lnDonacionesSaldoI + lnDonacionesOtrIng + lnDonacionesIng) - (lnDonacionesSal + lnDonacionesOtrSal)) = 0 Then
                          lbContinua = False
                       End If
                    End If
                    '
                    If lb_NOconsiderarSALDOcero = True Then
                       If (lnDonacionesOtrIng + lnDonacionesIng + TotIngresos) = 0 And _
                                                   (lnDonacionesSal + lnDonacionesOtrSal + TotSalidas) = 0 And _
                                                                  (lnDonacionesSaldoI + lnSaldoInicial) = 0 Then
                          lbContinua = False
                       End If
                    End If
                    '
                    If lbContinua Then
If Val(lcCodigo) = 4 Then
lcTexto1 = ""
End If
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!codigo = lcCodigo
                        mrs_Tmp.Fields!nombre = lcNombre
                        mrs_Tmp.Fields!precio = lnPrecio
                        If mb_EsDonaciones = True Then
                            mrs_Tmp.Fields!saldoI = lnDonacionesSaldoI
                            mrs_Tmp.Fields!ingresos = lnDonacionesOtrIng + lnDonacionesIng
                            mrs_Tmp.Fields!DevolucionesP = lnDonacionesSal
                            mrs_Tmp.Fields!otrasS = lnDonacionesOtrSal
                            mrs_Tmp.Fields!TotSalidas = lnDonacionesSal + lnDonacionesOtrSal
                            mrs_Tmp.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                        Else
                            mrs_Tmp.Fields!saldoI = lnSaldoInicial
                            mrs_Tmp.Fields!ingresos = lnIngresos + LnDevolucionesP
                            mrs_Tmp.Fields!DevolucionesP = 0   'LnDevolucionesP
                            mrs_Tmp.Fields!TotIngresos = TotIngresos
                            mrs_Tmp.Fields!Ventas = LnVentas
                            mrs_Tmp.Fields!sis = lnSis
                            mrs_Tmp.Fields!soat = lnSoat
                            mrs_Tmp.Fields!convenio = LnConvenio
                            mrs_Tmp.Fields!creditoH = lnCreditoH
                            mrs_Tmp.Fields!defensaN = lnDefensaN
                            mrs_Tmp.Fields!OsDevol = LnOsDevol
                            mrs_Tmp.Fields!OsVencim = LnOsVencim
                            mrs_Tmp.Fields!OsMerma = LnOsMerma
                            mrs_Tmp.Fields!Exonerac = LnExonerac
                            mrs_Tmp.Fields!IntervencionS = LnIntervencionS
                            mrs_Tmp.Fields!otrasS = LnOtrasS
                            mrs_Tmp.Fields!TotSalidas = TotSalidas
                            mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                        End If
                        mrs_Tmp.Fields!tipo = lcTipoMI
                        mrs_Tmp.Update
                        '***************************grabar el ICI-detalle**************************************

                        lnSaldofinal = lnSaldoInicial + lnIngresos + LnDevolucionesP - TotSalidas
                        lnSaldoFinalD = (lnDonacionesSaldoI + lnDonacionesOtrIng + lnDonacionesIng) - (lnDonacionesSal + lnDonacionesOtrSal)
                        lnPrecioItem = lnPrecio
                        oRsFox.AddNew
                        oRsFox.Fields!CODIGO_EJE = lcDisa
                        oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                        oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                        oRsFox.Fields!annoMes = lcAnioMes
                        oRsFox.Fields!codigo_med = Left(lcCodigo, 7)
                        oRsFox.Fields!saldo = lnSaldoInicial
                        oRsFox.Fields!precio = lnPrecioItem
                        oRsFox.Fields!ingre = lnIngresos + LnDevolucionesP
                        oRsFox.Fields!reingre = 0     'LnDevolucionesP
                        oRsFox.Fields!VENTA = LnVentas
                        oRsFox.Fields!sis = lnSis
                        oRsFox.Fields!intersan = LnIntervencionS
                        oRsFox.Fields!fac_perd = 0                        'falta
                        oRsFox.Fields!DEFNAC = lnDefensaN
                        oRsFox.Fields!exo = LnExonerac
                        oRsFox.Fields!soat = lnSoat
                        oRsFox.Fields!credHosp = lnCreditoH
                        oRsFox.Fields!otr_conv = LnConvenio
                        oRsFox.Fields!DEVOL = LnOsDevol
                        oRsFox.Fields!vencido = 0
                        oRsFox.Fields!merma = 0
                        oRsFox.Fields!distri = 0
                        oRsFox.Fields!transf = 0
                        oRsFox.Fields!ventaInst = 0
                        oRsFox.Fields!DEV_VEN = LnOsVencim
                        oRsFox.Fields!DEV_MERMA = LnOsMerma
                        oRsFox.Fields!otras_sal = LnOtrasS
                        oRsFox.Fields!STOCK_FIN = lnSaldofinal
                        oRsFox.Fields!stock_fin1 = lnSaldofinal
                        oRsFox.Fields!REQ = TotSalidas                               'falta preguntar al MINSA
                        oRsFox.Fields!Total = TotSalidas
                        oRsFox.Fields!do_saldo = lnDonacionesSaldoI
                        If mb_EsDonaciones = False Then
                           oRsFox.Fields!FEC_EXP = ldFechaVencimiento
                           oRsFox.Fields!do_saldo = 0
                        End If
                        
                        oRsFox.Fields!do_ingre = lnDonacionesOtrIng + lnDonacionesIng
                        oRsFox.Fields!do_con = lnDonacionesSal
                        oRsFox.Fields!do_otr = lnDonacionesOtrSal
                        oRsFox.Fields!do_tot = lnDonacionesSal + lnDonacionesOtrSal
                        oRsFox.Fields!do_stk = 0
                        If mb_EsDonaciones = True Then
                           oRsFox.Fields!do_fecExp = ldDonacionFechaVctoUlt
                           oRsFox.Fields!do_stk = lnSaldoFinalD
                        End If
                        oRsFox.Fields!fecha = Date
                        oRsFox.Fields!Usuario = " "
                        oRsFox.Fields!indiProc = " "
                        oRsFox.Fields!SIT = "1"
                        oRsFox.Fields!indiSiga = " "
                        oRsFox.Fields!dstkCero = 0
                        oRsFox.Fields!mptoRepo = 0
                        oRsFox.Update
                        'FormDetL
                        oRsFox1.AddNew
                        oRsFox1.Fields!CODIGO_EJE = lcDisa
                        oRsFox1.Fields!CODIGO_PRE = lc_CodigoSismed
                        oRsFox1.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                        oRsFox1.Fields!annoMes = lcAnioMes
                        oRsFox1.Fields!codigo_med = Left(lcCodigo, 7)
                        oRsFox1.Fields!Lote = lcLoteXitem
                        oRsFox1.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                        oRsFox1.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                        oRsFox1.Fields!SIT = "1"
                        oRsFox1.Update
                        'FormDetM
                        oRsFox2.AddNew
                        oRsFox2.Fields!CODIGO_EJE = lcDisa
                        oRsFox2.Fields!CODIGO_PRE = lc_CodigoSismed
                        oRsFox2.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                        oRsFox2.Fields!annoMes = lcAnioMes
                        oRsFox2.Fields!codigo_med = Left(lcCodigo, 7)
                        oRsFox2.Fields!Lote = lcLoteXitem
                        oRsFox2.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                        oRsFox2.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                        oRsFox2.Fields!SIT = "1"
                        oRsFox2.Update
                        '***************************grabar el ICI-detalle**************************************
                    End If
                    'Graba Datos en Temporal
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
GrabaParametro206 "procesa ICI antes carga sin movimie"
                'Carga Saldos de Items sin Movimientos
                If mb_ConsiderarSinMovimientos = True Then
                   Set rsTmp9 = mo_ReglasFarmacia.farmSaldoMensualFiltrarFecha(ldFechaHistoricoXmes, lc_AlmacenesParaICI, lnIdAlmacenRep)

                   If rsTmp9.RecordCount > 0 Then
                      rsTmp9.MoveFirst
                      Do While Not rsTmp9.EOF
                         lnIdAlmacen = rsTmp9.Fields!IdAlmacen
                         lcCodigo = rsTmp9.Fields!codigo
If Val(lcCodigo) = 56 Or Val(lcCodigo) = 1974 Then
lcNombre = ""
End If
                         lcNombre = rsTmp9.Fields!nombre
                         
                         
                         lnSaldoInicial = 0
                         Do While Not rsTmp9.EOF And lcCodigo = rsTmp9.Fields!codigo
                            If InStr(lc_AlmacenesParaICI, "/" & Trim(str(rsTmp9!IdAlmacen)) & "/") > 0 Then
                               lnSaldoInicial = lnSaldoInicial + rsTmp9.Fields!saldo
                            End If
                            rsTmp9.MoveNext
                            If rsTmp9.EOF Then
                               Exit Do
                            End If
                         Loop
                         
                         
                         
                         
                         If mrs_Tmp.RecordCount > 0 Then
                            mrs_Tmp.MoveFirst
                            mrs_Tmp.Find "codigo='" & lcCodigo & "'"
                            If mrs_Tmp.EOF Then
                                
                                If lb_ConsiderarSaldoInicialDelHistorico = True Then     'debb-10/12/2018
                                   lnSaldoInicial = 0
                                   Set rsErrores = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes(lcCodigo, lc_CodigoSismed, ldFechaHistoricoXmes, oConexion)
                                   If rsErrores.RecordCount > 0 Then
                                      If Not IsNull(rsErrores!STOCK_FIN) Then
                                         lnSaldoInicial = rsErrores!STOCK_FIN
                                      End If
                                   End If
                                   rsErrores.Close
                                End If
                                
                                lnSaldofinal = lnSaldoInicial
                                lnDonacionesSaldoI = lnSaldoInicial
                                lnSaldoFinalD = lnSaldoInicial
                                '
                                lnPrecio = 0
                                If rsTmp.State = 1 Then rsTmp.Close
                                If mb_EsDonaciones = True Then
                                    Set rsTmp = mo_reglasComunes.CatalogoBienesInsumosFiltrarDonacionesXcodigo(lcCodigo)
                                Else
                                    Set rsTmp = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodigo, 1)
                                End If
                                If rsTmp.RecordCount > 0 Then
                                   lnPrecio = rsTmp.Fields!PrecioUnitario
                                End If
                                rsTmp.Close
                                '
If Val(lcCodigo) = 4 Then
ldFechaVencimiento = Date
End If
                                ldFechaVencimiento = Date
                                lcLoteXitem = ""
                                Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacen, 0, lcCodigo)
                                If rsTmp.RecordCount > 0 Then
                                   ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                                   lcLoteXitem = rsTmp.Fields!Lote
                                End If
                                rsTmp.Close
                                ldDonacionFechaVctoUlt = Date
                                If mb_EsDonaciones = True Then
                                   ldDonacionFechaVctoUlt = ldFechaVencimiento
                                End If
                                '

                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!codigo = lcCodigo
                                mrs_Tmp.Fields!nombre = lcNombre
                                mrs_Tmp.Fields!precio = lnPrecio
                                If mb_EsDonaciones = True Then
                                    mrs_Tmp.Fields!saldoI = lnDonacionesSaldoI
                                    mrs_Tmp.Fields!ingresos = 0
                                    mrs_Tmp.Fields!DevolucionesP = 0
                                    mrs_Tmp.Fields!otrasS = 0
                                    mrs_Tmp.Fields!TotSalidas = 0
                                    mrs_Tmp.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                                Else
                                    mrs_Tmp.Fields!saldoI = lnSaldoInicial
                                    mrs_Tmp.Fields!ingresos = 0
                                    mrs_Tmp.Fields!DevolucionesP = 0
                                    mrs_Tmp.Fields!TotIngresos = 0
                                    mrs_Tmp.Fields!Ventas = 0
                                    mrs_Tmp.Fields!sis = 0
                                    mrs_Tmp.Fields!soat = 0
                                    mrs_Tmp.Fields!convenio = 0
                                    mrs_Tmp.Fields!creditoH = 0
                                    mrs_Tmp.Fields!defensaN = 0
                                    mrs_Tmp.Fields!OsDevol = 0
                                    mrs_Tmp.Fields!OsVencim = 0
                                    mrs_Tmp.Fields!OsMerma = 0
                                    mrs_Tmp.Fields!Exonerac = 0
                                    mrs_Tmp.Fields!IntervencionS = 0
                                    mrs_Tmp.Fields!otrasS = 0
                                    mrs_Tmp.Fields!TotSalidas = 0
                                    mrs_Tmp.Fields!fechaVencimiento = ldFechaVencimiento
                                End If
                                mrs_Tmp.Update
                                '***************************grabar el ICI-detalle**************************************
                                lnPrecioItem = lnPrecio
                                oRsFox.AddNew
                                oRsFox.Fields!CODIGO_EJE = lcDisa
                                oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox.Fields!annoMes = lcAnioMes
                                oRsFox.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox.Fields!saldo = lnSaldoInicial
                                oRsFox.Fields!precio = lnPrecioItem
                                oRsFox.Fields!ingre = 0
                                oRsFox.Fields!reingre = 0
                                oRsFox.Fields!VENTA = 0
                                oRsFox.Fields!sis = 0
                                oRsFox.Fields!intersan = 0
                                oRsFox.Fields!fac_perd = 0                        'falta
                                oRsFox.Fields!DEFNAC = 0
                                oRsFox.Fields!exo = 0
                                oRsFox.Fields!soat = 0
                                oRsFox.Fields!credHosp = 0
                                oRsFox.Fields!otr_conv = 0
                                oRsFox.Fields!DEVOL = 0
                                oRsFox.Fields!vencido = 0
                                oRsFox.Fields!merma = 0
                                oRsFox.Fields!distri = 0
                                oRsFox.Fields!transf = 0
                                oRsFox.Fields!ventaInst = 0
                                oRsFox.Fields!DEV_VEN = 0
                                oRsFox.Fields!DEV_MERMA = 0
                                oRsFox.Fields!otras_sal = 0
                                oRsFox.Fields!STOCK_FIN = lnSaldofinal
                                oRsFox.Fields!stock_fin1 = lnSaldofinal
                                oRsFox.Fields!REQ = 0
                                oRsFox.Fields!Total = 0
                                If mb_EsDonaciones = False Then
                                   oRsFox.Fields!FEC_EXP = ldFechaVencimiento
                                End If
                                If lnDonacionesSaldoI > 0 And lnDonacionesSaldoI < 9999999 Then
                                   oRsFox.Fields!do_saldo = 0
                                Else
                                   oRsFox.Fields!do_saldo = 0
                                End If
                                oRsFox.Fields!do_ingre = 0
                                oRsFox.Fields!do_con = 0
                                oRsFox.Fields!do_otr = 0
                                oRsFox.Fields!do_tot = 0
                                If lnDonacionesSaldoI > 0 And lnDonacionesSaldoI < 9999999 Then
                                   oRsFox.Fields!do_stk = 0
                                Else
                                   oRsFox.Fields!do_stk = 0
                                End If
                                If mb_EsDonaciones = True Then
                                   oRsFox.Fields!do_fecExp = ldDonacionFechaVctoUlt
                                End If
                                oRsFox.Fields!fecha = Date
                                oRsFox.Fields!Usuario = " "
                                oRsFox.Fields!indiProc = " "
                                oRsFox.Fields!SIT = "1"
                                oRsFox.Fields!indiSiga = " "
                                oRsFox.Fields!dstkCero = 0
                                oRsFox.Fields!mptoRepo = 0
                                oRsFox.Update
                                'FormDetL
                                oRsFox1.AddNew
                                oRsFox1.Fields!CODIGO_EJE = lcDisa
                                oRsFox1.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox1.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox1.Fields!annoMes = lcAnioMes
                                oRsFox1.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox1.Fields!Lote = lcLoteXitem
                                oRsFox1.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                                If lnSaldoFinalD < 9999999 Or lnSaldofinal < 9999999 Then
                                   oRsFox1.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                                Else
                                   oRsFox1.Fields!saldo = 0
                                End If
                                oRsFox1.Fields!SIT = "1"
                                oRsFox1.Update
                                'FormDetM
                                oRsFox2.AddNew
                                oRsFox2.Fields!CODIGO_EJE = lcDisa
                                oRsFox2.Fields!CODIGO_PRE = lc_CodigoSismed
                                oRsFox2.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                                oRsFox2.Fields!annoMes = lcAnioMes
                                oRsFox2.Fields!codigo_med = Left(lcCodigo, 7)
                                oRsFox2.Fields!Lote = lcLoteXitem
                                oRsFox2.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                                If lnSaldoFinalD < 9999999 Or lnSaldofinal < 9999999 Then
                                   oRsFox2.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldofinal)
                                Else
                                   oRsFox2.Fields!saldo = 0
                                End If
                                oRsFox2.Fields!SIT = "1"
                                oRsFox2.Update
                            End If
                         End If
                      Loop
                   End If
                   rsTmp9.Close
                End If
                
                
GrabaParametro206 "procesa ICI antes Unidosis.."

                    CodigosUnidosisDesagregaEnMedicInsumos oRsFox, oRsFox1, oRsFox2, lcDisa, mb_EsDonaciones, _
                                                           lcAnioMes, oConexion, mrs_Tmp
                
GrabaParametro206 "procesa ICI antes FuaPaquetes.."

                If Val(lcEstablecimiento) = 4370 Or Val(lcEstablecimiento) = 5195 Then   'hbt/hbt1
                'debb-10/12/2018
                FuaPaquetesFarmaciaDesagregaEnMedicInsumos oRsFox, oRsFox1, oRsFox2, lcDisa, mb_EsDonaciones, _
                                                              lcAnioMes, oConexion, mrs_Tmp
                End If
                If lb_SeGrabaICImensual = True And lcBuscaParametro.SeleccionaFilaParametro(575) = "S" Then
GrabaParametro206 "procesa ICI antes de formDet"
                    mo_ReglasFarmacia.Farm_formDetActualizar oRsFox, lcAnioMes, lc_CodigoSismed, oConexion, oRsFox1
                    lb_EsUnIciHistorico = True
                End If
GrabaParametro206 "procesa ICI antes ICI cabecera"



                '***************************grabar el ICI- cabecera*************************************
                Dim lnRec_crehos As Long, lnRec_soat As Long, lnRec_vtas As Long, lnRec_otrcon As Long
                Dim lnRec_sis As Long, lnRec_ints As Long, lnRec_dn As Long, lnRec_Exo As Long
                lnRec_crehos = 0: lnRec_soat = 0: lnRec_vtas = 0: lnRec_otrcon = 0
                lnRec_sis = 0: lnRec_ints = 0: lnRec_dn = 0: lnRec_Exo = 0:
                Set rsTmp = mo_ReglasFarmacia.farmMovimientoFiltrarXfechas(mda_FechaInicio, mda_FechaFin, oConexion)
                If rsTmp.RecordCount > 0 Then
                   rsTmp.MoveFirst
                   Do While Not rsTmp.EOF
                      Select Case rsTmp.Fields!IdTipoFinanciamiento
                      Case sghTipoFinanciamiento.sghSIS
                           lnRec_sis = lnRec_sis + 1
                      Case sghTipoFinanciamiento.sghSOAT
                           lnRec_soat = lnRec_soat + 1
                      Case sghTipoFinanciamiento.sghPacienteNormal
                           If rsTmp.Fields!idFuenteFinanciamiento = 5 Then
                              lnRec_crehos = lnRec_crehos + 1
                           Else
                              lnRec_vtas = lnRec_vtas + 1
                           End If
                      Case sghTipoFinanciamiento.sghServicioSocial
                           lnRec_Exo = lnRec_Exo + 1
                      Case Else
                           lnRec_otrcon = lnRec_otrcon + 1
                      End Select
                      rsTmp.MoveNext
                   Loop
                End If
                rsTmp.Close
                Set rsTmp = mo_ReglasFarmacia.farmMovimientoFiltrarIntervSanitaria(mda_FechaInicio, mda_FechaFin, oConexion)
                lnRec_ints = rsTmp.RecordCount
                rsTmp.Close
                '
                mo_ReglasFarmacia.dbfFormatoSeleccionarTodos oRsFox, oConexionFox
                oRsFox.AddNew
                oRsFox.Fields!CODIGO_EJE = lcDisa
                oRsFox.Fields!CODIGO_PRE = lc_CodigoSismed
                oRsFox.Fields!annoMes = lcAnioMes
                oRsFox.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                oRsFox.Fields!tipo_pre = "F"
                oRsFox.Fields!rec_vtas = lnRec_vtas
                oRsFox.Fields!rec_sis = lnRec_sis
                oRsFox.Fields!rec_ints = lnRec_ints
                oRsFox.Fields!rec_dn = lnRec_dn
                oRsFox.Fields!rec_Exo = lnRec_Exo
                oRsFox.Fields!rec_soat = lnRec_soat
                oRsFox.Fields!rec_crehos = lnRec_crehos
                oRsFox.Fields!rec_otrcon = lnRec_otrcon
                oRsFox.Fields!indiProc = "A"
                oRsFox.Fields!fecha = Date
                'oRsFox.Fields!fechaUlt = Date
                oRsFox.Fields!vers = "V2.0 04102011"
                oRsFox.Fields!SIT = "1"
                oRsFox.Fields!fdesde = Format(mda_FechaInicio, "dd/mm/yyyy")
                oRsFox.Fields!fhasta = Format(mda_FechaFin, "dd/mm/yyyy")
                oRsFox.Fields!ctrlcal = "P"
                oRsFox.Fields!catalogo = Date               'falta
                oRsFox.Fields!codpto = lcEstablecimiento
                oRsFox.Fields!tip_ins = "E"
                oRsFox.Update
                '***************************grabar el ICI- cabecera*************************************
                lcTexto3 = "N° Recetas: Vtas: " & Trim(str(lnRec_vtas)) & " | Sis: " & Trim(str(lnRec_sis)) & " | Soat: " & _
                         Trim(str(lnRec_soat)) & " | Cred.Hosp: " & Trim(str(lnRec_crehos)) & " | Conv: " & _
                         Trim(str(lnRec_otrcon)) & " | Int.San: " & Trim(str(lnRec_ints)) & " | Exon: " & Trim(str(lnRec_Exo))
                'debb-11/08/2017

                 
          End If
          oRsFox.Close
          oRsFox1.Close
          oRsFox2.Close
          Set oRsFox = Nothing
          Set oRsFox1 = Nothing
          Set oRsFox2 = Nothing
          Set rsTmp9 = Nothing
          Set RsTmp989 = Nothing
          Set RsTmp988 = Nothing
          oConexion.Close
          oConexionFox.Close
          Exit Sub
ErrICI:
          MsgBox Err.Description & Chr(13) & "Codigo: " & lcCodigo & Chr(13) & "MovNumero: " & lcTexto2
          oRsFox.Close
          oRsFox1.Close
          oRsFox2.Close
          Set oRsFox = Nothing
          Set oRsFox1 = Nothing
          Set oRsFox2 = Nothing
          oConexion.Close
          oConexionFox.Close
          Exit Sub
          Resume
End Sub


Sub DebugAgregar(lcCodigoX As String, lcMovTipoX As String, lcMovNumeroX As String, lnTipoConceptoX As Long, lnCantidadX As Long)
    If mb_EnArchivoExcel = True And Val(lc_CodigoItem) > 0 Then
            rsDebug.AddNew
            rsDebug.Fields!codigo = lcCodigoX
            rsDebug.Fields!MovTipo = lcMovTipoX
            rsDebug.Fields!movNumero = lcMovNumeroX
            rsDebug.Fields!TipoConcepto = lnTipoConceptoX
            rsDebug.Fields!Cantidad = lnCantidadX
            rsDebug.Update
    End If
End Sub

'el CODIGO DEL ITEM es un Paquete de Farmacia,se desagrega en CODIGOS DIGEMID   'debb-02/12/2016
Sub FuaPaquetesFarmaciaDesagregaEnMedicInsumos(ByRef oRsFoxP As Recordset, ByRef oRsFox1P As Recordset, ByRef oRsFox2P As Recordset, _
                                               lcDisa As String, mb_EsDonaciones As Boolean, lcAnioMes As String, _
                                               oConexion As Connection, ByRef mrs_TmpP As Recordset)
        Dim oRsPqte As New Recordset
        Dim mo_ReglasFarmacia As New ReglasFarmacia
        Dim oRsFarmTmp1 As New Recordset
        Dim rsTmp99 As New Recordset
        Dim lcCodigo As String, lnCantidadBolsasR As Long, lnPrecio As Double, lnCantidadR As Long, lnTotal As Double
        Dim lcTipo As String, lcDx As String, lnDxNro As Integer, lcFormaF As String, lnCantidadC As Long
        Dim lnCantidadBolsasC As Long, lcLoteXitem As String
        Dim lnIngre99 As Long, lnReingre99 As Long, lnVenta99 As Long, lnSis99 As Long, lnIntersan99 As Long, lnDefNac99 As Long
        Dim lnExo99 As Long, lnSoat99 As Long, lcCodigo_med99 As String, lnTotSalidas99 As Long, lnSaldoFinal99 As Long
        Dim lnCredHosp99 As Long, lnOtr_conv99 As Long, lnDevol99 As Long, lnDev_ven99 As Long, lnDev_merma99 As Long, lnOtras_sal99 As Long
        Dim ldFechaVencimiento As Date, lnDonacionesSaldoI As Long, lnSaldoFinalD As Long, ldDonacionFechaVctoUlt As Date
        Dim rsTmp_99 As New Recordset, lcNombre_99 As String, lcTipoMI_99 As String, lcCodMed991 As String, lnSaldo99 As Long
        
        ldFechaVencimiento = Date
        lnDonacionesSaldoI = 0
        lnSaldoFinalD = 0
        ldDonacionFechaVctoUlt = Date
        oRsFoxP.Filter = "sit='1' and codigo_pre='" & lc_CodigoSismed & "'"
        If oRsFoxP.RecordCount > 0 Then
            Set oRsFarmTmp1 = sighentidades.CopyRecordset(oRsFoxP, "")
            oRsFoxP.MoveFirst
            Do While Not oRsFoxP.EOF
               lcCodigo_med99 = oRsFoxP!codigo_med
               If mo_ReglasFarmacia.CatalogoDIGEMIDesCodigoPaquete(lcCodigo_med99, oConexion) = True Then
                  lnSaldo99 = oRsFoxP!saldo
                  lnIngre99 = oRsFoxP!ingre
                  lnReingre99 = oRsFoxP!reingre
                  lnVenta99 = oRsFoxP!VENTA
                  lnSis99 = oRsFoxP!sis
                  lnIntersan99 = oRsFoxP!intersan
                  lnDefNac99 = oRsFoxP!DEFNAC
                  lnExo99 = oRsFoxP!exo
                  lnSoat99 = oRsFoxP!soat
                  lnCredHosp99 = oRsFoxP!credHosp
                  lnOtr_conv99 = oRsFoxP!otr_conv
                  lnDevol99 = oRsFoxP!DEVOL
                  lnDev_ven99 = oRsFoxP!DEV_VEN
                  lnDev_merma99 = oRsFoxP!DEV_MERMA
                  lnOtras_sal99 = oRsFoxP!otras_sal
                  Set oRsPqte = mo_ReglasFarmacia.CatalogoDIGEMIDdevuelveITEMS(lcCodigo_med99)
                  If oRsPqte.RecordCount > 0 Then
                     oRsPqte.MoveFirst
                     Do While Not oRsPqte.EOF
                        lnPrecio = oRsPqte!precio
                        If oRsFarmTmp1.RecordCount > 0 Then
                            oRsFarmTmp1.MoveFirst
                            oRsFarmTmp1.Find "codigo_med='" & oRsPqte!codigo & "'"
                        End If
                        If oRsFarmTmp1.EOF Then
                           oRsFarmTmp1.AddNew
                           oRsFarmTmp1.Fields!codigo_med = oRsPqte!codigo
                           oRsFarmTmp1.Fields!precio = oRsPqte!precio
                           oRsFarmTmp1.Fields!saldo = oRsPqte!Cantidad * lnSaldo99
                           oRsFarmTmp1.Fields!ingre = oRsPqte!Cantidad * lnIngre99
                           oRsFarmTmp1.Fields!reingre = oRsPqte!Cantidad * lnReingre99
                           oRsFarmTmp1.Fields!VENTA = oRsPqte!Cantidad * lnVenta99
                           oRsFarmTmp1.Fields!sis = oRsPqte!Cantidad * lnSis99
                           oRsFarmTmp1.Fields!intersan = oRsPqte!Cantidad * lnIntersan99
                           oRsFarmTmp1.Fields!DEFNAC = oRsPqte!Cantidad * lnDefNac99
                           oRsFarmTmp1.Fields!exo = oRsPqte!Cantidad * lnExo99
                           oRsFarmTmp1.Fields!soat = oRsPqte!Cantidad * lnSoat99
                           oRsFarmTmp1.Fields!credHosp = oRsPqte!Cantidad * lnCredHosp99
                           oRsFarmTmp1.Fields!otr_conv = oRsPqte!Cantidad * lnOtr_conv99
                           oRsFarmTmp1.Fields!DEVOL = oRsPqte!Cantidad * lnDevol99
                           oRsFarmTmp1.Fields!DEV_VEN = oRsPqte!Cantidad * lnDev_ven99
                           oRsFarmTmp1.Fields!DEV_MERMA = oRsPqte!Cantidad * lnDev_merma99
                           oRsFarmTmp1.Fields!otras_sal = oRsPqte!Cantidad * lnOtras_sal99
                        Else
                           oRsFarmTmp1.Fields!saldo = oRsFarmTmp1.Fields!saldo + (oRsPqte!Cantidad * lnSaldo99)
                           oRsFarmTmp1.Fields!ingre = oRsFarmTmp1.Fields!ingre + (oRsPqte!Cantidad * lnIngre99)
                           oRsFarmTmp1.Fields!reingre = oRsFarmTmp1.Fields!reingre + (oRsPqte!Cantidad * lnReingre99)
                           oRsFarmTmp1.Fields!VENTA = oRsFarmTmp1.Fields!VENTA + (oRsPqte!Cantidad * lnVenta99)
                           oRsFarmTmp1.Fields!sis = oRsFarmTmp1.Fields!sis + (oRsPqte!Cantidad * lnSis99)
                           oRsFarmTmp1.Fields!intersan = oRsFarmTmp1.Fields!intersan + (oRsPqte!Cantidad * lnIntersan99)
                           oRsFarmTmp1.Fields!DEFNAC = oRsFarmTmp1.Fields!DEFNAC + (oRsPqte!Cantidad * lnDefNac99)
                           oRsFarmTmp1.Fields!exo = oRsFarmTmp1.Fields!exo + (oRsPqte!Cantidad * lnExo99)
                           oRsFarmTmp1.Fields!soat = oRsFarmTmp1.Fields!soat + (oRsPqte!Cantidad * lnSoat99)
                           oRsFarmTmp1.Fields!credHosp = oRsFarmTmp1.Fields!credHosp + (oRsPqte!Cantidad * lnCredHosp99)
                           oRsFarmTmp1.Fields!otr_conv = oRsFarmTmp1.Fields!otr_conv + (oRsPqte!Cantidad * lnOtr_conv99)
                           oRsFarmTmp1.Fields!DEVOL = oRsFarmTmp1.Fields!DEVOL + (oRsPqte!Cantidad * lnDevol99)
                           oRsFarmTmp1.Fields!DEV_VEN = oRsFarmTmp1.Fields!DEV_VEN + (oRsPqte!Cantidad * lnDev_ven99)
                           oRsFarmTmp1.Fields!DEV_MERMA = oRsFarmTmp1.Fields!DEV_MERMA + (oRsPqte!Cantidad * lnDev_merma99)
                           oRsFarmTmp1.Fields!otras_sal = oRsFarmTmp1.Fields!otras_sal + (oRsPqte!Cantidad * lnOtras_sal99)
                        End If
                        oRsFarmTmp1.Update
                        oRsPqte.MoveNext
                     Loop
                  End If
               End If
               oRsFoxP.MoveNext
            Loop
            If oRsFarmTmp1.RecordCount > 0 Then
               oRsFarmTmp1.MoveFirst
               Do While Not oRsFarmTmp1.EOF
                  lnTotSalidas99 = oRsFarmTmp1!VENTA + oRsFarmTmp1!sis + oRsFarmTmp1!intersan + oRsFarmTmp1!DEFNAC + oRsFarmTmp1!exo + _
                                   oRsFarmTmp1!soat + oRsFarmTmp1!credHosp + oRsFarmTmp1!otr_conv + oRsFarmTmp1!DEVOL + _
                                   oRsFarmTmp1!DEV_VEN + oRsFarmTmp1!DEV_MERMA + oRsFarmTmp1!otras_sal
                  lnSaldoFinal99 = oRsFarmTmp1.Fields!saldo + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre - lnTotSalidas99
                  lcCodigo_med99 = Left(oRsFarmTmp1!codigo_med, 7)
                  lcCodMed991 = Trim(oRsFarmTmp1!codigo_med)
                  oRsFoxP.MoveFirst
                  oRsFoxP.Find "codigo_med='" & lcCodMed991 & "'"
                  
                  If oRsFoxP.EOF Then
                    '
                    If rsTmp_99.State = 1 Then rsTmp_99.Close
                    Set rsTmp_99 = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodMed991, 1)
                    lcNombre_99 = rsTmp_99!nombreProducto
                    lcTipoMI_99 = IIf(rsTmp_99!TipoProducto = 1, "Insumo", "Medicamento")
                    mrs_TmpP.AddNew
                    mrs_TmpP.Fields!codigo = lcCodMed991
                    mrs_TmpP.Fields!nombre = lcNombre_99
                    mrs_TmpP.Fields!precio = oRsFarmTmp1!precio
                    If mb_EsDonaciones = True Then
                        mrs_TmpP.Fields!saldoI = oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!otrasS = oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = oRsFarmTmp1!DEVOL + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                    Else
                        mrs_TmpP.Fields!saldoI = oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = 0   'LnDevolucionesP
                        mrs_TmpP.Fields!TotIngresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!Ventas = oRsFarmTmp1!VENTA
                        mrs_TmpP.Fields!sis = oRsFarmTmp1!sis
                        mrs_TmpP.Fields!soat = oRsFarmTmp1!soat
                        mrs_TmpP.Fields!convenio = oRsFarmTmp1!otr_conv
                        mrs_TmpP.Fields!creditoH = oRsFarmTmp1!credHosp
                        mrs_TmpP.Fields!defensaN = oRsFarmTmp1!DEFNAC
                        mrs_TmpP.Fields!OsDevol = oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!OsVencim = oRsFarmTmp1!DEV_VEN
                        mrs_TmpP.Fields!OsMerma = oRsFarmTmp1!DEV_MERMA
                        mrs_TmpP.Fields!Exonerac = oRsFarmTmp1!exo
                        mrs_TmpP.Fields!IntervencionS = oRsFarmTmp1!intersan
                        mrs_TmpP.Fields!otrasS = oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = lnTotSalidas99
                        mrs_TmpP.Fields!fechaVencimiento = ldFechaVencimiento
                    End If
                    mrs_TmpP.Fields!tipo = lcTipoMI_99
                    mrs_TmpP.Update
                    '
                    
                    oRsFoxP.AddNew
                    oRsFoxP.Fields!CODIGO_EJE = lcDisa
                    oRsFoxP.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFoxP.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFoxP.Fields!annoMes = lcAnioMes
                    oRsFoxP.Fields!codigo_med = lcCodigo_med99
                    oRsFoxP.Fields!saldo = oRsFarmTmp1!saldo
                    oRsFoxP.Fields!precio = oRsFarmTmp1!precio
                    oRsFoxP.Fields!ingre = oRsFarmTmp1!ingre
                    oRsFoxP.Fields!reingre = oRsFarmTmp1!reingre
                    oRsFoxP.Fields!VENTA = oRsFarmTmp1!VENTA
                    oRsFoxP.Fields!sis = oRsFarmTmp1!sis
                    oRsFoxP.Fields!intersan = oRsFarmTmp1!intersan
                    oRsFoxP.Fields!fac_perd = 0                        'falta
                    oRsFoxP.Fields!DEFNAC = oRsFarmTmp1!DEFNAC
                    oRsFoxP.Fields!exo = oRsFarmTmp1!exo
                    oRsFoxP.Fields!soat = oRsFarmTmp1!soat
                    oRsFoxP.Fields!credHosp = oRsFarmTmp1!credHosp
                    oRsFoxP.Fields!otr_conv = oRsFarmTmp1!otr_conv
                    oRsFoxP.Fields!DEVOL = oRsFarmTmp1!DEVOL
                    oRsFoxP.Fields!vencido = 0
                    oRsFoxP.Fields!merma = 0
                    oRsFoxP.Fields!distri = 0
                    oRsFoxP.Fields!transf = 0
                    oRsFoxP.Fields!ventaInst = 0
                    oRsFoxP.Fields!DEV_VEN = oRsFarmTmp1!DEV_VEN
                    oRsFoxP.Fields!DEV_MERMA = oRsFarmTmp1!DEV_MERMA
                    oRsFoxP.Fields!otras_sal = oRsFarmTmp1!otras_sal
                    oRsFoxP.Fields!STOCK_FIN = lnSaldoFinal99
                    oRsFoxP.Fields!stock_fin1 = lnSaldoFinal99
                    oRsFoxP.Fields!REQ = lnTotSalidas99
                    oRsFoxP.Fields!Total = lnTotSalidas99
                    If mb_EsDonaciones = False Then
                       oRsFoxP.Fields!FEC_EXP = ldFechaVencimiento
                    End If
                    If lnDonacionesSaldoI < 9999999 Then
                       oRsFoxP.Fields!do_saldo = lnDonacionesSaldoI
                    Else
                       oRsFoxP.Fields!do_saldo = 0
                    End If
                    oRsFoxP.Fields!do_ingre = 0
                    oRsFoxP.Fields!do_con = 0
                    oRsFoxP.Fields!do_otr = 0
                    oRsFoxP.Fields!do_tot = 0
                    If lnDonacionesSaldoI < 9999999 Then
                       oRsFoxP.Fields!do_stk = lnSaldoFinalD
                    Else
                       oRsFoxP.Fields!do_stk = 0
                    End If
                    If mb_EsDonaciones = True Then
                       oRsFoxP.Fields!do_fecExp = ldDonacionFechaVctoUlt
                    End If
                    oRsFoxP.Fields!fecha = Date
                    oRsFoxP.Fields!Usuario = " "
                    oRsFoxP.Fields!indiProc = " "
                    oRsFoxP.Fields!SIT = "1"
                    oRsFoxP.Fields!indiSiga = " "
                    oRsFoxP.Fields!dstkCero = 0
                    oRsFoxP.Fields!mptoRepo = 0
                    oRsFoxP.Update
                    '
                    ldFechaVencimiento = Date
                    lcLoteXitem = ""
                    For lnFor = 1 To Len(lc_AlmacenesParaICI)
                        If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                           lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                           lnFor = Len(lc_AlmacenesParaICI)
                        Else
                            lcTexto1 = ""
                            Do While True
                               If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                                  Exit Do
                               Else
                                  lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                                  lnFor = lnFor + 1
                               End If
                            Loop
                            lnIdAlmacenRep = Val(lcTexto1)
                        End If
                        
                        Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacenRep, 0, lcCodigo)
                        If rsTmp.RecordCount > 0 Then
                           If rsTmp.Fields!fechaVencimiento < ldFechaVencimiento Then
                                ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                                lcLoteXitem = rsTmp.Fields!Lote
                           End If
                        End If
                        rsTmp.Close
                    Next
                    'FormDetL
                    oRsFox1P.AddNew
                    oRsFox1P.Fields!CODIGO_EJE = lcDisa
                    oRsFox1P.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFox1P.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFox1P.Fields!annoMes = lcAnioMes
                    oRsFox1P.Fields!codigo_med = lcCodigo_med99
                    oRsFox1P.Fields!Lote = lcLoteXitem
                    oRsFox1P.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                    If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                       oRsFox1P.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                    Else
                       oRsFox1P.Fields!saldo = 0
                    End If
                    oRsFox1P.Fields!SIT = "1"
                    oRsFox1P.Update
                    'FormDetM
                    oRsFox2P.AddNew
                    oRsFox2P.Fields!CODIGO_EJE = lcDisa
                    oRsFox2P.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFox2P.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFox2P.Fields!annoMes = lcAnioMes
                    oRsFox2P.Fields!codigo_med = lcCodigo_med99
                    oRsFox2P.Fields!Lote = lcLoteXitem
                    oRsFox2P.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                    If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                       oRsFox2P.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                    Else
                       oRsFox2P.Fields!saldo = 0
                    End If
                    oRsFox2P.Fields!SIT = "1"
                    oRsFox2P.Update
                  Else
                    '********** ya existe el ITEM
                    mrs_TmpP.MoveFirst
                    mrs_TmpP.Find "codigo='" & lcCodMed991 & "'"
                    If Not mrs_TmpP.EOF Then
                    If mb_EsDonaciones = True Then
                        mrs_TmpP.Fields!saldoI = mrs_TmpP.Fields!saldoI + oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = mrs_TmpP.Fields!ingresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = mrs_TmpP.Fields!DevolucionesP + oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!otrasS = mrs_TmpP.Fields!otrasS + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = mrs_TmpP.Fields!TotSalidas + oRsFarmTmp1!DEVOL + oRsFarmTmp1!otras_sal
                    Else
                        mrs_TmpP.Fields!saldoI = mrs_TmpP.Fields!saldoI + oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = mrs_TmpP.Fields!ingresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = 0   'LnDevolucionesP
                        mrs_TmpP.Fields!TotIngresos = mrs_TmpP.Fields!TotIngresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!Ventas = mrs_TmpP.Fields!Ventas + oRsFarmTmp1!VENTA
                        mrs_TmpP.Fields!sis = mrs_TmpP.Fields!sis + oRsFarmTmp1!sis
                        mrs_TmpP.Fields!soat = mrs_TmpP.Fields!soat + oRsFarmTmp1!soat
                        mrs_TmpP.Fields!convenio = mrs_TmpP.Fields!convenio + oRsFarmTmp1!otr_conv
                        mrs_TmpP.Fields!creditoH = mrs_TmpP.Fields!creditoH + oRsFarmTmp1!credHosp
                        mrs_TmpP.Fields!defensaN = mrs_TmpP.Fields!defensaN + oRsFarmTmp1!DEFNAC
                        mrs_TmpP.Fields!OsDevol = mrs_TmpP.Fields!OsDevol + oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!OsVencim = mrs_TmpP.Fields!OsVencim + oRsFarmTmp1!DEV_VEN
                        mrs_TmpP.Fields!OsMerma = mrs_TmpP.Fields!OsMerma + oRsFarmTmp1!DEV_MERMA
                        mrs_TmpP.Fields!Exonerac = mrs_TmpP.Fields!Exonerac + oRsFarmTmp1!exo
                        mrs_TmpP.Fields!IntervencionS = mrs_TmpP.Fields!IntervencionS + oRsFarmTmp1!intersan
                        mrs_TmpP.Fields!otrasS = mrs_TmpP.Fields!otrasS + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = mrs_TmpP.Fields!TotSalidas + lnTotSalidas99
                    End If
                    mrs_TmpP.Update
                    End If
                    '
                    oRsFoxP.Fields!saldo = oRsFoxP.Fields!saldo + oRsFarmTmp1!saldo
                    oRsFoxP.Fields!ingre = oRsFoxP.Fields!ingre + oRsFarmTmp1!ingre
                    oRsFoxP.Fields!reingre = oRsFoxP.Fields!reingre + oRsFarmTmp1!reingre
                    oRsFoxP.Fields!VENTA = oRsFoxP.Fields!VENTA + oRsFarmTmp1!VENTA
                    oRsFoxP.Fields!sis = oRsFoxP.Fields!sis + oRsFarmTmp1!sis
                    oRsFoxP.Fields!intersan = oRsFoxP.Fields!intersan + oRsFarmTmp1!intersan
                    oRsFoxP.Fields!DEFNAC = oRsFoxP.Fields!DEFNAC + oRsFarmTmp1!DEFNAC
                    oRsFoxP.Fields!exo = oRsFoxP.Fields!exo + oRsFarmTmp1!exo
                    oRsFoxP.Fields!soat = oRsFoxP.Fields!soat + oRsFarmTmp1!exo
                    oRsFoxP.Fields!credHosp = oRsFoxP.Fields!credHosp + oRsFarmTmp1!credHosp
                    oRsFoxP.Fields!otr_conv = oRsFoxP.Fields!otr_conv + oRsFarmTmp1!otr_conv
                    oRsFoxP.Fields!DEVOL = oRsFoxP.Fields!DEVOL + oRsFarmTmp1!DEVOL
                    oRsFoxP.Fields!DEV_VEN = oRsFoxP.Fields!DEV_VEN + oRsFarmTmp1!DEV_VEN
                    oRsFoxP.Fields!DEV_MERMA = oRsFoxP.Fields!DEV_MERMA + oRsFarmTmp1!DEV_MERMA
                    oRsFoxP.Fields!otras_sal = oRsFoxP.Fields!otras_sal + oRsFarmTmp1!otras_sal
                    oRsFoxP.Fields!STOCK_FIN = oRsFoxP.Fields!STOCK_FIN + lnSaldoFinal99
                    oRsFoxP.Fields!stock_fin1 = oRsFoxP.Fields!stock_fin1 + lnSaldoFinal99
                    oRsFoxP.Fields!REQ = oRsFoxP.Fields!REQ + lnTotSalidas99
                    oRsFoxP.Fields!Total = oRsFoxP.Fields!Total + lnTotSalidas99
                    oRsFoxP.Update
                    'FormDetL
                    oRsFox1P.MoveFirst
                    oRsFox1P.Find "codigo_med='" & lcCodigo_med99 & "'"
                    If Not oRsFox1P.EOF Then
                        If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                           oRsFox1P.Fields!saldo = oRsFox1P.Fields!saldo + IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                        Else
                           oRsFox1P.Fields!saldo = 0
                        End If
                        oRsFox1P.Update
                    End If
                    'FormDetM
                    oRsFox2P.MoveFirst
                    oRsFox2P.Find "codigo_med='" & lcCodigo_med99 & "'"
                    If Not oRsFox2P.EOF Then
                        If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                           oRsFox2P.Fields!saldo = oRsFox2P.Fields!saldo + IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                        Else
                           oRsFox2P.Fields!saldo = 0
                        End If
                        oRsFox2P.Update
                     End If
                  End If
                  oRsFarmTmp1.MoveNext
               Loop
            End If
            'eliminando Paquetes
            mrs_TmpP.MoveFirst
            Do While Not mrs_TmpP.EOF
               lcCodigo_med99 = mrs_TmpP!codigo
               If mo_ReglasFarmacia.CatalogoDIGEMIDesCodigoPaquete(lcCodigo_med99, oConexion) = True Then
                  mrs_TmpP.Delete
                  mrs_TmpP.Update
               End If
               mrs_TmpP.MoveNext
            Loop
            '
            oRsFoxP.MoveFirst
            Do While Not oRsFoxP.EOF
               lcCodigo_med99 = oRsFoxP!codigo_med
               If mo_ReglasFarmacia.CatalogoDIGEMIDesCodigoPaquete(lcCodigo_med99, oConexion) = True Then
                  oRsFoxP.Delete
                  oRsFoxP.Update
                  oRsFox1P.Filter = "codigo_med='" & lcCodigo_med99 & "'"
                  If oRsFox1P.RecordCount > 0 Then
                     oRsFox1P.MoveFirst
                     Do While Not oRsFox1P.EOF
                        oRsFox1P.Delete
                        oRsFox1P.Update
                        oRsFox1P.MoveNext
                     Loop
                  End If
                  oRsFox2P.Filter = "codigo_med='" & lcCodigo_med99 & "'"
                  If oRsFox2P.RecordCount > 0 Then
                     oRsFox2P.MoveFirst
                     Do While Not oRsFox2P.EOF
                        oRsFox2P.Delete
                        oRsFox2P.Update
                        oRsFox2P.MoveNext
                     Loop
                  End If
                  
               End If
               oRsFoxP.MoveNext
            Loop
        End If
        oRsFox1P.Filter = ""
        oRsFox2P.Filter = ""
        oRsFoxP.Filter = ""
        Set oRsPqte = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set oRsFarmTmp1 = Nothing
        Set rsTmp99 = Nothing
        Set rsTmp_99 = Nothing

End Sub

'el CODIGO DEL ITEM es un Paquete de Farmacia,se desagrega en CODIGOS DIGEMID   'debb-02/12/2016
Sub CodigosUnidosisDesagregaEnMedicInsumos(ByRef oRsFoxP As Recordset, ByRef oRsFox1P As Recordset, ByRef oRsFox2P As Recordset, _
                                               lcDisa As String, mb_EsDonaciones As Boolean, lcAnioMes As String, _
                                               oConexion As Connection, ByRef mrs_TmpP As Recordset)
'Exit Sub
        On Error GoTo errCUnid
        Dim oRsPqte As New Recordset
        Dim mo_ReglasFarmacia As New ReglasFarmacia
        Dim oRsFarmTmp1 As New Recordset
        Dim rsTmp99 As New Recordset
        Dim lcCodigo As String, lnCantidadBolsasR As Long, lnPrecio As Double, lnCantidadR As Long, lnTotal As Double
        Dim lcTipo As String, lcDx As String, lnDxNro As Integer, lcFormaF As String, lnCantidadC As Long, lnSaldoInicial991 As Long
        Dim lnCantidadBolsasC As Long, lcLoteXitem As String, lcCodigoSinPunto As String, lnConvertir As Long
        Dim lnIngre99 As Long, lnReingre99 As Long, lnVenta99 As Long, lnSis99 As Long, lnIntersan99 As Long, lnDefNac99 As Long
        Dim lnExo99 As Long, lnSoat99 As Long, lcCodigo_med99 As String, lnTotSalidas99 As Long, lnSaldoFinal99 As Long
        Dim lnCredHosp99 As Long, lnOtr_conv99 As Long, lnDevol99 As Long, lnDev_ven99 As Long, lnDev_merma99 As Long, lnOtras_sal99 As Long
        Dim ldFechaVencimiento As Date, lnDonacionesSaldoI As Long, lnSaldoFinalD As Long, ldDonacionFechaVctoUlt As Date
        Dim rsTmp_99 As New Recordset, lcNombre_99 As String, lcTipoMI_99 As String, lcCodMed991 As String, lnSaldo99 As Long
        Dim oRsItemsUnidosis As New Recordset
        Set oRsItemsUnidosis = mo_ReglasFarmacia.farmUnidosisSeleccionarTodos
        ldFechaVencimiento = Date
        lnDonacionesSaldoI = 0
        lnSaldoFinalD = 0
        ldDonacionFechaVctoUlt = Date
        oRsFoxP.Filter = "sit='1' and codigo_pre='" & lc_CodigoSismed & "'"
        'oRsFoxP.Filter = "sit='9'"
        If oRsFoxP.RecordCount > 0 Then
            Set oRsFarmTmp1 = sighentidades.CopyRecordset(oRsFoxP, "")
            oRsFoxP.MoveFirst
            Do While Not oRsFoxP.EOF
               lcCodigo_med99 = Trim(oRsFoxP!codigo_med)
               If Right(lcCodigo_med99, 1) = sighentidades.Pto Then
                  lcCodigoSinPunto = Trim(Left(lcCodigo_med99, InStr(lcCodigo_med99, sighentidades.Pto) - 1))
                  oRsItemsUnidosis.MoveFirst
                  oRsItemsUnidosis.Find "codigo='" & Trim(lcCodigoSinPunto) & "'"
                  If Not oRsItemsUnidosis.EOF Then
                        If rsTmp_99.State = 1 Then rsTmp_99.Close
                        Set rsTmp_99 = mo_ReglasFacturacion.FacturacionBienesPorCodigo(lcCodigoSinPunto, 1, oConexion)
                        If rsTmp_99.RecordCount > 0 Then
                            lnConvertir = Val(oRsItemsUnidosis!convertir)
                            
                            If oRsFoxP!saldo = 0 Then
                                lnSaldo99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!saldo / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                   lnSaldo99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                   lnSaldo99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                   lnSaldo99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!ingre = 0 Then
                                lnIngre99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!ingre / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnIngre99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnIngre99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnIngre99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!reingre = 0 Then
                                lnReingre99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!reingre / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnReingre99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnReingre99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnReingre99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!VENTA = 0 Then
                                lnVenta99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!VENTA / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnVenta99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnVenta99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnVenta99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!sis = 0 Then
                                lnSis99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!sis / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnSis99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnSis99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnSis99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!intersan = 0 Then
                                lnIntersan99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!intersan / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnIntersan99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnIntersan99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnIntersan99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!DEFNAC = 0 Then
                                lnDefNac99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!DEFNAC / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnDefNac99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnDefNac99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnDefNac99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!exo = 0 Then
                                lnExo99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!exo / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnExo99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnExo99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnExo99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!soat = 0 Then
                                lnSoat99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!soat / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnSoat99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnSoat99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnSoat99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!credHosp = 0 Then
                                lnCredHosp99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!credHosp / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnCredHosp99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnCredHosp99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnCredHosp99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!otr_conv = 0 Then
                                lnOtr_conv99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!otr_conv / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnOtr_conv99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnOtr_conv99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnOtr_conv99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!DEVOL = 0 Then
                                lnDevol99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!DEVOL / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnDevol99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnDevol99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnDevol99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!DEV_VEN = 0 Then
                                lnDev_ven99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!DEV_VEN / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnDev_ven99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnDev_ven99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnDev_ven99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!DEV_MERMA = 0 Then
                                lnDev_merma99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!DEV_MERMA / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                    lnDev_merma99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                    lnDev_merma99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                    lnDev_merma99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            
                            If oRsFoxP!otras_sal = 0 Then
                                lnOtras_sal99 = 0
                            Else
                                lcTipoMI_99 = CCur(oRsFoxP!otras_sal / lnConvertir)
                                If InStr(lcTipoMI_99, ".") = 0 Then
                                   lnOtras_sal99 = Val(lcTipoMI_99)
                                ElseIf Val(Mid(lcTipoMI_99, InStr(lcTipoMI_99, ".") + 1, 10)) > 0 Then
                                   lnOtras_sal99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1)) + 1
                                Else
                                   lnOtras_sal99 = Val(Left(lcTipoMI_99, InStr(lcTipoMI_99, ".") - 1))
                                End If
                            End If
                            
                            oRsFarmTmp1.AddNew
                            oRsFarmTmp1.Fields!codigo_med = lcCodigoSinPunto
                            oRsFarmTmp1.Fields!precio = rsTmp_99!PrecioUnitario
                            oRsFarmTmp1.Fields!saldo = lnSaldo99
                            oRsFarmTmp1.Fields!ingre = lnIngre99
                            oRsFarmTmp1.Fields!reingre = lnReingre99
                            oRsFarmTmp1.Fields!VENTA = lnVenta99
                            oRsFarmTmp1.Fields!sis = lnSis99
                            oRsFarmTmp1.Fields!intersan = lnIntersan99
                            oRsFarmTmp1.Fields!DEFNAC = lnDefNac99
                            oRsFarmTmp1.Fields!exo = lnExo99
                            oRsFarmTmp1.Fields!soat = lnSoat99
                            oRsFarmTmp1.Fields!credHosp = lnCredHosp99
                            oRsFarmTmp1.Fields!otr_conv = lnOtr_conv99
                            oRsFarmTmp1.Fields!DEVOL = lnDevol99
                            oRsFarmTmp1.Fields!DEV_VEN = lnDev_ven99
                            oRsFarmTmp1.Fields!DEV_MERMA = lnDev_merma99
                            oRsFarmTmp1.Fields!otras_sal = lnOtras_sal99
                            oRsFarmTmp1.Update
                        End If
                   End If
               End If
               oRsFoxP.MoveNext
            Loop
            If oRsFarmTmp1.RecordCount > 0 Then
               oRsFarmTmp1.MoveFirst
               Do While Not oRsFarmTmp1.EOF
                  lnTotSalidas99 = oRsFarmTmp1!VENTA + oRsFarmTmp1!sis + oRsFarmTmp1!intersan + oRsFarmTmp1!DEFNAC + oRsFarmTmp1!exo + _
                                   oRsFarmTmp1!soat + oRsFarmTmp1!credHosp + oRsFarmTmp1!otr_conv + oRsFarmTmp1!DEVOL + _
                                   oRsFarmTmp1!DEV_VEN + oRsFarmTmp1!DEV_MERMA + oRsFarmTmp1!otras_sal
                  lnSaldoFinal99 = oRsFarmTmp1.Fields!saldo + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre - lnTotSalidas99
                  lcCodigo_med99 = Left(oRsFarmTmp1!codigo_med, 7)
                  lcCodMed991 = Trim(oRsFarmTmp1!codigo_med)
                  oRsFoxP.MoveFirst
                  oRsFoxP.Find "codigo_med='" & lcCodMed991 & "'"
                  
                  If oRsFoxP.EOF Then
                    If rsTmp_99.State = 1 Then rsTmp_99.Close
                    Set rsTmp_99 = mo_ReglasFacturacion.FacturacionBienesPorCodigoTipoFinanciamiento(oConexion, lcCodMed991, 1)
                    lcNombre_99 = rsTmp_99!nombreProducto
                    lcTipoMI_99 = IIf(rsTmp_99!TipoProducto = 1, "Insumo", "Medicamento")
                    '
                    lnSaldoInicial991 = 0
                    If lb_ConsiderarSaldoInicialDelHistorico = True Then     'debb-10/12/2018
                       If rsErrores.State = 1 Then rsErrores.Close
                       Set rsErrores = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes(lcCodMed991, lc_CodigoSismed, ldFechaHistoricoXmes, oConexion)
                       If rsErrores.RecordCount > 0 Then
                          If Not IsNull(rsErrores!STOCK_FIN) Then
                             lnSaldoInicial991 = rsErrores!STOCK_FIN
                          End If
                       End If
                       rsErrores.Close
                     Else
                        For lnFor = 1 To Len(lc_AlmacenesParaICI)
                            If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                               lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                               lnFor = Len(lc_AlmacenesParaICI)
                            Else
                                lcTexto1 = ""
                                Do While True
                                   If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                                      Exit Do
                                   Else
                                      lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                                      lnFor = lnFor + 1
                                   End If
                                Loop
                                lnIdAlmacenRep = Val(lcTexto1)
                            End If
                            If lnIdAlmacenRep > 3 Then
                                If rsErrores.State = 1 Then rsErrores.Close
                                Set rsErrores = mo_ReglasFarmacia.FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(rsTmp_99!idProducto, lnIdAlmacenRep, ldFechaHistoricoXmes, oConexion)
                                Do While Not rsErrores.EOF
                                    lnSaldoInicial991 = lnSaldoInicial991 + rsErrores.Fields!saldo
                                    rsErrores.MoveNext
                                Loop
                                rsErrores.Close
                            End If
                        Next
                    End If
                    lnSaldoFinal99 = lnSaldoFinal99 + lnSaldoInicial991
                    '
                    '
                    mrs_TmpP.AddNew
                    mrs_TmpP.Fields!codigo = lcCodMed991
                    mrs_TmpP.Fields!nombre = lcNombre_99
                    mrs_TmpP.Fields!precio = oRsFarmTmp1!precio
                    If mb_EsDonaciones = True Then
                        mrs_TmpP.Fields!saldoI = oRsFarmTmp1!saldo + lnSaldoInicial991
                        mrs_TmpP.Fields!ingresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!otrasS = oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = oRsFarmTmp1!DEVOL + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!fechaVencimiento = ldDonacionFechaVctoUlt
                    Else
                        mrs_TmpP.Fields!saldoI = oRsFarmTmp1!saldo + lnSaldoInicial991
                        mrs_TmpP.Fields!ingresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = 0   'LnDevolucionesP
                        mrs_TmpP.Fields!TotIngresos = oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!Ventas = oRsFarmTmp1!VENTA
                        mrs_TmpP.Fields!sis = oRsFarmTmp1!sis
                        mrs_TmpP.Fields!soat = oRsFarmTmp1!soat
                        mrs_TmpP.Fields!convenio = oRsFarmTmp1!otr_conv
                        mrs_TmpP.Fields!creditoH = oRsFarmTmp1!credHosp
                        mrs_TmpP.Fields!defensaN = oRsFarmTmp1!DEFNAC
                        mrs_TmpP.Fields!OsDevol = oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!OsVencim = oRsFarmTmp1!DEV_VEN
                        mrs_TmpP.Fields!OsMerma = oRsFarmTmp1!DEV_MERMA
                        mrs_TmpP.Fields!Exonerac = oRsFarmTmp1!exo
                        mrs_TmpP.Fields!IntervencionS = oRsFarmTmp1!intersan
                        mrs_TmpP.Fields!otrasS = oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = lnTotSalidas99
                        mrs_TmpP.Fields!fechaVencimiento = ldFechaVencimiento
                    End If
                    mrs_TmpP.Fields!tipo = lcTipoMI_99
                    mrs_TmpP.Update
                    '
                    
                    oRsFoxP.AddNew
                    oRsFoxP.Fields!CODIGO_EJE = lcDisa
                    oRsFoxP.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFoxP.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFoxP.Fields!annoMes = lcAnioMes
                    oRsFoxP.Fields!codigo_med = lcCodigo_med99
                    oRsFoxP.Fields!saldo = oRsFarmTmp1!saldo + lnSaldoInicial991
                    oRsFoxP.Fields!precio = oRsFarmTmp1!precio
                    oRsFoxP.Fields!ingre = oRsFarmTmp1!ingre
                    oRsFoxP.Fields!reingre = oRsFarmTmp1!reingre
                    oRsFoxP.Fields!VENTA = oRsFarmTmp1!VENTA
                    oRsFoxP.Fields!sis = oRsFarmTmp1!sis
                    oRsFoxP.Fields!intersan = oRsFarmTmp1!intersan
                    oRsFoxP.Fields!fac_perd = 0                        'falta
                    oRsFoxP.Fields!DEFNAC = oRsFarmTmp1!DEFNAC
                    oRsFoxP.Fields!exo = oRsFarmTmp1!exo
                    oRsFoxP.Fields!soat = oRsFarmTmp1!soat
                    oRsFoxP.Fields!credHosp = oRsFarmTmp1!credHosp
                    oRsFoxP.Fields!otr_conv = oRsFarmTmp1!otr_conv
                    oRsFoxP.Fields!DEVOL = oRsFarmTmp1!DEVOL
                    oRsFoxP.Fields!vencido = 0
                    oRsFoxP.Fields!merma = 0
                    oRsFoxP.Fields!distri = 0
                    oRsFoxP.Fields!transf = 0
                    oRsFoxP.Fields!ventaInst = 0
                    oRsFoxP.Fields!DEV_VEN = oRsFarmTmp1!DEV_VEN
                    oRsFoxP.Fields!DEV_MERMA = oRsFarmTmp1!DEV_MERMA
                    oRsFoxP.Fields!otras_sal = oRsFarmTmp1!otras_sal
                    oRsFoxP.Fields!STOCK_FIN = lnSaldoFinal99
                    oRsFoxP.Fields!stock_fin1 = lnSaldoFinal99
                    oRsFoxP.Fields!REQ = lnTotSalidas99
                    oRsFoxP.Fields!Total = lnTotSalidas99
                    If mb_EsDonaciones = False Then
                       oRsFoxP.Fields!FEC_EXP = ldFechaVencimiento
                    End If
                    If lnDonacionesSaldoI < 9999999 Then
                       oRsFoxP.Fields!do_saldo = lnDonacionesSaldoI
                    Else
                       oRsFoxP.Fields!do_saldo = 0
                    End If
                    oRsFoxP.Fields!do_ingre = 0
                    oRsFoxP.Fields!do_con = 0
                    oRsFoxP.Fields!do_otr = 0
                    oRsFoxP.Fields!do_tot = 0
                    If lnDonacionesSaldoI < 9999999 Then
                       oRsFoxP.Fields!do_stk = lnSaldoFinalD
                    Else
                       oRsFoxP.Fields!do_stk = 0
                    End If
                    If mb_EsDonaciones = True Then
                       oRsFoxP.Fields!do_fecExp = ldDonacionFechaVctoUlt
                    End If
                    oRsFoxP.Fields!fecha = Date
                    oRsFoxP.Fields!Usuario = " "
                    oRsFoxP.Fields!indiProc = " "
                    oRsFoxP.Fields!SIT = "1"
                    oRsFoxP.Fields!indiSiga = " "
                    oRsFoxP.Fields!dstkCero = 0
                    oRsFoxP.Fields!mptoRepo = 0
                    oRsFoxP.Update
                    '
                    ldFechaVencimiento = Date
                    lcLoteXitem = ""
                    For lnFor = 1 To Len(lc_AlmacenesParaICI)
                        If InStr(lc_AlmacenesParaICI, "/") = 0 Then
                           lnIdAlmacenRep = Val(lc_AlmacenesParaICI)
                           lnFor = Len(lc_AlmacenesParaICI)
                        Else
                            lcTexto1 = ""
                            Do While True
                               If Mid(lc_AlmacenesParaICI, lnFor, 1) = "/" Then
                                  Exit Do
                               Else
                                  lcTexto1 = lcTexto1 & Mid(lc_AlmacenesParaICI, lnFor, 1)
                                  lnFor = lnFor + 1
                               End If
                            Loop
                            lnIdAlmacenRep = Val(lcTexto1)
                        End If
                        
                        Set rsTmp = mo_ReglasFarmacia.FarmDevuelveSaldosConLotesSegunIdAlmacen(oConexion, lnIdAlmacenRep, 0, lcCodigo)
                        If rsTmp.RecordCount > 0 Then
                           If rsTmp.Fields!fechaVencimiento < ldFechaVencimiento Then
                                ldFechaVencimiento = rsTmp.Fields!fechaVencimiento
                                lcLoteXitem = rsTmp.Fields!Lote
                           End If
                        End If
                        rsTmp.Close
                    Next
                    'FormDetL
                    oRsFox1P.AddNew
                    oRsFox1P.Fields!CODIGO_EJE = lcDisa
                    oRsFox1P.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFox1P.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFox1P.Fields!annoMes = lcAnioMes
                    oRsFox1P.Fields!codigo_med = lcCodigo_med99
                    oRsFox1P.Fields!Lote = lcLoteXitem
                    oRsFox1P.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                    If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                       oRsFox1P.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                    Else
                       oRsFox1P.Fields!saldo = 0
                    End If
                    oRsFox1P.Fields!SIT = "1"
                    oRsFox1P.Update
                    'FormDetM
                    oRsFox2P.AddNew
                    oRsFox2P.Fields!CODIGO_EJE = lcDisa
                    oRsFox2P.Fields!CODIGO_PRE = lc_CodigoSismed
                    oRsFox2P.Fields!TIPSUM = IIf(mb_EsDonaciones = True, "D", "S")
                    oRsFox2P.Fields!annoMes = lcAnioMes
                    oRsFox2P.Fields!codigo_med = lcCodigo_med99
                    oRsFox2P.Fields!Lote = lcLoteXitem
                    oRsFox2P.Fields!fechVto = IIf(mb_EsDonaciones = True, ldDonacionFechaVctoUlt, ldFechaVencimiento)
                    If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                       oRsFox2P.Fields!saldo = IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                    Else
                       oRsFox2P.Fields!saldo = 0
                    End If
                    oRsFox2P.Fields!SIT = "1"
                    oRsFox2P.Update
                  Else
                    '********** ya existe el ITEM
                    mrs_TmpP.MoveFirst
                    mrs_TmpP.Find "codigo='" & lcCodMed991 & "'"
                    If Not mrs_TmpP.EOF Then
                    If mb_EsDonaciones = True Then
                        mrs_TmpP.Fields!saldoI = mrs_TmpP.Fields!saldoI + oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = mrs_TmpP.Fields!ingresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = mrs_TmpP.Fields!DevolucionesP + oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!otrasS = mrs_TmpP.Fields!otrasS + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = mrs_TmpP.Fields!TotSalidas + oRsFarmTmp1!DEVOL + oRsFarmTmp1!otras_sal
                    Else
                        mrs_TmpP.Fields!saldoI = mrs_TmpP.Fields!saldoI + oRsFarmTmp1!saldo
                        mrs_TmpP.Fields!ingresos = mrs_TmpP.Fields!ingresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!DevolucionesP = 0   'LnDevolucionesP
                        mrs_TmpP.Fields!TotIngresos = mrs_TmpP.Fields!TotIngresos + oRsFarmTmp1!ingre + oRsFarmTmp1!reingre
                        mrs_TmpP.Fields!Ventas = mrs_TmpP.Fields!Ventas + oRsFarmTmp1!VENTA
                        mrs_TmpP.Fields!sis = mrs_TmpP.Fields!sis + oRsFarmTmp1!sis
                        mrs_TmpP.Fields!soat = mrs_TmpP.Fields!soat + oRsFarmTmp1!soat
                        mrs_TmpP.Fields!convenio = mrs_TmpP.Fields!convenio + oRsFarmTmp1!otr_conv
                        mrs_TmpP.Fields!creditoH = mrs_TmpP.Fields!creditoH + oRsFarmTmp1!credHosp
                        mrs_TmpP.Fields!defensaN = mrs_TmpP.Fields!defensaN + oRsFarmTmp1!DEFNAC
                        mrs_TmpP.Fields!OsDevol = mrs_TmpP.Fields!OsDevol + oRsFarmTmp1!DEVOL
                        mrs_TmpP.Fields!OsVencim = mrs_TmpP.Fields!OsVencim + oRsFarmTmp1!DEV_VEN
                        mrs_TmpP.Fields!OsMerma = mrs_TmpP.Fields!OsMerma + oRsFarmTmp1!DEV_MERMA
                        mrs_TmpP.Fields!Exonerac = mrs_TmpP.Fields!Exonerac + oRsFarmTmp1!exo
                        mrs_TmpP.Fields!IntervencionS = mrs_TmpP.Fields!IntervencionS + oRsFarmTmp1!intersan
                        mrs_TmpP.Fields!otrasS = mrs_TmpP.Fields!otrasS + oRsFarmTmp1!otras_sal
                        mrs_TmpP.Fields!TotSalidas = mrs_TmpP.Fields!TotSalidas + lnTotSalidas99
                    End If
                    mrs_TmpP.Update
                    End If
                    '
                    oRsFoxP.Fields!saldo = oRsFoxP.Fields!saldo + oRsFarmTmp1!saldo
                    oRsFoxP.Fields!ingre = oRsFoxP.Fields!ingre + oRsFarmTmp1!ingre
                    oRsFoxP.Fields!reingre = oRsFoxP.Fields!reingre + oRsFarmTmp1!reingre
                    oRsFoxP.Fields!VENTA = oRsFoxP.Fields!VENTA + oRsFarmTmp1!VENTA
                    oRsFoxP.Fields!sis = oRsFoxP.Fields!sis + oRsFarmTmp1!sis
                    oRsFoxP.Fields!intersan = oRsFoxP.Fields!intersan + oRsFarmTmp1!intersan
                    oRsFoxP.Fields!DEFNAC = oRsFoxP.Fields!DEFNAC + oRsFarmTmp1!DEFNAC
                    oRsFoxP.Fields!exo = oRsFoxP.Fields!exo + oRsFarmTmp1!exo
                    oRsFoxP.Fields!soat = oRsFoxP.Fields!soat + oRsFarmTmp1!exo
                    oRsFoxP.Fields!credHosp = oRsFoxP.Fields!credHosp + oRsFarmTmp1!credHosp
                    oRsFoxP.Fields!otr_conv = oRsFoxP.Fields!otr_conv + oRsFarmTmp1!otr_conv
                    oRsFoxP.Fields!DEVOL = oRsFoxP.Fields!DEVOL + oRsFarmTmp1!DEVOL
                    oRsFoxP.Fields!DEV_VEN = oRsFoxP.Fields!DEV_VEN + oRsFarmTmp1!DEV_VEN
                    oRsFoxP.Fields!DEV_MERMA = oRsFoxP.Fields!DEV_MERMA + oRsFarmTmp1!DEV_MERMA
                    oRsFoxP.Fields!otras_sal = oRsFoxP.Fields!otras_sal + oRsFarmTmp1!otras_sal
                    oRsFoxP.Fields!STOCK_FIN = oRsFoxP.Fields!STOCK_FIN + lnSaldoFinal99
                    oRsFoxP.Fields!stock_fin1 = oRsFoxP.Fields!stock_fin1 + lnSaldoFinal99
                    oRsFoxP.Fields!REQ = oRsFoxP.Fields!REQ + lnTotSalidas99
                    oRsFoxP.Fields!Total = oRsFoxP.Fields!Total + lnTotSalidas99
                    oRsFoxP.Update
                    'FormDetL
                    oRsFox1P.MoveFirst
                    oRsFox1P.Find "codigo_med='" & lcCodigo_med99 & "'"
                    If Not oRsFox1P.EOF Then
                        If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                           oRsFox1P.Fields!saldo = oRsFox1P.Fields!saldo + IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                        Else
                           oRsFox1P.Fields!saldo = 0
                        End If
                        oRsFox1P.Update
                    End If
                    'FormDetM
                    oRsFox2P.MoveFirst
                    oRsFox2P.Find "codigo_med='" & lcCodigo_med99 & "'"
                    If Not oRsFox2P.EOF Then
                        If lnSaldoFinalD < 9999999 Or lnSaldoFinal99 < 9999999 Then
                           oRsFox2P.Fields!saldo = oRsFox2P.Fields!saldo + IIf(mb_EsDonaciones = True, lnSaldoFinalD, lnSaldoFinal99)
                        Else
                           oRsFox2P.Fields!saldo = 0
                        End If
                        oRsFox2P.Update
                     End If
                  End If
                  oRsFarmTmp1.MoveNext
               Loop
            End If
            'eliminando codigos con puntos
            mrs_TmpP.MoveFirst
            Do While Not mrs_TmpP.EOF
               lcCodigo_med99 = mrs_TmpP!codigo
               If Right(lcCodigo_med99, 1) = sighentidades.Pto Then
                  mrs_TmpP.Delete
                  mrs_TmpP.Update
               End If
               mrs_TmpP.MoveNext
            Loop
            '
            oRsFoxP.MoveFirst
            Do While Not oRsFoxP.EOF
               lcCodigo_med99 = oRsFoxP!codigo_med
               If Right(lcCodigo_med99, 1) = sighentidades.Pto Then
                  oRsFoxP.Delete
                  oRsFoxP.Update
                  oRsFox1P.Filter = "codigo_med='" & lcCodigo_med99 & "'"
                  If oRsFox1P.RecordCount > 0 Then
                     oRsFox1P.MoveFirst
                     Do While Not oRsFox1P.EOF
                        oRsFox1P.Delete
                        oRsFox1P.Update
                        oRsFox1P.MoveNext
                     Loop
                  End If
                  oRsFox2P.Filter = "codigo_med='" & lcCodigo_med99 & "'"
                  If oRsFox2P.RecordCount > 0 Then
                     oRsFox2P.MoveFirst
                     Do While Not oRsFox2P.EOF
                        oRsFox2P.Delete
                        oRsFox2P.Update
                        oRsFox2P.MoveNext
                     Loop
                  End If
                  
               End If
               oRsFoxP.MoveNext
            Loop
        End If
errCUnid:
        oRsFox1P.Filter = ""
        oRsFox2P.Filter = ""
        oRsFoxP.Filter = ""
        Set oRsPqte = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set oRsFarmTmp1 = Nothing
        Set rsTmp99 = Nothing
        Set rsTmp_99 = Nothing
        Set oRsItemsUnidosis = Nothing
        Exit Sub
        Resume
End Sub







