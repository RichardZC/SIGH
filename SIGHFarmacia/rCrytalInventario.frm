VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rCrytalInventario 
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8475
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
      DisplayGroupTree=   0   'False
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
Attribute VB_Name = "rCrytalInventario"
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
Dim lc_TextoDelFiltro As String, lcTexto10 As String
Dim mrs_Tmp As New Recordset
Dim mrs_tmp1 As New Recordset
Dim rsReporte As New ADODB.Recordset
Dim rsReporteAgrupado As New Recordset
Dim rsTmp As New Recordset
Dim mb_SolooConsolido As Boolean

'Dim ml_mesanio  As String 'MARIANO 07112014
Dim ml_mes  As Long 'MARIANO 07112014
Dim mb_Rreportes As String 'MARIANO 07112014

Dim mo_DoFarmMovimientoVentas As New DoFarmMovimientoVentas
Dim oFarmMovimientoDetalle As New farmMovimientoDetalle
'm
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_HoraInicio As String
Dim ml_HoraFin As String
Dim lnIdAlmacenOrigen As Long
Dim lnIdAlmacenDestino As Long
Dim ml_Proveedor As Long

Dim lnIdAlmacen As Long
Dim lnOrdenadoPor As Long: Dim lnIdProducto As Long

Dim mb_ConsiderarSinMovimientos As Boolean
Dim mb_SeMuestraLotes As Boolean
Dim mb_StockMinimoMayorAcantidad As Boolean
Dim ml_idUsuario As Long
Dim ml_idProducto  As Long



Dim ml_IdConcepto As Long
Dim ml_MovTipo As String
Dim ml_IdEstado  As Long
Dim lc_AlmacenesParaICI As String
Dim ml_IdAnio As Long
Dim ml_IdCuenta As Long
Dim ml_Dias  As Long

Dim ml_Almacen As String
Dim ml_AlmacenO As String

Dim ml_Documento As String

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
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
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
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String
Dim lc_TipoReporte As String, lcCodigo As String, lcNombre As String

Property Let CodigoSismed(lValue As String)
    lc_CodigoSismed = lValue
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


'MARIANO 07112014
Property Let Rreportes(lValue As String)
    mb_Rreportes = lValue
End Property
Property Let SoloConsolidado(lValue As Boolean)
    mb_SolooConsolido = lValue
End Property
'Property Let Mesanio(lValue As String)
 '   ml_mesanio = lValue
'End Property
Property Let Mes(lValue As Long)
    ml_mes = lValue
End Property
'mariano 20112014
Property Let IdProveedores(lValue As Long)
    ml_Proveedor = lValue
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
'm
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

Private Sub Form_Activate()
    If Len(lc_TextoDelFiltro) > 250 Then
       lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
    End If

Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition
Dim lcAnioMes As String, lnErrorEnOdbc As Integer, lbContinuar As Boolean, lnConsumoFarmacia1 As Double
Dim lcSerieB As String, lcDocumentoB As String, lnRedondeoB As Double, lbTienePagoAcuenta As Boolean
Dim lnTotalBol As Double, lnTotalExo As Double, lnTotalAde As Double, lbEsNuevoDocumento As Boolean
On Error GoTo ErrHandler
lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)

Screen.MousePointer = vbHourglass
Select Case lc_TipoReporte
Case "Inventario"
    'AGREGADO POR MARIANO 11112014
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set rsReporte = mo_ReglasFarmacia.FarmInventarioSeleccionarXdocumentoAlmacen(ml_Documento, lnIdAlmacenDestino, oConexion)
        If mb_Rreportes = "InventarioC" Then  'reporte por conteo
                        'Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                        GenerarRecordsetTemporalInventario
                        Do While Not rsReporte.EOF
                           mrs_Tmp.AddNew
                           mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                           mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                           mrs_Tmp.Fields!registroSanitario = rsReporte.Fields!registroSanitario
                           mrs_Tmp.Fields!FormaFarmaceutica = rsReporte.Fields!FormaFarmaceutica
                           mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                           mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
                           mrs_Tmp.Update
                           'Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                           rsReporte.MoveNext
                        Loop
                         mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\farmInventarioConteo.rpt", 1)
                         ' Parametros del reporte
                         Set crParamDefs = crReport.ParameterFields
                         For Each crParamDef In crParamDefs
                             Select Case crParamDef.ParameterFieldName
                                Case "Pml_Documento"
                                    crParamDef.AddCurrentValue ("Inventario de Productos N° " & ml_Documento & " (Para Conteo)")
                                Case "PAlmacenDestino"
                                    crParamDef.AddCurrentValue (lnIdAlmacenDestino)
                                Case "lcAlmacenDestino"
                                    crParamDef.AddCurrentValue (ml_Almacen)
                                Case "lcFechaNI"
                                    crParamDef.AddCurrentValue (ml_HoraInicio)
                                Case "cshos"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(205))
                                Case "red"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(240))
                                Case "microred"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(241))
                                Case "lcEESS"
                                    crParamDef.AddCurrentValue (lcTitEESS)
                                Case "lcEESSdireccion"
                                    crParamDef.AddCurrentValue (lcTitDireccion)
                                Case "lcEESStelefono"
                                    crParamDef.AddCurrentValue (lcTitTelefono)
                             End Select
                         Next
                         crReport.Database.SetDataSource mrs_Tmp
         'End If
         ElseIf mb_Rreportes = "InventarioD" Then 'Reporte Detallado
                        GenerarRecordsetTemporalInventario
                        Do While Not rsReporte.EOF
                           mrs_Tmp.AddNew
                           mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                           mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                           mrs_Tmp.Fields!registroSanitario = rsReporte.Fields!registroSanitario
                           mrs_Tmp.Fields!FormaFarmaceutica = rsReporte.Fields!FormaFarmaceutica
                           mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                           mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
                           mrs_Tmp.Fields!FechaVencimiento = Format(rsReporte.Fields!FechaVencimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                           mrs_Tmp.Fields!Precio = rsReporte.Fields!Precio
                           mrs_Tmp.Update
                           rsReporte.MoveNext
                        Loop
                         mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\farmInventarioDetallado.rpt", 1)
                        Set crParamDefs = crReport.ParameterFields
                         For Each crParamDef In crParamDefs
                            Select Case crParamDef.ParameterFieldName
                                Case "Pml_Documento"
                                    crParamDef.AddCurrentValue ("Inventario de Productos N° " & ml_Documento & " (Detallado)")
                                Case "PAlmacenDestino"
                                    crParamDef.AddCurrentValue (lnIdAlmacenDestino)
                                Case "lcAlmacenDestino"
                                    crParamDef.AddCurrentValue (ml_Almacen)
                                Case "lcFechaNI"
                                    crParamDef.AddCurrentValue (ml_HoraInicio)
                                Case "cshos"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(205))
                                Case "red"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(240))
                                Case "microred"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(241))
                                Case "lcEESS"
                                    crParamDef.AddCurrentValue (lcTitEESS)
                                Case "lcEESSdireccion"
                                    crParamDef.AddCurrentValue (lcTitDireccion)
                                Case "lcEESStelefono"
                                    crParamDef.AddCurrentValue (lcTitTelefono)
                             End Select
                         Next
                         crReport.Database.SetDataSource mrs_Tmp
        'End If
    ElseIf mb_Rreportes = "InventarioG" Then 'Reporte General
                        GenerarRecordsetTemporalInventario
                        Do While Not rsReporte.EOF
                           mrs_Tmp.AddNew
                           mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                           mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                           mrs_Tmp.Fields!FormaFarmaceutica = rsReporte.Fields!FormaFarmaceutica
                           mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
                           mrs_Tmp.Fields!Precio = rsReporte.Fields!Precio
                           mrs_Tmp.Fields!CantidadSaldo = rsReporte.Fields!CantidadSaldo
                           mrs_Tmp.Fields!Totalactual = rsReporte.Fields!CantidadSaldo * rsReporte.Fields!Precio
                           mrs_Tmp.Fields!CantidadFaltante = rsReporte.Fields!CantidadFaltante
                           mrs_Tmp.Fields!totalf = rsReporte.Fields!CantidadFaltante * rsReporte.Fields!Precio
                           mrs_Tmp.Fields!CantidadSobrante = rsReporte.Fields!CantidadSobrante
                           mrs_Tmp.Fields!totals = rsReporte.Fields!CantidadSobrante * rsReporte.Fields!Precio
                           mrs_Tmp.Fields!totalgen = rsReporte.Fields!Cantidad * rsReporte.Fields!Precio
                           mrs_Tmp.Update
                           rsReporte.MoveNext
                        Loop
                         mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\farmInventarioGeneral.rpt", 1)
                        Set crParamDefs = crReport.ParameterFields
                         For Each crParamDef In crParamDefs
                             Select Case crParamDef.ParameterFieldName
                                Case "Pml_Documento"
                                     crParamDef.AddCurrentValue ("Inventario de Productos N° " & ml_Documento & " (General)")
                                Case "PAlmacenDestino"
                                     crParamDef.AddCurrentValue (lnIdAlmacenDestino)
                                Case "lcAlmacenDestino"
                                        crParamDef.AddCurrentValue (ml_Almacen)
                                Case "lcFechaNI"
                                    crParamDef.AddCurrentValue (ml_HoraInicio)
                                Case "cshos"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(205))
                                Case "red"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(240))
                                Case "microred"
                                     crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(241))
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
'''.......................................................
    Case "rConsumoPorCuenta"
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        'Filtra los Datos
        'Editado Mariano 07112014
                        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuentaXmes(ml_IdCuenta, oConexion)
                        rsReporte.Filter = "idFuenteFinanciamiento<>1"
                        lnTotalRegistros = rsReporte.RecordCount
                        If lnTotalRegistros = 0 Then
                             MsgBox "No hay consumos para esa Cuenta", vbInformation, "Consumo por Cuenta"
                        Else
                '            Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                            GenerarRecordsetTemporalConsumoCUENTAXmes
                            rsReporte.MoveFirst
                            lnPrecio = 0
                            Do While Not rsReporte.EOF
                            lcTexto1 = ""
                            Set mrs_tmp1 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(rsReporte.Fields!movNumero, "S", rsReporte.Fields!idProducto, oConexion)
                            If mrs_tmp1.RecordCount > 0 Then
                              If mrs_tmp1.Fields!idEstadoFacturacion = 4 And mrs_tmp1.Fields!IdComprobantePago > 0 Then
                                 lcTexto1 = "Pago"
                                End If
                            End If
                                mrs_Tmp.AddNew
                                'mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                                mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                                mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                                mrs_Tmp.Fields!Precio = Round(rsReporte.Fields!Precio, 2)
                                'saldo anterior
                                'mrs_Tmp.Fields!saldoanterior = IIf(Month(rsReporte.Fields!fechaCreacion) < Val(ml_mes), rsReporte.Fields!cantidad, "0")
                                If Year(rsReporte.Fields!fechaCreacion) < Val(ml_IdAnio) Then
                                        mrs_Tmp.Fields!saldoanterior = IIf(IsNull(rsReporte.Fields!Cantidad), "0", rsReporte.Fields!Cantidad)
                                ElseIf Year(rsReporte.Fields!fechaCreacion) > Val(ml_IdAnio) Then
                                        mrs_Tmp.Fields!saldoposterior = IIf(IsNull(rsReporte.Fields!Cantidad), "0", rsReporte.Fields!Cantidad)
                                Else
                                        mrs_Tmp.Fields!saldoanterior = IIf(Month(rsReporte.Fields!fechaCreacion) < Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                        mrs_Tmp.Fields!saldoposterior = IIf(Month(rsReporte.Fields!fechaCreacion) > Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                End If
                                mrs_Tmp.Fields!uno = IIf(Day(rsReporte.Fields!fechaCreacion) = "1" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dos = IIf(Day(rsReporte.Fields!fechaCreacion) = "2" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!tres = IIf(Day(rsReporte.Fields!fechaCreacion) = "3" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!cuatro = IIf(Day(rsReporte.Fields!fechaCreacion) = "4" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!cinco = IIf(Day(rsReporte.Fields!fechaCreacion) = "5" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!seis = IIf(Day(rsReporte.Fields!fechaCreacion) = "6" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!siete = IIf(Day(rsReporte.Fields!fechaCreacion) = "7" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!ocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "8" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!nueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "9" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diez = IIf(Day(rsReporte.Fields!fechaCreacion) = "10" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!once = IIf(Day(rsReporte.Fields!fechaCreacion) = "11" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!doce = IIf(Day(rsReporte.Fields!fechaCreacion) = "12" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!trece = IIf(Day(rsReporte.Fields!fechaCreacion) = "13" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!catorce = IIf(Day(rsReporte.Fields!fechaCreacion) = "14" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!quince = IIf(Day(rsReporte.Fields!fechaCreacion) = "15" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dieciseis = IIf(Day(rsReporte.Fields!fechaCreacion) = "16" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diecisiete = IIf(Day(rsReporte.Fields!fechaCreacion) = "17" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dieciocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "18" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diecinueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "19" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinte = IIf(Day(rsReporte.Fields!fechaCreacion) = "20" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinteuno = IIf(Day(rsReporte.Fields!fechaCreacion) = "21" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintedos = IIf(Day(rsReporte.Fields!fechaCreacion) = "22" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintetres = IIf(Day(rsReporte.Fields!fechaCreacion) = "23" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintecuatro = IIf(Day(rsReporte.Fields!fechaCreacion) = "24" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintecinco = IIf(Day(rsReporte.Fields!fechaCreacion) = "25" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintesies = IIf(Day(rsReporte.Fields!fechaCreacion) = "26" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintesiete = IIf(Day(rsReporte.Fields!fechaCreacion) = "27" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinteocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "28" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintenueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "29" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!treinta = IIf(Day(rsReporte.Fields!fechaCreacion) = "30" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!treintayuno = IIf(Day(rsReporte.Fields!fechaCreacion) = "31" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                'mrs_Tmp.Fields!saldoposterior = IIf(Month(rsReporte.Fields!fechaCreacion) > Val(ml_mes), rsReporte.Fields!cantidad, "0")
                    
                                mrs_Tmp.Update
                                If lcTexto1 = "" Then
                                    lnPrecio = lnPrecio + rsReporte.Fields!total
                                End If
                               'Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                               rsReporte.MoveNext
                            Loop
                            
                           ' Devoluciones
                        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoNotaIngresoSeleccionarPorCuenta(ml_IdCuenta)
                        rsReporte.Filter = "idTipoConcepto=21"    'solo DEVOLUCIONES del PACIENTE
                        lnTotalRegistros = rsReporte.RecordCount
                        
                        If lnTotalRegistros > 0 Then
                           'Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
                           Do While Not rsReporte.EOF
                                mrs_Tmp.AddNew
                                'mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                                mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                                mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                                mrs_Tmp.Fields!Precio = Round(rsReporte.Fields!Precio, 2)
                                'mrs_Tmp.Fields!Precio = rsReporte.Fields!total
                                'saldo anterior
                                'mrs_Tmp.Fields!saldoanterior = IIf(Month(rsReporte.Fields!fechaCreacion) < Val(ml_mes), rsReporte.Fields!cantidad, "0")

                                mrs_Tmp.Fields!uno = IIf(Day(rsReporte.Fields!fechaCreacion) = "1" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dos = IIf(Day(rsReporte.Fields!fechaCreacion) = "2" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!tres = IIf(Day(rsReporte.Fields!fechaCreacion) = "3" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!cuatro = IIf(Day(rsReporte.Fields!fechaCreacion) = "4" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!cinco = IIf(Day(rsReporte.Fields!fechaCreacion) = "5" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!seis = IIf(Day(rsReporte.Fields!fechaCreacion) = "6" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!siete = IIf(Day(rsReporte.Fields!fechaCreacion) = "7" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!ocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "8" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!nueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "9" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diez = IIf(Day(rsReporte.Fields!fechaCreacion) = "10" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!once = IIf(Day(rsReporte.Fields!fechaCreacion) = "11" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!doce = IIf(Day(rsReporte.Fields!fechaCreacion) = "12" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!trece = IIf(Day(rsReporte.Fields!fechaCreacion) = "13" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!catorce = IIf(Day(rsReporte.Fields!fechaCreacion) = "14" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!quince = IIf(Day(rsReporte.Fields!fechaCreacion) = "15" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dieciseis = IIf(Day(rsReporte.Fields!fechaCreacion) = "16" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diecisiete = IIf(Day(rsReporte.Fields!fechaCreacion) = "17" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!dieciocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "18" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!diecinueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "19" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinte = IIf(Day(rsReporte.Fields!fechaCreacion) = "20" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinteuno = IIf(Day(rsReporte.Fields!fechaCreacion) = "21" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintedos = IIf(Day(rsReporte.Fields!fechaCreacion) = "22" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintetres = IIf(Day(rsReporte.Fields!fechaCreacion) = "23" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintecuatro = IIf(Day(rsReporte.Fields!fechaCreacion) = "24" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintecinco = IIf(Day(rsReporte.Fields!fechaCreacion) = "25" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintesies = IIf(Day(rsReporte.Fields!fechaCreacion) = "26" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintesiete = IIf(Day(rsReporte.Fields!fechaCreacion) = "27" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veinteocho = IIf(Day(rsReporte.Fields!fechaCreacion) = "28" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!veintenueve = IIf(Day(rsReporte.Fields!fechaCreacion) = "29" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!treinta = IIf(Day(rsReporte.Fields!fechaCreacion) = "30" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                mrs_Tmp.Fields!treintayuno = IIf(Day(rsReporte.Fields!fechaCreacion) = "31" And Month(rsReporte.Fields!fechaCreacion) = Val(ml_mes), -rsReporte.Fields!Cantidad, "0")
                                If Year(rsReporte.Fields!fechaCreacion) < Val(ml_IdAnio) Then
                                        mrs_Tmp.Fields!saldoanterior = IIf(IsNull(rsReporte.Fields!Cantidad), "0", rsReporte.Fields!Cantidad)
                                ElseIf Year(rsReporte.Fields!fechaCreacion) > Val(ml_IdAnio) Then
                                        mrs_Tmp.Fields!saldoposterior = IIf(IsNull(rsReporte.Fields!Cantidad), "0", rsReporte.Fields!Cantidad)
                                Else
                                        mrs_Tmp.Fields!saldoanterior = IIf(Month(rsReporte.Fields!fechaCreacion) < Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                        mrs_Tmp.Fields!saldoposterior = IIf(Month(rsReporte.Fields!fechaCreacion) > Val(ml_mes), rsReporte.Fields!Cantidad, "0")
                                End If

                                mrs_Tmp.Update
                                If rsReporte.Fields!idEstadoMovimiento = 1 Then
                                   lnPrecio = lnPrecio - rsReporte.Fields!total
                                End If
                                'Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                                rsReporte.MoveNext
                            Loop
                        End If
                            
                             'Reporte
                             mflgContinuar = True
                             Set crReport = crApp.OpenReport(App.Path & "\plantillas\farmConsumoPorCuentaXmes.rpt", 1)
                             ' Parametros del reporte
                             Set crParamDefs = crReport.ParameterFields
                             For Each crParamDef In crParamDefs
                                 Select Case crParamDef.ParameterFieldName
                                     Case "subTitulo"
                                         crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                                     Case "total"
                                         crParamDef.AddCurrentValue (lnPrecio)
                                     Case "MeSelec"
                                        crParamDef.AddCurrentValue (MonthName(ml_mes) & " " & Val(ml_IdAnio)) 'ml_mesanio
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
    Case "rProductosIngresados"
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        'Filtra los Datos
                        If ml_Proveedor = "0" Then 'sin proveedor
                            Set rsReporte = mo_ReglasFarmacia.FarmMovimientoSeleccionPorAlmacenProductosIngresados(mda_FechaInicio, mda_FechaFin, lnIdAlmacenDestino, lnIdAlmacenOrigen, oConexion)
                        Else 'con proveedor
                            Set rsReporte = mo_ReglasFarmacia.FarmMovimientoSeleccionPorProveedorProductosIngresados(mda_FechaInicio, mda_FechaFin, lnIdAlmacenDestino, lnIdAlmacenOrigen, ml_Proveedor, oConexion)
                        End If
                            GenerarRecordsetTemporalProductosIngresados
                        Do While Not rsReporte.EOF
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                           'mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
                            mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                            mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                            mrs_Tmp.Fields!preciou = rsReporte.Fields!PrecioUnitario
                            mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
                            mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                            mrs_Tmp.Fields!Concepto = rsReporte.Fields!Concepto
                            mrs_Tmp.Fields!FechaVencimiento = Format(rsReporte.Fields!FechaVencimiento, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                            mrs_Tmp.Fields!monto = rsReporte.Fields!monto
                            
                            mrs_Tmp.Update
                           'Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: DoEvents
                           rsReporte.MoveNext
                        Loop
                         'Reporte
                         'mflgContinuar = True
                         If mb_Rreportes = "ProductosIngConso" Then
                            Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmProductosIngresadosConsolidado.rpt", 1)
                         ElseIf mb_Rreportes = "ProductosIngDet" Then
                            Set crReport = crApp.OpenReport(App.Path & "\plantillas\FarmProductosIngresadosDetallado.rpt", 1)
                         End If
                         ' Parametros del reporte
                         Set crParamDefs = crReport.ParameterFields
                         For Each crParamDef In crParamDefs
                             Select Case crParamDef.ParameterFieldName
                            Case "IdAlmorigen"
                                  crParamDef.AddCurrentValue (ml_AlmacenO) 'AlmacenO
                            Case "IdAlmDestino"
                                  crParamDef.AddCurrentValue (ml_Almacen) 'Almacen
                            Case "RangoFec"
                                  crParamDef.AddCurrentValue ("DEL " & mda_FechaInicio & " Al " & mda_FechaFin)
                            Case "cshos"
                                  crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(205))
                            Case "red"
                                  crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(240))
                            Case "microred"
                                  crParamDef.AddCurrentValue (lcBuscaParametro.SeleccionaFilaParametro(241))
                            Case "lcEESS"
                                crParamDef.AddCurrentValue (lcTitEESS)
                            Case "lcEESSdireccion"
                                crParamDef.AddCurrentValue (lcTitDireccion)
                            Case "lcEESStelefono"
                                crParamDef.AddCurrentValue (lcTitTelefono)
                        End Select
                         Next
                         crReport.Database.SetDataSource mrs_Tmp
                    'End If
        oConexion.Close

End Select



'
'        '
'        If mb_EnArchivoExcel = True Then
'            Select Case lc_TipoReporte
'            Case "rConsumoPorCuenta"
'                mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Consumo por Cuenta por Mes", lc_TextoDelFiltro, "", Me.hwnd
'            End Select
'        End If
'        '



        CrvReportes.ReportSource = crReport
        CrvReportes.ViewReport
        CrvReportes.Zoom 120
'
        mo_ReglasComunes.grabaTablaAuditoria (crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))   'debb-27/05/2015
                
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    LimpiarVariablesDeMemoria
    Exit Sub
ErrHandler:
    If lnErrorEnOdbc = 0 Then
       Resume Next
    ElseIf Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
    Resume
    Screen.MousePointer = vbDefault
End Sub
Sub GenerarRecordsetTemporalInventario()
    With mrs_Tmp
        .Fields.Append "movNumero", adVarChar, 9, adFldIsNullable
        .Fields.Append "movTipo", adVarChar, 1, adFldIsNullable
        .Fields.Append "idproducto", adInteger, 4, adFldIsNullable
        .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
        .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
        .Fields.Append "item", adInteger, 4, adFldIsNullable
        .Fields.Append "Cantidad", adInteger, 4, adFldIsNullable
        .Fields.Append "Precio", adDouble
        .Fields.Append "total", adDouble
        .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
        .Fields.Append "codigo", adVarChar, 10, adFldIsNullable
        .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
        .Fields.Append "Presentacion", adVarChar, 100, adFldIsNullable
        .Fields.Append "idAlmcen", adInteger, 4, adFldIsNullable
        .Fields.Append "Totalactual", adDouble
        .Fields.Append "CantidadSaldo", adDouble
        .Fields.Append "CantidadFaltante", adDouble
        .Fields.Append "CantidadSobrante", adDouble
        .Fields.Append "totalgen", adDouble
        .Fields.Append "totalf", adDouble
        .Fields.Append "totals", adDouble
        .Fields.Append "FormaFarmaceutica", adVarChar, 10, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Sub GenerarRecordsetTemporalConsumoCUENTA()
    With mrs_Tmp
          .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
          '.Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
          .Fields.Append "MovNumero", adVarChar, 15, adFldIsNullable
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "cantidad", adInteger, 4, adFldIsNullable
          .Fields.Append "saldoa", adInteger, 4, adFldIsNullable
          .Fields.Append "saldop", adInteger, 4, adFldIsNullable
          .Fields.Append "totalgen", adInteger, 4, adFldIsNullable
          .Fields.Append "d1", adInteger, 4, adFldIsNullable
          .Fields.Append "d2", adInteger, 4, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "total", adDouble
          .Fields.Append "cantidadv", adInteger, 4, adFldIsNullable
          .Fields.Append "cantidadd", adInteger, 4, adFldIsNullable
          .Fields.Append "mes", adVarChar, 10, adFldIsNullable
          .Fields.Append "dia", adVarChar, 10, adFldIsNullable
          
          '.Fields.Append "Estado", adVarChar, 20, adFldIsNullable
          '.Fields.Append "dAlmacen", adVarChar, 100, adFldIsNullable
          '.Fields.Append "dFinanciamiento", adVarChar, 50, adFldIsNullable
          '.Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
          .Fields.Append "Tipo", adVarChar, 30, adFldIsNullable
          '.Fields.Append "saldoanterior", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

'MARIANO 07112014
Sub GenerarRecordsetTemporalConsumoCUENTAXmes()
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
            'agregado x Mariano 07112014
            .Fields.Append "uno", adInteger, 4, adFldIsNullable
            .Fields.Append "dos", adInteger, 4, adFldIsNullable
            .Fields.Append "tres", adInteger, 4, adFldIsNullable
            .Fields.Append "cuatro", adInteger, 4, adFldIsNullable
            .Fields.Append "cinco", adInteger, 4, adFldIsNullable
            .Fields.Append "seis", adInteger, 4, adFldIsNullable
            .Fields.Append "siete", adInteger, 4, adFldIsNullable
            .Fields.Append "ocho", adInteger, 4, adFldIsNullable
            .Fields.Append "nueve", adInteger, 4, adFldIsNullable
            .Fields.Append "diez", adInteger, 4, adFldIsNullable
            .Fields.Append "once", adInteger, 4, adFldIsNullable
            .Fields.Append "doce", adInteger, 4, adFldIsNullable
            .Fields.Append "trece", adInteger, 4, adFldIsNullable
            .Fields.Append "catorce", adInteger, 4, adFldIsNullable
            .Fields.Append "quince", adInteger, 4, adFldIsNullable
            .Fields.Append "dieciseis", adInteger, 4, adFldIsNullable
            .Fields.Append "diecisiete", adInteger, 4, adFldIsNullable
            .Fields.Append "dieciocho", adInteger, 4, adFldIsNullable
            .Fields.Append "diecinueve", adInteger, 4, adFldIsNullable
            .Fields.Append "veinte", adInteger, 4, adFldIsNullable
            .Fields.Append "veinteuno", adInteger, 4, adFldIsNullable
            .Fields.Append "veintedos", adInteger, 4, adFldIsNullable
            .Fields.Append "veintetres", adInteger, 4, adFldIsNullable
            .Fields.Append "veintecuatro", adInteger, 4, adFldIsNullable
            .Fields.Append "veintecinco", adInteger, 4, adFldIsNullable
            .Fields.Append "veintesies", adInteger, 4, adFldIsNullable
            .Fields.Append "veintesiete", adInteger, 4, adFldIsNullable
            .Fields.Append "veinteocho", adInteger, 4, adFldIsNullable
            .Fields.Append "veintenueve", adInteger, 4, adFldIsNullable
            .Fields.Append "treinta", adInteger, 4, adFldIsNullable
            .Fields.Append "treintayuno", adInteger, 4, adFldIsNullable
            .Fields.Append "totalconsumo", adDouble
            .Fields.Append "saldoanterior", adInteger, 4, adFldIsNullable
            .Fields.Append "saldoposterior", adInteger, 4, adFldIsNullable
            .Fields.Append "totalmes", adInteger, 4, adFldIsNullable
            .Fields.Append "devol", adInteger, 4, adFldIsNullable
            .Fields.Append "mes", adDate, 10, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
'mariano 19112014
Sub GenerarRecordsetTemporalProductosIngresados()
    With mrs_Tmp
        .Fields.Append "fechaCreacion", adDBTimeStamp, 10, adFldIsNullable
        .Fields.Append "codigo", adVarChar, 10, adFldIsNullable
        .Fields.Append "nombre", adVarChar, 510, adFldIsNullable
        .Fields.Append "preciou", adDouble
        .Fields.Append "concepto", adVarChar, 50, adFldIsNullable
        .Fields.Append "idAlmacenDestino", adInteger, 4, adFldIsNullable
        .Fields.Append "idAlmacenOrigen", adInteger, 4, adFldIsNullable
        .Fields.Append "lote", adVarChar, 20, adFldIsNullable
        .Fields.Append "cantidad", adInteger, 4, adFldIsNullable
        .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
        .Fields.Append "monto", adDouble
        .LockType = adLockOptimistic
        .Open
    End With
End Sub



Private Sub Form_Resize()
    CrvReportes.Top = 0 '500
    CrvReportes.Left = 0
    CrvReportes.Height = ScaleHeight
    CrvReportes.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
    LimpiarVariablesDeMemoria
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mrs_Tmp = Nothing
    Set rsReporte = Nothing
    Set rsTmp = Nothing
    'Set rsErrores = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasCaja = Nothing
    Set mo_ReglasReportes = Nothing
    Set oConexion = Nothing
    Set mo_DoFarmMovimientoVentas = Nothing
End Sub




