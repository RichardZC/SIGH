VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasFarmacia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Farmacia
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Const lcConstanteMovimientoEntrada As String = "E"
Const lcConstanteMovimientoSalida As String = "S"
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim oRsDetalleProductosConLotes As New Recordset
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Const lcLetrasCodigoPaquete As String = "PQ"
Dim mo_idOrdenPago As Long

Property Get idOrdenPago() As Long
   idOrdenPago = mo_idOrdenPago
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

' MODIFICADO POR YAMILL PALOMINO
Sub BloqueaTabla(oRs As Recordset)
    On Error GoTo ErrBloqueo


Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "BloqueaTabla"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
Exit Sub
ErrBloqueo:
    MsgBox "Tabla esta Bloqueada"
    
    
End Sub

Function ChequeaQueSaldosSinLotesSeaPositivo(lnIdAlmacen As Long, lnIdProducto As Long, cantidad As Long) As Boolean
    Dim oRsTmp As New Recordset
    ChequeaQueSaldosSinLotesSeaPositivo = False
    Set oRsTmp = farmDevuelveSaldosSegunAlmacenProducto(lnIdAlmacen, lnIdProducto)
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!cantidad >= cantidad Then
          ChequeaQueSaldosSinLotesSeaPositivo = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function ChequeaQueSaldosConLotesSeaPositivo(oConexion As ADODB.Connection, lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnIdTipoSalidaBienInsumo As Long, cantidad As Long) As Boolean
    Dim oRsTmp As New Recordset
    Dim oFarmMovimientoDetalle As New FarmMovimientoDetalle
    ChequeaQueSaldosConLotesSeaPositivo = False
    Set oFarmMovimientoDetalle.Conexion = oConexion
    Set oRsTmp = oFarmMovimientoDetalle.farmDevuelveSaldosSegunAlmacenProductoLote(lnIdAlmacen, lnIdProducto, lcLote, ldFechaVencimiento, lnIdTipoSalidaBienInsumo)
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!cantidad >= cantidad Then
          ChequeaQueSaldosConLotesSeaPositivo = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oFarmMovimientoDetalle = Nothing
End Function

Function DevuelveProductosConLotes()
    Set DevuelveProductosConLotes = oRsDetalleProductosConLotes.Clone
End Function

Function DevuelveDetalleProductosConLotes(oConexion As ADODB.Connection, lnIdAlmacen As Long, oRsConsolidado As Recordset, _
                                          ldFechaActual As Date, lnIdTipoSalidaBienInsumo As Long, wxParametro347 As String) As Boolean
    On Error Resume Next
    DevuelveDetalleProductosConLotes = False
    If oRsDetalleProductosConLotes.State = 1 Then
       Set oRsDetalleProductosConLotes = Nothing
    End If
    With oRsDetalleProductosConLotes
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "Codigo", adChar, 20, adFldIsNullable
          .Fields.Append "NombreProducto", adChar, 300, adFldIsNullable
          .Fields.Append "Lote", adChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "idTipoSalidaBienInsumo", adInteger
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "NumeroDocumento", adVarChar, 20, adFldIsNullable
          .Fields.Append "RegistroSanitario", adVarChar, 50, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
          'Carga Lotes/FechaVencimiento
          Dim oRsSaldosEnEsteMomento As New Recordset
          Dim lnCantidadCargar As Integer
          Dim lnCantSaldo As Long
          Dim oFarmMovimientoDetalle As New FarmMovimientoDetalle
          Dim lbProductoConLoteOK As Boolean
          Dim LdFechaMinimaDespacho As Date
          LdFechaMinimaDespacho = ldFechaActual + Val(lcBuscaParametro.SeleccionaFilaParametro(220))
          Set oFarmMovimientoDetalle.Conexion = oConexion
          'limpiar
          If .RecordCount > 0 Then
             .MoveFirst
             Do While Not .EOF
                .Delete
                .Update
                .MoveNext
             Loop
          End If
          ms_MensajeError = ""
          oRsConsolidado.MoveFirst
          Do While Not oRsConsolidado.EOF
             Set oRsSaldosEnEsteMomento = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen, 0, _
                                          Trim(oRsConsolidado.Fields!Codigo), oConexion)
             If wxParametro347 <> "S" Then
                oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & lnIdTipoSalidaBienInsumo
             End If
             lbProductoConLoteOK = False
             If oRsSaldosEnEsteMomento.RecordCount > 0 Then
                lnCantSaldo = oRsConsolidado.Fields!cantidad
                oRsSaldosEnEsteMomento.MoveFirst
                Do While Not oRsSaldosEnEsteMomento.EOF
                   If oRsSaldosEnEsteMomento.Fields!saldo > 0 And oRsSaldosEnEsteMomento.Fields!FechaVencimiento >= LdFechaMinimaDespacho Then
                        lbProductoConLoteOK = True
                        .AddNew
                        .Fields!idProducto = oRsConsolidado.Fields!idProducto
                        .Fields!Codigo = oRsConsolidado.Fields!Codigo
                        .Fields!nombreProducto = oRsConsolidado.Fields!nombreProducto
                        .Fields!Lote = oRsSaldosEnEsteMomento.Fields!Lote
                        .Fields!FechaVencimiento = oRsSaldosEnEsteMomento.Fields!FechaVencimiento
                        .Fields!idTipoSalidaBienInsumo = oRsSaldosEnEsteMomento.Fields!idTipoSalidaBienInsumoSaldo
                        .Fields!Precio = oRsConsolidado.Fields!Precio
                        If lnCantSaldo >= oRsSaldosEnEsteMomento.Fields!saldo Then
                           lnCantidadCargar = oRsSaldosEnEsteMomento.Fields!saldo
                           .Fields!cantidad = lnCantidadCargar
                           .Fields!total = Round(lnCantidadCargar * oRsConsolidado.Fields!Precio, 2)
                           .Update
                           lnCantSaldo = lnCantSaldo - oRsSaldosEnEsteMomento.Fields!saldo
                           If lnCantSaldo <= 0 Then
                              Exit Do
                           End If
                        Else
                           lnCantidadCargar = lnCantSaldo
                           .Fields!cantidad = lnCantidadCargar
                           .Fields!total = Round(lnCantidadCargar * oRsConsolidado.Fields!Precio, 2)
                           .Update
                           lnCantSaldo = 0
                           Exit Do
                        End If
                   End If
                   oRsSaldosEnEsteMomento.MoveNext
                Loop
                If lnCantSaldo > 0 Then
                   lbProductoConLoteOK = False
                End If
                
             Else
                ms_MensajeError = ms_MensajeError & Trim(oRsConsolidado.Fields!Codigo) & " " & oRsConsolidado.Fields!nombreProducto & Chr(13)
                Exit Function
             End If
             oRsSaldosEnEsteMomento.Close
             If lbProductoConLoteOK = False Then
                 ms_MensajeError = ms_MensajeError & "(F.Mínima despacho: " & Format(LdFechaMinimaDespacho, sighentidades.DevuelveFechaSoloFormato_DMY) & ") " & Trim(oRsConsolidado.Fields!Codigo) & " " & oRsConsolidado.Fields!nombreProducto & Chr(13)
                 Exit Function
             End If
             oRsConsolidado.MoveNext
          Loop
          Set oRsSaldosEnEsteMomento = Nothing
          Set oFarmMovimientoDetalle = Nothing
          DevuelveDetalleProductosConLotes = True
    End With
End Function

''********************************  CAJA   ************************************
'Function CajaComprobantePagoBienesRegistraBoleta(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, oDOFactOrdenBien As DoFactOrdenesBienes, oRsDetalleProductos As Recordset, ml_IdUsuarioAuditoria As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oFactOrdenesBienes As New FactOrdenesBienes
'Dim oCajaComprobantePago As New CajaComprobantesPago
'Dim oCajaNroDocumento As New CajaNroDocumento
'Dim oMovimiento As New farmMovimiento
'Dim oDoMovimiento As New DoFarmMovimiento
'Dim oFarmMovimientoVentas As New farmMovimientoVentas
'Dim oDoFarmMovimientoVentas As New DoFarmMovimientoVentas
'Dim oMovimientoVentasDetalle As New FarmMovimientoVentasDetalle
'Dim oDoMovimientoVentasDetalle As New DoFarmMovimientoVentasDetalle
'Dim oMovimientoDetalle As New FarmMovimientoDetalle
'Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
'Dim oDoFarmPreVenta As New DoFarmPreVenta
'Dim oFarmPreventa As New FarmPreVenta
'Dim oDoRecetaDetalle As New DoRecetaDetalle, oRecetaDetalle As New RecetaDetalle
'Dim oDoRecetaCabecera As New DoRecetaCabecera, oRecetaCabecera As New RecetaCabecera
'Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
'Dim lcCorrelativo As String
'Dim bProcesoOK As Boolean
'Dim sResultado As String
'Dim oDoCajaNroDocumento As New DOCajaNroDocumento
'Dim lcTicket As String
'Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim lnItem As Long
'Dim ml_IdFuenteFinanciamiento As Long
'Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
'Dim lnIdReceta As Long
'Dim lcSql As String
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'
'    CajaComprobantePagoBienesRegistraBoleta = False
'
'    'MODIFICADO POR YAMILL PALOMINO
'
'    oConexion.CursorLocation = adUseClient
'    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
'
'    oConexion.BeginTrans
'    Set oFactOrdenesBienes.Conexion = oConexion
'    Set oCajaComprobantePago.Conexion = oConexion
'    Set oMovimientoDetalle.Conexion = oConexion
'    Set oMovimiento.Conexion = oConexion
'    Set oFarmPreventa.Conexion = oConexion
'    Set oFarmMovimientoVentas.Conexion = oConexion
'    Set oMovimientoVentasDetalle.Conexion = oConexion
'    Set oCajaNroDocumento.Conexion = oConexion
'    Set oRecetaCabecera.Conexion = oConexion
'    Set oRecetaDetalleItem.Conexion = oConexion
'
'    '*********comprobante pago
'
'    oDOCajaComprobantePago.IdTipoOrden = 2  'Bienes e Insumos
'    oDOCajaComprobantePago.idEstadoComprobante = 4    'pagado
'    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'    '
'    oDoFarmPreVenta.idPreVenta = oDOFactOrdenBien.idPreVenta
'    If Not oFarmPreventa.SeleccionarPorId(oDoFarmPreVenta) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'    '
'
'    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
'    '*********  graba tabla correlativos farmTipoDocumentos  ***************
'    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
'    '*********  graba tabla Movimientos  ***************
'    lcTicket = Trim(oDOCajaComprobantePago.NroSerie) & "-" & Trim(oDOCajaComprobantePago.NroDocumento)
'    With oDoMovimiento
'       .FechaCreacion = oDOCajaComprobantePago.FechaCobranza
'       .idAlmacenDestino = 0  'Ninguno
'       .idAlmacenOrigen = oDOCajaComprobantePago.idFarmacia
'       .idEstadoMovimiento = 1
'       .idTipoConcepto = 10
'       .idUsuario = oDOCajaComprobantePago.IdUsuarioAuditoria
'       .MovTipo = "S"
'       .MovNumero = lcCorrelativo
'       .DocumentoNumero = lcTicket
'       .DocumentoIdtipo = 26
'       .total = oDOCajaComprobantePago.total
'       .idUsuario = oDoFarmPreVenta.idVendedor
'       .IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'    End With
'
'    If Not oMovimiento.Insertar(oDoMovimiento) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'    '*************vuelve a chequear saldos ****************
'    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloVenta) Then
'        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
'         bProcesoOK = False: GoTo ErrorManager1
'    End If
'    oRsDetalleProductosConLotes.MoveFirst
'    Do While Not oRsDetalleProductosConLotes.EOF
'       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
'                                    oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, _
'                                    oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
'                                    oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
'                                    oRsDetalleProductosConLotes.Fields!cantidad) Then
'             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + _
'                               Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + _
'                               Trim(oRsDetalleProductosConLotes.Fields!nombreProducto) + Chr(13) + _
'                               "(Farmacia debe modificar PREVENTA y controlar mejor su STOCK MINIMO)" + Chr(13)
'             bProcesoOK = False: GoTo ErrorManager1
'       ElseIf ChequeaQueSaldosSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
'                                          oRsDetalleProductosConLotes.Fields!idProducto, _
'                                          oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
'                                          oRsDetalleProductosConLotes.Fields!cantidad) = False Then
'             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + _
'                               Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + _
'                               Trim(oRsDetalleProductosConLotes.Fields!nombreProducto) + Chr(13) + _
'                               "(Farmacia debe modificar PREVENTA y controlar mejor su STOCK MINIMO)" + Chr(13)
'             bProcesoOK = False: GoTo ErrorManager1
'       End If
'       oRsDetalleProductosConLotes.MoveNext
'    Loop
'    '********* preventas
'    oDoFarmPreVenta.idEstadoPreventa = 2   'con documento emitido
'    oDoFarmPreVenta.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'    If Not oFarmPreventa.Modificar(oDoFarmPreVenta) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'    ml_IdFuenteFinanciamiento = 0    'Particular= tenga o  no "Nro CUENTA", Historia
'    '*********  graba tabla MovimientosVentas  ***************
'    With oDoFarmMovimientoVentas
'       .MovTipo = oDoMovimiento.MovTipo
'       .MovNumero = oDoMovimiento.MovNumero
'       .IdCuentaAtencion = oDoFarmPreVenta.IdCuentaAtencion
'       .IdDiagnostico = oDoFarmPreVenta.IdDiagnostico
'       .idPaciente = oDoFarmPreVenta.idPaciente
'       .idPrescriptor = oDoFarmPreVenta.idPrescriptor
'       .idPreVenta = oDoFarmPreVenta.idPreVenta
'       .IdTipoFinanciamiento = oDoFarmPreVenta.IdTipoFinanciamiento
'       .idTipoReceta = oDoFarmPreVenta.idTipoReceta
'       .IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'       .tipoVenta = "P"
'       .idFuenteFinanciamiento = ml_IdFuenteFinanciamiento
'       If .IdCuentaAtencion > 0 Then
'          .IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(.IdCuentaAtencion, Format(oDoMovimiento.FechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY), Format(oDoMovimiento.FechaCreacion, "hh:mm"))
'       Else
'          .IdServicioPaciente = 0
'       End If
'       .FechaHoraPrescribe = oDoFarmPreVenta.FechaHoraPrescribe
'    End With
'    If Not oFarmMovimientoVentas.Insertar(oDoFarmMovimientoVentas) Then
'             bProcesoOK = False: GoTo ErrorManager1
'    End If
'    '************ FactOrdenBien
'
'    oDOFactOrdenBien.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
'    oDOFactOrdenBien.IdEstadoFacturacion = 4  'pagada
'    oDOFactOrdenBien.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'    oDOFactOrdenBien.MovNumero = oDoMovimiento.MovNumero
'    oDOFactOrdenBien.MovTipo = oDoMovimiento.MovTipo
'    If Not oFactOrdenesBienes.Modificar(oDOFactOrdenBien) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
'
'    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'    oDoMovimientoDetalle.MovNumero = oDoMovimiento.MovNumero
'    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
'    lnItem = 1
'    oRsDetalleProductosConLotes.MoveFirst
'    Do While Not oRsDetalleProductosConLotes.EOF
'       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
'       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
'       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
'       oDoMovimientoDetalle.Item = lnItem
'       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
'       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
'       oDoMovimientoDetalle.RegistroSanitario = ""
'       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
'       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
'       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
'           bProcesoOK = False: GoTo ErrorManager1
'       End If
'       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(oDoMovimiento.MovTipo, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
'           bProcesoOK = False: GoTo ErrorManager1
'       End If
'       lnItem = lnItem + 1
'       oRsDetalleProductosConLotes.MoveNext
'    Loop
'    '*********  graba tabla farmMovimientosVentasDetalle, FacturacionBienesFinanciamientos ***************
'
'    oDoMovimientoVentasDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'    oDoMovimientoVentasDetalle.MovNumero = oDoMovimiento.MovNumero
'    oDoMovimientoVentasDetalle.MovTipo = oDoMovimiento.MovTipo
'    oRsDetalleProductos.MoveFirst
'    Do While Not oRsDetalleProductos.EOF
'       oDoMovimientoVentasDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
'       oDoMovimientoVentasDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
'       oDoMovimientoVentasDetalle.Precio = oRsDetalleProductos.Fields!Precio
'       oDoMovimientoVentasDetalle.total = oRsDetalleProductos.Fields!total
'       If Not oMovimientoVentasDetalle.Insertar(oDoMovimientoVentasDetalle) Then
'           bProcesoOK = False: GoTo ErrorManager1
'       End If
'       oRsDetalleProductos.MoveNext
'    Loop
'    '************* Actualiza Nro Serie y Documento
'
'    oDoCajaNroDocumento.IdCaja = oDOGestionCaja.IdCaja
'    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
'    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
'    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
'    If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
'        bProcesoOK = False: GoTo ErrorManager1
'    End If
'
'
'    'Receta automatica debb-03/04/2012
'    If oDoFarmPreVenta.idPreVenta > 0 Then
'        If oRecetaCabecera.ConsultaRecetaCabeceraFarmacia(oDoRecetaCabecera, Trim(Str(oDoFarmPreVenta.idPreVenta))) = True Then
'             lnIdReceta = oDoRecetaCabecera.idReceta
'             oDoRecetaCabecera.IdEstado = 3
'             oDoRecetaCabecera.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
'             oDoRecetaCabecera.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
'             If oRecetaCabecera.Modificar(oDoRecetaCabecera) = False Then
'                bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo ErrorManager1
'             End If
'             Set oRecetaDetalle.Conexion = oConexion
'             Set oRsTmp1 = oRecetaDetalle.SeleccionarPorIdReceta(lnIdReceta)
'
'             oRsDetalleProductos.MoveFirst
'             Do While Not oRsDetalleProductos.EOF
'               oRsTmp1.MoveFirst
'               oRsTmp1.Find "idItem=" & oRsDetalleProductos.Fields!idProducto
'               If Not oRsTmp1.EOF Then
'                     oDoRecetaDetalle.idReceta = lnIdReceta
'                     oDoRecetaDetalle.idItem = oRsDetalleProductos.Fields!idProducto
'                     oDoRecetaDetalle.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
'                     If oRecetaDetalle.RecetaDetalleActualizaCantDespachada(oDoRecetaDetalle) = False Then
'                         bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo ErrorManager1
'                     End If
'                     oDoRecetaDetalleItem.idReceta = lnIdReceta
'                     oDoRecetaDetalleItem.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
'                     oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDoFarmPreVenta.idPreVenta))
'                     oDoRecetaDetalleItem.idItem = oRsDetalleProductos.Fields!idProducto
'                     oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoFarmPreVenta.IdUsuarioAuditoria
'                     If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
'                        bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo ErrorManager1
'                     End If
'               End If
'               oRsDetalleProductos.MoveNext
'             Loop
'             oRsTmp1.Close
'        End If
'    End If
'    '
'    oConexion.CommitTrans
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oCommand = Nothing
'    ms_MensajeError = ""
'    CajaComprobantePagoBienesRegistraBoleta = True
'    Exit Function
'ErrorManager1:
'    oConexion.RollbackTrans
'    oConexion.Close
'    Set oConexion = Nothing
'    ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oCajaComprobantePago.MensajeError <> "", oCajaComprobantePago.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oCajaNroDocumento.MensajeError <> "", oCajaNroDocumento.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oFarmMovimientoVentas.MensajeError <> "", oFarmMovimientoVentas.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentasDetalle.MensajeError <> "", oMovimientoVentasDetalle.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
'    ms_MensajeError = ms_MensajeError + IIf(oFarmPreventa.MensajeError <> "", oFarmPreventa.MensajeError + Chr(13), "")
'    CajaComprobantePagoBienesRegistraBoleta = False
'    Set oFactOrdenesBienes = Nothing
'    Set oCajaComprobantePago = Nothing
'    Set oCajaNroDocumento = Nothing
'    Set oMovimiento = Nothing
'    Set oFarmMovimientoVentas = Nothing
'    Set oDoFarmMovimientoVentas = Nothing
'    Set oMovimientoVentasDetalle = Nothing
'    Set oDoMovimientoDetalle = Nothing
'    Set oDoFarmPreVenta = Nothing
'    Set oFarmPreventa = Nothing
'    Set oDoCajaNroDocumento = Nothing
'    Set lcBuscaParametro = Nothing
'    Set oRsTmp1 = Nothing
'    Set oDoRecetaDetalle = Nothing
'    Set oRecetaDetalle = Nothing
'    Set oDoRecetaCabecera = Nothing
'    Set oRecetaCabecera = Nothing
'    Set oRecetaDetalleItem = Nothing
'    Set oDoRecetaDetalleItem = Nothing
'End Function

Function CajaComprobantePagoBienesAnulaBoleta(oDOCajaComprobantePago As DOCajaComprobantesPago, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFacturacionPagos As New FacturacionBienesPagos
Dim oDoFacturacionPagos As New DoFacturacionBienesPagos
Dim oFactOrdenesBienes As New FactOrdenesBienes
Dim oDOFactOrdenBien As New DoFactOrdenesBienes
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim oMovimiento As New farmMovimiento
Dim oDoMovimiento As New DoFarmMovimiento
Dim oMovimientoDetalle As New FarmMovimientoDetalle
Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
Dim oDoFarmPreVenta As New DoFarmPreVenta
Dim oFarmPreventa As New FarmPreVenta
Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
Dim oDoFacturacionBienesFinanciam As New DoFacturacionBienesFinanciam
Dim oDoFinanciamientoCatalogoBien As New DoFinanciamientoCatalogoBien
Dim oFinanciamientoCatalogoBien As New FinanciamientoCatalogoBien
Dim oFarmMovimientoVentas As New farmMovimientoVentas
Dim oDoFarmMovimientoVentas As New DoFarmMovimientoVentas
Dim mo_ReglasImagenes As New ReglasImagenes
Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
Dim lcCorrelativo As String
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim lcTicket As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnItem As Long
Dim oRsListaOrdenes As New Recordset
Dim oRsListaOrdenesDetalle As New Recordset
Dim rs As New Recordset
Dim oRsTmp1 As New Recordset
Dim lnIdOrden As Long
Dim ldFechaAnulacion As Date
Dim lcSql As String, lnIdReceta As Long
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    CajaComprobantePagoBienesAnulaBoleta = False
    
    ' Modificado por Yamill Palomino
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oCajaComprobantePago.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oMovimiento.Conexion = oConexion
    Set oFarmPreventa.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    Set oFacturacionPagos.Conexion = oConexion
    Set oFarmMovimientoVentas.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    '
    ldFechaAnulacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
    '*********comprobante pago
    oDOCajaComprobantePago.idEstadoComprobante = 9   'anulada
    If Not oCajaComprobantePago.modificar(oDOCajaComprobantePago) Then
        bProcesoOK = False: GoTo ErrorManager1El
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "E", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
    
    Set oRsListaOrdenes = FactOrdenesBienesSeleccionarPorIdComprobantePago(oDOCajaComprobantePago.IdComprobantePago)
    If oRsListaOrdenes.RecordCount > 0 Then
        oRsListaOrdenes.MoveFirst
        Do While Not oRsListaOrdenes.EOF
            '************ FactOrdenBien
            oDOFactOrdenBien.IdOrden = oRsListaOrdenes.Fields!IdOrden
            If Not oFactOrdenesBienes.SeleccionarPorId(oDOFactOrdenBien) Then
                bProcesoOK = False: GoTo ErrorManager1El
            End If
            oDOFactOrdenBien.IdEstadoFacturacion = 9   'anulado
            oDOFactOrdenBien.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
            oDOFactOrdenBien.idPreVenta = 0
            If IsNull(oRsListaOrdenes.Fields!idPreVenta) Or oRsListaOrdenes.Fields!idPreVenta = 0 Then
               'es una Cuenta pagada, no debe mostrar pagos en VENTANA=ESTADO CUENTA, para que  vuelva a pagar
               oDOFactOrdenBien.movNumero = ""
               oDOFactOrdenBien.MovTipo = ""
            End If
            oDOFactOrdenBien.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
            If Not oFactOrdenesBienes.modificar(oDOFactOrdenBien) Then
                bProcesoOK = False: GoTo ErrorManager1El
            End If
            '
            '****************************************Solo cuando es Preventa****************************************
            If oRsListaOrdenes.Fields!idPreVenta > 0 And Not IsNull(oRsListaOrdenes.Fields!movNumero) Then
               '************ FarmPreVenta
               oDoFarmPreVenta.idPreVenta = oRsListaOrdenes.Fields!idPreVenta
               If Not oFarmPreventa.SeleccionarPorId(oDoFarmPreVenta) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               oDoFarmPreVenta.idEstadoPreventa = 1  'por Emitir Documento
               oDoFarmPreVenta.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               If Not oFarmPreventa.modificar(oDoFarmPreVenta) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               '************ FarmMovimiento
               oDoMovimiento.movNumero = oRsListaOrdenes.Fields!movNumero
               oDoMovimiento.MovTipo = oRsListaOrdenes.Fields!MovTipo
               If Not oMovimiento.SeleccionarPorId(oDoMovimiento) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               oDoMovimiento.idEstadoMovimiento = 0   'Anulado
               oDoMovimiento.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               If Not oMovimiento.modificar(oDoMovimiento) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               '************FarmMovimientoVenta
               oDoFarmMovimientoVentas.movNumero = oRsListaOrdenes.Fields!movNumero
               oDoFarmMovimientoVentas.MovTipo = oRsListaOrdenes.Fields!MovTipo
               If Not oFarmMovimientoVentas.SeleccionarPorId(oDoFarmMovimientoVentas) Then
                  bProcesoOK = False: GoTo ErrorManager1El
               End If
               oDoFarmMovimientoVentas.idPreVenta = 0
               oDoFarmMovimientoVentas.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               If Not oFarmMovimientoVentas.modificar(oDoFarmMovimientoVentas) Then
                  bProcesoOK = False: GoTo ErrorManager1El
               End If
               '*********  devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
               Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
               If rs.RecordCount > 0 Then
                    oMovimientoDetalle.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
                    rs.MoveFirst
                    Do While Not rs.EOF
                        If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                           bProcesoOK = False: GoTo ErrorManager1El
                        End If
                        rs.MoveNext
                    Loop
               End If
               rs.Close
               '*********  agrega CABECERA de pagos ***************
               oDOFactOrdenBien.FechaCreacion = oDoFarmPreVenta.FechaCreacion
               oDOFactOrdenBien.idCuentaAtencion = oDoFarmPreVenta.idCuentaAtencion
               oDOFactOrdenBien.IdEstadoFacturacion = 1 'registrado
               oDOFactOrdenBien.idPaciente = oDoFarmPreVenta.idPaciente
               oDOFactOrdenBien.idPreVenta = oDoFarmPreVenta.idPreVenta
               oDOFactOrdenBien.IdPuntoCarga = 5  'farmacia
               oDOFactOrdenBien.idUsuario = oDoFarmPreVenta.idUsuario
               oDOFactOrdenBien.IdComprobantePago = 0
               oDOFactOrdenBien.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               If Not oFactOrdenesBienes.Insertar(oDOFactOrdenBien) Then
                  bProcesoOK = False: GoTo ErrorManager1El
               End If
               'Agrega Detalle de PAGOS
               oDoFacturacionPagos.IdOrden = oDOFactOrdenBien.IdOrden
               oDoFacturacionPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               Set rs = farmMovimientoVentasDetalleSeleccionarPorId(oDoMovimiento.movNumero, oDoMovimiento.MovTipo)
               If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        oDoFacturacionPagos.CantidadPagar = rs.Fields!cantidad
                        oDoFacturacionPagos.idProducto = rs.Fields!idProducto
                        oDoFacturacionPagos.PrecioVenta = rs.Fields!Precio
                        oDoFacturacionPagos.TotalPagar = rs.Fields!total
                        If Not oFacturacionPagos.Insertar(oDoFacturacionPagos) Then
                           bProcesoOK = False: GoTo ErrorManager1El
                        End If
                       rs.MoveNext
                    Loop
               End If
               rs.Close
            End If
            '
            '***********Solo cuando tiene Nro cuenta, devuelve Cantidades financiadas****************************************
            If oDOFactOrdenBien.idCuentaAtencion > 0 And (IsNull(oRsListaOrdenes.Fields!idPreVenta) Or oRsListaOrdenes.Fields!idPreVenta = 0) Then
               'agregar Cabecera Pendiente de Pago
               oDOFactOrdenBien.movNumero = oRsListaOrdenes.Fields!movNumero
               oDOFactOrdenBien.MovTipo = oRsListaOrdenes.Fields!MovTipo
               oDOFactOrdenBien.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               oDOFactOrdenBien.IdEstadoFacturacion = 1    'registrada
               oDOFactOrdenBien.IdComprobantePago = 0
               If Not oFactOrdenesBienes.Insertar(oDOFactOrdenBien) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               'agregar Detalle Pendiente de Pago
               lnIdOrden = oRsListaOrdenes.Fields!IdOrden
               Do While Not oRsListaOrdenes.EOF And lnIdOrden = oRsListaOrdenes.Fields!IdOrden
                    oDoFacturacionPagos.IdOrden = oDOFactOrdenBien.IdOrden
                    oDoFacturacionPagos.CantidadPagar = oRsListaOrdenes.Fields!CantidadPagar
                    oDoFacturacionPagos.idProducto = oRsListaOrdenes.Fields!idProducto
                    oDoFacturacionPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
                    oDoFacturacionPagos.PrecioVenta = oRsListaOrdenes.Fields!PrecioVenta
                    oDoFacturacionPagos.TotalPagar = oRsListaOrdenes.Fields!TotalPagar
                    If Not oFacturacionPagos.Insertar(oDoFacturacionPagos) Then
                        bProcesoOK = False: GoTo ErrorManager1El
                    End If
                    oRsListaOrdenes.MoveNext
                    If oRsListaOrdenes.EOF Then
                       Exit Do
                    End If
               Loop
            Else
               lnIdOrden = oRsListaOrdenes.Fields!IdOrden
               Do While Not oRsListaOrdenes.EOF And lnIdOrden = oRsListaOrdenes.Fields!IdOrden
                  oRsListaOrdenes.MoveNext
                  If oRsListaOrdenes.EOF Then
                     Exit Do
                  End If
               Loop
            End If
        Loop
    End If
    oRsListaOrdenes.Close
    
    'Receta automatica debb-03/04/2012
    If oDoFarmPreVenta.idPreVenta > 0 Then
        If oRecetaCabecera.ConsultaRecetaCabeceraFarmacia(oDoRecetaCabecera, Trim(Str(oDoFarmPreVenta.idPreVenta))) = True Then
           lnIdReceta = oDoRecetaCabecera.idReceta
           mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalleItem oDOCajaComprobantePago.IdComprobantePago, oConexion
        End If
    End If
    
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagoBienesAnulaBoleta = True
ErrorManager1El:
    If CajaComprobantePagoBienesAnulaBoleta = False Then
        oConexion.RollbackTrans
        oConexion.Close
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oCajaComprobantePago.MensajeError <> "", oCajaComprobantePago.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFarmPreventa.MensajeError <> "", oFarmPreventa.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFarmMovimientoVentas.MensajeError <> "", oFarmMovimientoVentas.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionPagos.MensajeError <> "", oFacturacionPagos.MensajeError + Chr(13), "")
    End If
    Set oConexion = Nothing
    Set oFacturacionPagos = Nothing
    Set oDoFacturacionPagos = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oDOFactOrdenBien = Nothing
    Set oCajaComprobantePago = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oDoFarmPreVenta = Nothing
    Set oFarmPreventa = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oDoFacturacionBienesFinanciam = Nothing
    Set oDoFinanciamientoCatalogoBien = Nothing
    Set oFinanciamientoCatalogoBien = Nothing
    Set lcBuscaParametro = Nothing
    Set oFarmMovimientoVentas = Nothing
    Set oDoFarmMovimientoVentas = Nothing
    Set oRsListaOrdenes = Nothing
    Set oRsListaOrdenesDetalle = Nothing
    Set rs = Nothing
    Set oRsTmp1 = Nothing
    Set mo_ReglasImagenes = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaCabecera = Nothing
End Function
'********************************  CAJA   ************************************


'********************************   venta Directa  ***************************************
Function AgregaDatosDeVentaDirecta(ByVal oDoMovimiento As DoFarmMovimiento, oDoFarmMovimientoVentas As DoFarmMovimientoVentas, _
                                   oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                   mo_lcNombrePc As String, lnIdReceta As Long, lnEpsPorcentaje As Double) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoVentas As New farmMovimientoVentas
    Dim oMovimientoVentasDetalle As New FarmMovimientoVentasDetalle
    Dim oDoMovimientoVentasDetalle As New DoFarmMovimientoVentasDetalle
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinanciam As New DoFacturacionBienesFinanciam
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim rsTmp As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lcCorrelativo As String: Dim lcTicket As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oDoCajaNroDocumento As New DOCajaNroDocumento
    Dim oCajaCaja As New CajaCaja
    Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
    Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
    Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
    Dim lbGeneraPago As Boolean
    Dim lcSql As String, lnImporteTotal As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oCajaCaja.Conexion = oConexion
    Set oMovimientoVentas.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oMovimientoVentasDetalle.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalleItem.Conexion = oConexion
    
    
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla correlativos CajaNroDocumento (Devuelve Ticket=serie+correlativo)  ***************
    oDoCajaNroDocumento.idCaja = Val(lcBuscaParametro.SeleccionaFilaParametro(225))    'Farmacia
    oDoCajaNroDocumento.idTipoComprobante = Val(lcBuscaParametro.SeleccionaFilaParametro(226))   'Ticket
    If Not oCajaCaja.ObtenerSiguienteNumeroDocumentoYgrabarlo(oDoCajaNroDocumento) Then
       bProcesoOK = False:  ms_MensajeError = ms_MensajeError + "Problemas con parametros 225 y/o 226": GoTo TerminarIVDTA
    End If
    lcTicket = Trim(oDoCajaNroDocumento.NroSerie) & "-" & oDoCajaNroDocumento.NroDocumento
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
       .DocumentoNumero = lcTicket
       .DocumentoIdtipo = Val(lcBuscaParametro.SeleccionaFilaParametro(227))   'Ticket
    End With
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla MOVIMIENTO": GoTo TerminarIVDTA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********** revisa nuevamente los saldos *************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloVenta, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarIVDTA
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
        ms_MensajeError = ms_MensajeError + "No existe Productos"
        bProcesoOK = False: GoTo TerminarIVDTA
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarIVDTA
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '*********  graba tabla MovimientosVentas  ***************
    With oDoFarmMovimientoVentas
       .movNumero = lcCorrelativo
       If .idCuentaAtencion > 0 Then
          .IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(.idCuentaAtencion, _
                      Format(oDoMovimiento.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), _
                      Format(oDoMovimiento.FechaCreacion, "hh:mm"), oConexion)
       Else
          .IdServicioPaciente = 0
       End If
    End With
    If Not oMovimientoVentas.Insertar(oDoFarmMovimientoVentas) Then
            bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla MOVIMIENTOVENTAS": GoTo TerminarIVDTA
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla MOVIMIENTODETALLE": GoTo TerminarIVDTA
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con ACTUALIZARSALDOSPORPRODUCTO: GoTo TerminarIVDTA"
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '
    lbGeneraPago = ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFarmMovimientoVentas.IdTipoFinanciamiento, oConexion)
    Dim lbEsPlanAtencion As Boolean
    
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "ConsultaBloqueaTablaXidTipoFinanciamiento"
          Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, oDoFarmMovimientoVentas.IdTipoFinanciamiento): .Parameters.Append oParameter
         Set rsTmp = .Execute
         Set rsTmp.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    If rsTmp.Fields!esFuenteFinanciamiento = True Then
       lbEsPlanAtencion = True
    Else
       lbEsPlanAtencion = False
       If rsTmp.Fields!IdTipoFinanciamiento = 9 Then  'Servicio Social genera datos en tablas de FACTURACION
          lbEsPlanAtencion = True
          lbGeneraPago = False
       End If
    End If
    rsTmp.Close
    '
    lnImporteTotal = 0
    If lbGeneraPago = True Then
        oDoFactOrdenesBienes.FechaCreacion = oDoMovimiento.FechaCreacion
        oDoFactOrdenesBienes.idCuentaAtencion = oDoFarmMovimientoVentas.idCuentaAtencion
        oDoFactOrdenesBienes.IdEstadoFacturacion = 1 'registrado
        oDoFactOrdenesBienes.idPaciente = oDoFarmMovimientoVentas.idPaciente
        oDoFactOrdenesBienes.idPreVenta = 0
        oDoFactOrdenesBienes.IdPuntoCarga = 5  'farmacia
        oDoFactOrdenesBienes.idUsuario = oDoFarmMovimientoVentas.IdUsuarioAuditoria
        oDoFactOrdenesBienes.IdUsuarioAuditoria = oDoFarmMovimientoVentas.IdUsuarioAuditoria
        oDoFactOrdenesBienes.movNumero = oDoMovimiento.movNumero
        oDoFactOrdenesBienes.MovTipo = oDoMovimiento.MovTipo
        If Not oFactOrdenesBienes.Insertar(oDoFactOrdenesBienes) Then
                bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla FactOrdenesBienes": GoTo TerminarIVDTA
        End If
    End If
    '*********  graba tabla farmMovimientosVentasDetalle, FacturacionBienesFinanciamientos ***************
    oDoMovimientoVentasDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoVentasDetalle.movNumero = lcCorrelativo
    oDoMovimientoVentasDetalle.MovTipo = lcConstanteMovimientoSalida
    '
    oDoFacturacionBienesFinanciam.IdTipoFinanciamiento = oDoFarmMovimientoVentas.IdTipoFinanciamiento
    oDoFacturacionBienesFinanciam.FechaAutoriza = oDoMovimiento.FechaCreacion
    oDoFacturacionBienesFinanciam.IdUsuarioAuditoria = oDoMovimiento.idUsuario
    oDoFacturacionBienesFinanciam.IdUsuarioAutoriza = oDoMovimiento.idUsuario
    oDoFacturacionBienesFinanciam.movNumero = oDoMovimiento.movNumero
    oDoFacturacionBienesFinanciam.MovTipo = oDoMovimiento.MovTipo
    oDoFacturacionBienesFinanciam.idFuenteFinanciamiento = oDoFarmMovimientoVentas.idFuenteFinanciamiento
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoVentasDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoVentasDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoVentasDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoVentasDetalle.total = oRsDetalleProductos.Fields!total
       If Not oMovimientoVentasDetalle.Insertar(oDoMovimientoVentasDetalle) Then
          bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla MovimientoVentasDetalle": GoTo TerminarIVDTA
       End If
       If lbGeneraPago = False And lbEsPlanAtencion = True Then
          oDoFacturacionBienesFinanciam.CantidadFinanciada = oRsDetalleProductos.Fields!cantidad
          oDoFacturacionBienesFinanciam.idProducto = oRsDetalleProductos.Fields!idProducto
          oDoFacturacionBienesFinanciam.PrecioFinanciado = oRsDetalleProductos.Fields!precioDelSeguro
          oDoFacturacionBienesFinanciam.TotalFinanciado = Round(oRsDetalleProductos.Fields!precioDelSeguro * oRsDetalleProductos.Fields!cantidad, 2)
          If Not oFacturacionBienesFinanciam.Insertar(oDoFacturacionBienesFinanciam) Then
               bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla FacturacionBienesFinanciamiento": GoTo TerminarIVDTA
          End If
          lnImporteTotal = lnImporteTotal + oDoFacturacionBienesFinanciam.TotalFinanciado
       End If
       '*********  graba tabla FacturacionBienesPagos ,solo si idtipofinanciamiento=5 (cred.hospit) ***************
       If lbGeneraPago = True And lbEsPlanAtencion = True Then
            oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
            oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFactOrdenesBienes.idUsuario
            oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!cantidad
            oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
            oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!Precio
            oDoFacturacionBienesPagos.TotalPagar = oRsDetalleProductos.Fields!total
            If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
                     bProcesoOK = False: ms_MensajeError = ms_MensajeError + "Problemas con tabla FacturacionBienesPagos ": GoTo TerminarIVDTA
            End If
       End If
       oRsDetalleProductos.MoveNext
    Loop

    'Receta automatica debb-03/04/2012

    If lnIdReceta > 0 Then
        oDoRecetaCabecera.idReceta = lnIdReceta
        If oRecetaCabecera.SeleccionarPorId(oDoRecetaCabecera) = True Then
             oDoRecetaCabecera.IdEstado = 2
             oDoRecetaCabecera.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
             If oRecetaCabecera.modificar(oDoRecetaCabecera) = False Then
                bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo TerminarIVDTA
             End If
             Set oRecetaDetalle.Conexion = oConexion
             Set oRsTmp1 = oRecetaDetalle.SeleccionarPorIdReceta(lnIdReceta)
             
             oRsDetalleProductos.MoveFirst
             Do While Not oRsDetalleProductos.EOF
               oRsTmp1.MoveFirst
               oRsTmp1.Find "idItem=" & oRsDetalleProductos.Fields!idProducto
               If Not oRsTmp1.EOF Then
                     oDoRecetaDetalle.idReceta = lnIdReceta
                     oDoRecetaDetalle.idItem = oRsDetalleProductos.Fields!idProducto
                     oDoRecetaDetalle.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
                     oDoRecetaDetalle.idEstadoDetalle = 2
                     If oRecetaDetalle.RecetaDetalleActualizaCantDespachada(oDoRecetaDetalle) = False Then
                         bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarIVDTA
                     End If
                     oDoRecetaDetalleItem.idReceta = lnIdReceta
                     oDoRecetaDetalleItem.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
                     oDoRecetaDetalleItem.DocumentoDespacho = oDoMovimiento.DocumentoNumero
                     oDoRecetaDetalleItem.idItem = oRsDetalleProductos.Fields!idProducto
                     oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
                     If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
                        bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIVDTA
                     End If
               End If
               oRsDetalleProductos.MoveNext
             Loop
             oRsTmp1.Close
        End If
    End If
    '
    AsignaPorcentajeEPS oConexion, oDoFarmMovimientoVentas.idCuentaAtencion, lnEpsPorcentaje, mo_lnIdTablaLISTBARITEMS, _
                         sghAgregar, oDoMovimiento.FechaCreacion, _
                         oDoFarmMovimientoVentas.IdServicioPaciente, oDoFarmMovimientoVentas.idPaciente, _
                         oDoFarmMovimientoVentas.IdUsuarioAuditoria, oDoFarmMovimientoVentas.movNumero, lnImporteTotal, _
                         sghPuntosCargaBasicos.sghPtoCargaFarmacia, oDoFarmMovimientoVentas.MovTipo, 0
    
TerminarIVDTA:
    If bProcesoOK Then
        AgregaDatosDeVentaDirecta = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeVentaDirecta = False
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentas.MensajeError <> "", oMovimientoVentas.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentasDetalle.MensajeError <> "", oMovimientoVentasDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesFinanciam.MensajeError <> "", oFacturacionBienesFinanciam.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoVentas = Nothing
    Set oMovimientoVentasDetalle = Nothing
    Set oDoMovimientoVentasDetalle = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oDoFacturacionBienesFinanciam = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    Set oFacturacionBienesPagos = Nothing
    Set rsTmp = Nothing
    Set oRsLookTablas = Nothing
    Set mo_ReglasCaja = Nothing
    Set oCajaCaja = Nothing
    Set oRecetaCabecera = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaDetalle = Nothing
    Set oRecetaDetalleItem = Nothing
    Set oDoRecetaDetalleItem = Nothing
End Function

Sub AsignaPorcentajeEPS(oConexion As Connection, lnIdCuentaAtencion As Long, lnEpsPorcentaje As Double, _
                        lnIdTablaLISTBARITEMS As sghOpcionGalenHos, lnOpcion As sghOpciones, _
                        ldFechaCreacion As Date, lnIdServicioIngreso As Long, lnIdPaciente As Long, lnIdUsuario As Long, _
                        lcMovNumero As String, lnImporteTotal As Double, lnIdPuntoCarga As Long, lcMovTipo As String, _
                        lnIdOrdenPago As Long)
    On Error GoTo ErrAsgPorc
    If lnEpsPorcentaje > 0 And lnEpsPorcentaje < 100 Then
       Dim lnImporteQpagara As Double
       Select Case lnIdTablaLISTBARITEMS
       Case sghVentasFarmacia
            If lnOpcion = sghAgregar Or lnOpcion = sghModificar Then
               lnImporteQpagara = Round(lnImporteTotal * ((100 - lnEpsPorcentaje) / 100), 2)
            
               '****disminuir precios,Totales segun porcentaje en tabla Seguros -CPT
               If EPSdisminuyePrecio(oConexion, lnEpsPorcentaje, lcMovNumero, lcMovTipo) = True Then
                  '****crear ORDEN DE PAGO por el porcentaje que no cubre
                  EPScrearOrdenPago oConexion, ldFechaCreacion, lnImporteQpagara, lnIdCuentaAtencion, lnIdServicioIngreso, _
                                    lnIdPaciente, lnIdUsuario, lnIdPuntoCarga, lnIdOrdenPago
               End If
            ElseIf lnOpcion = sghEliminar Then
            
            End If
       End Select
    End If
ErrAsgPorc:
End Sub

Function EPSdisminuyePrecio(oConexion As Connection, lnEpsPorcentaje As Double, lcMovNumero As String, lcMovTipo As String) As Boolean
    On Error GoTo ErrEPSDISMY
    EPSdisminuyePrecio = False
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "FacturacionBienesFinanciamientosActualizaPrecio"
        Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@porcentaje", adCurrency, adParamInput, 0, lnEpsPorcentaje): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    EPSdisminuyePrecio = True
    Exit Function
ErrEPSDISMY:
End Function


Function EPScrearOrdenPago(oConexion As Connection, ldFechaCreacion As Date, lnImporteQpagara As Double, lnIdCuenta As Long, lnIdServicioIngreso As Long, _
                          lnIdPaciente As Long, lnIdUsuario As Long, lnIdPuntoCarga As Long, lnIdOrdenPago As Long) As Boolean
      On Error GoTo ErrEPScrearOP
      EPScrearOrdenPago = False
      Dim oDoFactOrdenServ As New DoFactOrdenServ, oFactOrdenesServicio As New FactOrdenServicio
      Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho, oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos, oFactServicioPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos, oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim lnIdProductoEPS As Long, lnIdOrden As Long
      lnIdProductoEPS = Val(lcBuscaParametro.SeleccionaFilaParametro(566))
      Set oFactOrdenesServicio.Conexion = oConexion
      Set oFacturacionServicioDespacho.Conexion = oConexion
      Set oFactServicioPagos.Conexion = oConexion
      Set oFacturacionServicioPagos.Conexion = oConexion
      '
      If lnIdOrdenPago > 0 Then
            oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
            oDOFactOrdenServicioPagos.idUsuario = lnIdUsuario
            If oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) = True Then
                lnIdOrden = oDOFactOrdenServicioPagos.IdOrden
                oDoFacturacionServicioPagos.idOrdenPago = lnIdOrdenPago
                oDoFacturacionServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                oDoFacturacionServicioPagos.idProducto = lnIdProductoEPS
                If oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                    oDOFactOrdenServicioPagos.idOrdenPago = lnIdOrdenPago
                    oDOFactOrdenServicioPagos.IdUsuarioAuditoria = lnIdUsuario
                    If oFactServicioPagos.Eliminar(oDOFactOrdenServicioPagos) = True Then
                       oDoFacturacionServicioDesp.IdOrden = lnIdOrden
                       oDoFacturacionServicioDesp.IdUsuarioAuditoria = lnIdUsuario
                       If oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
                          oDoFactOrdenServ.IdOrden = lnIdOrden
                          oDoFactOrdenServ.IdUsuarioAuditoria = lnIdUsuario
                          If oFactOrdenesServicio.Eliminar(oDoFactOrdenServ) = True Then
                          End If
                       End If
                    End If
                End If
            End If
      End If
      With oDoFactOrdenServ
        .FechaCreacion = ldFechaCreacion
        .idCuentaAtencion = lnIdCuenta
        .IdServicioPaciente = lnIdServicioIngreso
        .IdEstadoFacturacion = 1               'Registrado
        .idPaciente = lnIdPaciente
        .IdPuntoCarga = lnIdPuntoCarga
        .IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal
        .idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente
        .idUsuario = lnIdUsuario
        .IdUsuarioAuditoria = lnIdUsuario
        .fechaDespacho = .FechaCreacion
        .IdUsuarioDespacho = lnIdUsuario
      End With
      If oFactOrdenesServicio.Insertar(oDoFactOrdenServ) = True Then
       With oDoFacturacionServicioDesp
          .cantidad = 1
          .IdOrden = oDoFactOrdenServ.IdOrden
          .idProducto = lnIdProductoEPS
          .IdUsuarioAuditoria = lnIdUsuario
          .Precio = lnImporteQpagara
          .total = lnImporteQpagara
       End With
       If oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
          With oDOFactOrdenServicioPagos
              .idOrdenPago = 0
              .FechaCreacion = oDoFactOrdenServ.FechaCreacion
              .IdEstadoFacturacion = 1   'Registrado
              .IdOrden = oDoFactOrdenServ.IdOrden
              .idUsuario = lnIdUsuario
              .IdUsuarioAuditoria = lnIdUsuario
          End With
          If oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) = True Then
                 With oDoFacturacionServicioPagos
                     .cantidad = 1
                     .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                     .idProducto = lnIdProductoEPS
                     .IdUsuarioAuditoria = lnIdUsuario
                     .Precio = lnImporteQpagara
                     .total = lnImporteQpagara
                 End With
                 If oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                    mo_idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                    EPScrearOrdenPago = True
                 End If
          End If
        End If
      End If
      Set oDoFactOrdenServ = Nothing
      Set oFactOrdenesServicio = Nothing
      Set oDoFacturacionServicioDesp = Nothing
      Set oFacturacionServicioDespacho = Nothing
      Set oDOFactOrdenServicioPagos = Nothing
      Set oFactServicioPagos = Nothing
      Set oDoFacturacionServicioPagos = Nothing
      Set oFacturacionServicioPagos = Nothing
      Exit Function
ErrEPScrearOP:

End Function


Function ModificaDatosVentaDirecta(oDoMovimiento As DoFarmMovimiento, oDoFarmMovimientoVentas As DoFarmMovimientoVentas, _
                                   oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                   mo_lcNombrePc As String, lnIdReceta As Long, lnEpsPorcentaje As Double, _
                                   lnIdOrdenPago As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoVentas As New farmMovimientoVentas
    Dim oMovimientoVentasDetalle As New FarmMovimientoVentasDetalle
    Dim oDoMovimientoVentasDetalle As New DoFarmMovimientoVentasDetalle
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinanciam As New DoFacturacionBienesFinanciam
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset, oRsTmp2 As New Recordset
    Dim oRsLookTablas As New ADODB.Recordset, oRsTmp1 As New Recordset
    Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
    Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
    Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
    Dim lbGeneraPago As Boolean
    Dim lcSql As String, lnImporteTotal As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    'Modificado por Yamill Palomino
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoVentas.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    Set oMovimientoVentasDetalle.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalleItem.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarMVTDTA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "M", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloVenta, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarMVTDTA
    End If
    
    '*********  graba tabla MovimientosVentas  ***************
    If oDoFarmMovimientoVentas.idCuentaAtencion > 0 Then
       oDoFarmMovimientoVentas.IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFarmMovimientoVentas.idCuentaAtencion, Format(oDoMovimiento.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), Format(oDoMovimiento.FechaCreacion, "hh:mm"))
    Else
       oDoFarmMovimientoVentas.IdServicioPaciente = 0
    End If
    If Not oMovimientoVentas.modificar(oDoFarmMovimientoVentas) Then
            bProcesoOK = False: GoTo TerminarMVTDTA
    End If
    '*********  Elimina tabla farmMovimientosDetalle y devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
           oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
           oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
           oDoMovimientoDetalle.idProducto = rs.Fields!idProducto
           oDoMovimientoDetalle.Lote = rs.Fields!Lote
           oDoMovimientoDetalle.FechaVencimiento = rs.Fields!FechaVencimiento
           oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
           oDoMovimientoDetalle.idTipoSalidaBienInsumo = rs.Fields!idTipoSalidaBienInsumo
           If Not FarmMovimientoDetalleEliminar(oConexion, oDoMovimientoDetalle) Then
              bProcesoOK = False: GoTo TerminarMVTDTA
           End If
           If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
              bProcesoOK = False: GoTo TerminarMVTDTA
           End If
           rs.MoveNext
        Loop
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarMVTDTA
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarMVTDTA
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '
    lbGeneraPago = ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFarmMovimientoVentas.IdTipoFinanciamiento, oConexion)
    Dim lbEsPlanAtencion As Boolean
    Dim rsTmp As New Recordset
    
    'Modificado por Yamill Palomino
        With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultaTiposFinanciamientoXidTipoFinanciamiento"
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, oDoFarmMovimientoVentas.IdTipoFinanciamiento): .Parameters.Append oParameter
        Set rsTmp = .Execute
        Set rsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
    
    
    If rsTmp.Fields!esFuenteFinanciamiento = True Then
       lbEsPlanAtencion = True
    Else
       lbEsPlanAtencion = False
       If rsTmp.Fields!IdTipoFinanciamiento = 9 Then  'Servicio Social genera datos en tablas de FACTURACION
          lbEsPlanAtencion = True
          lbGeneraPago = False
       End If
    End If
    rsTmp.Close
    Set rsTmp = Nothing
    '*********  elimina los productos de tabla FacturacionBienesPagos, ***************
    If lbEsPlanAtencion = True Then
        oDoFactOrdenesBienes.movNumero = oDoMovimiento.movNumero
        oDoFactOrdenesBienes.MovTipo = oDoMovimiento.MovTipo
        If Not oFactOrdenesBienes.SeleccionarPorMovNumero(oDoFactOrdenesBienes) Then
                bProcesoOK = False: GoTo TerminarMVTDTA
        End If
        If oDoFactOrdenesBienes.IdEstadoFacturacion = 1 Then
            oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
            If Not oFacturacionBienesPagos.Eliminar(oDoFacturacionBienesPagos) Then
                     bProcesoOK = False: GoTo TerminarMVTDTA
            End If
        End If
    End If
    lnImporteTotal = 0
    '*********  graba tabla farmMovimientosVentasDetalle, FacturacionBienesFinanciamientos ***************
    oDoFacturacionBienesFinanciam.movNumero = oDoMovimiento.movNumero
    oDoFacturacionBienesFinanciam.MovTipo = oDoMovimiento.MovTipo
    If Not oFacturacionBienesFinanciam.Eliminar(oDoFacturacionBienesFinanciam) Then
       bProcesoOK = False: GoTo TerminarMVTDTA
    End If
    oDoFacturacionBienesFinanciam.IdTipoFinanciamiento = oDoFarmMovimientoVentas.IdTipoFinanciamiento
    oDoFacturacionBienesFinanciam.FechaAutoriza = oDoMovimiento.FechaCreacion
    oDoFacturacionBienesFinanciam.IdUsuarioAuditoria = oDoMovimiento.idUsuario
    oDoFacturacionBienesFinanciam.IdUsuarioAutoriza = oDoMovimiento.idUsuario
    oDoFacturacionBienesFinanciam.idFuenteFinanciamiento = oDoFarmMovimientoVentas.idFuenteFinanciamiento
    oDoMovimientoVentasDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoVentasDetalle.MovTipo = oDoMovimiento.MovTipo
    oDoMovimientoVentasDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    If Not oMovimientoVentasDetalle.Eliminar(oDoMovimientoVentasDetalle) Then
       bProcesoOK = False: GoTo TerminarMVTDTA
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoVentasDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoVentasDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoVentasDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoVentasDetalle.total = oRsDetalleProductos.Fields!total
       If Not oMovimientoVentasDetalle.Insertar(oDoMovimientoVentasDetalle) Then
          bProcesoOK = False: GoTo TerminarMVTDTA
       End If
       If lbGeneraPago = False And lbEsPlanAtencion = True Then
          oDoFacturacionBienesFinanciam.CantidadFinanciada = oRsDetalleProductos.Fields!cantidad
          oDoFacturacionBienesFinanciam.idProducto = oRsDetalleProductos.Fields!idProducto
          oDoFacturacionBienesFinanciam.PrecioFinanciado = oRsDetalleProductos.Fields!precioDelSeguro
          oDoFacturacionBienesFinanciam.TotalFinanciado = Round(oRsDetalleProductos.Fields!precioDelSeguro * oRsDetalleProductos.Fields!cantidad, 2)
          If Not oFacturacionBienesFinanciam.Insertar(oDoFacturacionBienesFinanciam) Then
               bProcesoOK = False: GoTo TerminarMVTDTA
          End If
          lnImporteTotal = lnImporteTotal + oDoFacturacionBienesFinanciam.TotalFinanciado
       End If
       '*********  graba tabla FacturacionBienesPagos ,solo si idtipofinanciamiento=5 (cred.hospit) ***************
       If lbGeneraPago = True And lbEsPlanAtencion = True Then
            If oDoFactOrdenesBienes.IdEstadoFacturacion = 1 Then
                oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
                oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFactOrdenesBienes.idUsuario
                oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!cantidad
                oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
                oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!Precio
                oDoFacturacionBienesPagos.TotalPagar = oRsDetalleProductos.Fields!total
                If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
                         bProcesoOK = False: GoTo TerminarMVTDTA
                End If
            End If
       End If
       oRsDetalleProductos.MoveNext
    Loop
    'Receta automatica debb-03/04/2012
    If lnIdReceta > 0 Then
        mo_ReglasImagenes.RecetaAutomaticaCantidadDespachadaCERO lnIdReceta, oConexion
        Set oRecetaDetalle.Conexion = oConexion
        Set oRsTmp1 = oRecetaDetalle.SeleccionarPorIdReceta(lnIdReceta)
        
        oRsDetalleProductos.MoveFirst
        Do While Not oRsDetalleProductos.EOF
           oRsTmp1.MoveFirst
           oRsTmp1.Find "idItem=" & oRsDetalleProductos.Fields!idProducto
           If Not oRsTmp1.EOF Then
                 oDoRecetaDetalle.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
           Else
                 oDoRecetaDetalle.CantidadDespachada = 0
           End If
           oDoRecetaDetalle.idReceta = lnIdReceta
           oDoRecetaDetalle.idItem = oRsDetalleProductos.Fields!idProducto
           oDoRecetaDetalle.idEstadoDetalle = 2
           If oRecetaDetalle.RecetaDetalleActualizaCantDespachada(oDoRecetaDetalle) = False Then
                bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarMVTDTA
           End If
           oDoRecetaDetalleItem.idReceta = lnIdReceta
           oDoRecetaDetalleItem.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
           oDoRecetaDetalleItem.DocumentoDespacho = oDoMovimiento.DocumentoNumero
           oDoRecetaDetalleItem.idItem = oRsDetalleProductos.Fields!idProducto
           oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
           If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
               bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarMVTDTA
           End If
           oRsDetalleProductos.MoveNext
        Loop

    End If
    Set oRsTmp1 = Nothing
    '
    AsignaPorcentajeEPS oConexion, oDoFarmMovimientoVentas.idCuentaAtencion, lnEpsPorcentaje, mo_lnIdTablaLISTBARITEMS, _
                         sghAgregar, oDoMovimiento.FechaCreacion, _
                         oDoFarmMovimientoVentas.IdServicioPaciente, oDoFarmMovimientoVentas.idPaciente, _
                         oDoFarmMovimientoVentas.IdUsuarioAuditoria, oDoFarmMovimientoVentas.movNumero, lnImporteTotal, _
                         sghPuntosCargaBasicos.sghPtoCargaFarmacia, oDoFarmMovimientoVentas.MovTipo, lnIdOrdenPago
    
TerminarMVTDTA:
    If bProcesoOK Then
        ModificaDatosVentaDirecta = True
        oConexion.CommitTrans
    Else
        ModificaDatosVentaDirecta = False
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentas.MensajeError <> "", oMovimientoVentas.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentasDetalle.MensajeError <> "", oMovimientoVentasDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesFinanciam.MensajeError <> "", oFacturacionBienesFinanciam.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoVentas = Nothing
    Set oMovimientoVentasDetalle = Nothing
    Set oDoMovimientoVentasDetalle = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    Set oFacturacionBienesPagos = Nothing
    Set rs = Nothing
    Set oRsLookTablas = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRecetaCabecera = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaDetalle = Nothing
    Set mo_ReglasImagenes = Nothing
    Set oRecetaDetalleItem = Nothing
    Set oDoRecetaDetalleItem = Nothing
End Function


Function farmMovimientoVentasSeleccionarPorId(ByVal oDoMovimientoVentas As DoFarmMovimientoVentas, Optional oConexion1 As Connection) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimientoVtas As New farmMovimientoVentas
    
    oConexion.Open sighentidades.CadenaConexion
    If oConexion1 Is Nothing Then
       Set oMovimientoVtas.Conexion = oConexion
    Else
       Set oMovimientoVtas.Conexion = oConexion1
    End If
    If oMovimientoVtas.SeleccionarPorId(oDoMovimientoVentas) Then
        farmMovimientoVentasSeleccionarPorId = True
    Else
        farmMovimientoVentasSeleccionarPorId = False
        ms_MensajeError = oMovimientoVtas.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientoVtas = Nothing
End Function
'********************************   venta Directa ***************************************

'********************************   Preventa  ***************************************
Function AgregaDatosDePreVenta(ByVal oDoPreventa As DoFarmPreVenta, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdReceta As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oPreVenta As New FarmPreVenta
    Dim oPreVentaDetalle As New FarmPreVentaDetalle
    Dim oDoPreventaDetalle As New DoFarmPreVentaDetalle
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    Set oPreVenta.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oPreVentaDetalle.Conexion = oConexion
    '*********** revisa nuevamente los saldos *************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoPreventa.IdAlmacen, oRsDetalleProductos, oDoPreventa.FechaCreacion, _
                                      sghTipoSalidaItemFarmacia.sghSoloVenta, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarIPV
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
        ms_MensajeError = ms_MensajeError + "No existe Productos"
        bProcesoOK = False: GoTo TerminarIPV
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
        If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoPreventa.IdAlmacen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad) Then
              ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
              bProcesoOK = False: GoTo TerminarIPV
        End If
        oRsDetalleProductosConLotes.MoveNext
    Loop
    
    '*********  graba tabla PreVentas  ***************
   
    If Not oPreVenta.Insertar(oDoPreventa) Then
      bProcesoOK = False: GoTo TerminarIPV
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPreventa.IdUsuarioAuditoria, "A", oDoPreventa.idPreVenta, "FarmPreVenta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '*********  graba tabla FactOrdenesBienes  ***************
    oDoFactOrdenesBienes.FechaCreacion = oDoPreventa.FechaCreacion
    oDoFactOrdenesBienes.idCuentaAtencion = oDoPreventa.idCuentaAtencion
    oDoFactOrdenesBienes.IdEstadoFacturacion = 1 'registrado
    oDoFactOrdenesBienes.idPaciente = oDoPreventa.idPaciente
    oDoFactOrdenesBienes.idPreVenta = oDoPreventa.idPreVenta
    oDoFactOrdenesBienes.IdPuntoCarga = 5  'farmacia
    oDoFactOrdenesBienes.idUsuario = oDoPreventa.idUsuario
    oDoFactOrdenesBienes.IdUsuarioAuditoria = oDoPreventa.idUsuario
    If Not oFactOrdenesBienes.Insertar(oDoFactOrdenesBienes) Then
      bProcesoOK = False: GoTo TerminarIPV
    End If
    
    oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
    oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFactOrdenesBienes.idUsuario
    
    oDoPreventaDetalle.idPreVenta = oDoPreventa.idPreVenta
    oDoPreventaDetalle.IdUsuarioAuditoria = oDoPreventa.IdUsuarioAuditoria
    lnItem = 1
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
        oDoPreventaDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
        oDoPreventaDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
        oDoPreventaDetalle.Importe = oRsDetalleProductos.Fields!total
        oDoPreventaDetalle.Item = lnItem
        oDoPreventaDetalle.Precio = oRsDetalleProductos.Fields!Precio
        If Not oPreVentaDetalle.Insertar(oDoPreventaDetalle) Then
                 bProcesoOK = False: GoTo TerminarIPV
        End If
        oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!cantidad
        oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
        oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!Precio
        oDoFacturacionBienesPagos.TotalPagar = oRsDetalleProductos.Fields!total
        If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
                 bProcesoOK = False: GoTo TerminarIPV
        End If
        lnItem = lnItem + 1
        oRsDetalleProductos.MoveNext
    Loop
    
    'Receta automatica
    'debb-18/05/2016 (inicio)
    If lnIdReceta > 0 Then
        Dim oDoRecetaDetalleItem As New DoRecetaDetalleItem
        Dim oRecetaDetalleItem As New RecetaDetalleItem
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RecetaCabeceraModificaIdEstado"
            Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 20, Trim(Str(oDoPreventa.idPreVenta))): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
             .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oRecetaDetalleItem.Conexion = oConexion
        oRsDetalleProductos.MoveFirst
        Do While Not oRsDetalleProductos.EOF
           oDoRecetaDetalleItem.idReceta = lnIdReceta
           oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoPreventa.IdUsuarioAuditoria
           oDoRecetaDetalleItem.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
           oDoRecetaDetalleItem.DocumentoDespacho = oDoPreventa.idPreVenta
           oDoRecetaDetalleItem.IdComprobantePago = 0
           oDoRecetaDetalleItem.idItem = oRsDetalleProductos.Fields!idProducto
           oDoRecetaDetalleItem.IdEstado = sghRecetaEstados.sighRecetaDespachada
           If Not oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) Then
              bProcesoOK = False: GoTo TerminarIPV
           End If
           oRsDetalleProductos.MoveNext
        Loop
        Set oDoRecetaDetalleItem = Nothing
        Set oRecetaDetalleItem = Nothing
    End If
    'debb-18/05/2016 (fin)
    Set oRsTmp1 = Nothing
   
    '
TerminarIPV:
    If bProcesoOK Then
  AgregaDatosDePreVenta = True
  oConexion.CommitTrans
    Else
  AgregaDatosDePreVenta = False
  ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oPreVenta.MensajeError <> "", oPreVenta.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oPreVentaDetalle.MensajeError <> "", oPreVentaDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesPagos.MensajeError <> "", oFacturacionBienesPagos.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    
    Set oConexion = Nothing
    Set oPreVenta = Nothing
    Set oPreVentaDetalle = Nothing
    Set oDoPreventaDetalle = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oDoFacturacionBienesPagos = Nothing

    Set oConexion = Nothing
    Set oCommand = Nothing
End Function
Function ModificaDatosDePreVenta(ByVal oDoPreventa As DoFarmPreVenta, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oPreVenta As New FarmPreVenta
    Dim oPreVentaDetalle As New FarmPreVentaDetalle
    Dim oDoPreventaDetalle As New DoFarmPreVentaDetalle
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oPreVenta.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oPreVentaDetalle.Conexion = oConexion
    '*********** revisa nuevamente los saldos *************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoPreventa.IdAlmacen, oRsDetalleProductos, oDoPreventa.FechaCreacion, _
                                            sghTipoSalidaItemFarmacia.sghSoloVenta, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarMPTA
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
        ms_MensajeError = ms_MensajeError + "No existe Productos"
        bProcesoOK = False: GoTo TerminarMPTA
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoPreventa.IdAlmacen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarMPTA
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '*********  graba tabla PreVentas  ***************
    If Not oPreVenta.modificar(oDoPreventa) Then
            bProcesoOK = False: GoTo TerminarMPTA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPreventa.IdUsuarioAuditoria, "M", oDoPreventa.idPreVenta, "FarmPreVenta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '*********  graba tabla FactOrdenesBienes  ***************
    
    oDoFactOrdenesBienes.idPreVenta = oDoPreventa.idPreVenta
    If Not oFactOrdenesBienes.SeleccionarPorIdpreventa(oDoFactOrdenesBienes) Then
       bProcesoOK = False: GoTo TerminarMPTA
    End If
    oDoFactOrdenesBienes.idCuentaAtencion = oDoPreventa.idCuentaAtencion
    oDoFactOrdenesBienes.idPaciente = oDoPreventa.idPaciente
    oDoFactOrdenesBienes.IdUsuarioAuditoria = oDoPreventa.idUsuario
    If Not oFactOrdenesBienes.modificar(oDoFactOrdenesBienes) Then
            bProcesoOK = False: GoTo TerminarMPTA
    End If

    '*********  Elimina productos tabla FacturacionBienesPagos  ***************
    
    oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
    oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFactOrdenesBienes.idUsuario
    If Not oFacturacionBienesPagos.Eliminar(oDoFacturacionBienesPagos) Then
           bProcesoOK = False: GoTo TerminarMPTA
    End If
    '*********  Elimina productos tabla farmPreventaDetalle  ***************
    oDoPreventaDetalle.idPreVenta = oDoPreventa.idPreVenta
    If Not FarmPreVentaEliminaTodosItems(oDoPreventaDetalle) Then
           bProcesoOK = False: GoTo TerminarMPTA
    End If
   '*********  graba tabla PreVentaDetalle,farmSaldo,farmSaldoDetalle,FacturacionBienesPagos    ***************
    
    oDoPreventaDetalle.idPreVenta = oDoPreventa.idPreVenta
    oDoPreventaDetalle.IdUsuarioAuditoria = oDoPreventa.IdUsuarioAuditoria
    lnItem = 1
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoPreventaDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoPreventaDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoPreventaDetalle.Importe = oRsDetalleProductos.Fields!total
       oDoPreventaDetalle.Item = lnItem
       oDoPreventaDetalle.Precio = oRsDetalleProductos.Fields!Precio
       If Not oPreVentaDetalle.Insertar(oDoPreventaDetalle) Then
                bProcesoOK = False: GoTo TerminarMPTA
       End If
       oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!cantidad
       oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!Precio
       oDoFacturacionBienesPagos.TotalPagar = oRsDetalleProductos.Fields!total
       If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
                bProcesoOK = False: GoTo TerminarMPTA
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarMPTA:
    If bProcesoOK Then
        ModificaDatosDePreVenta = True
        oConexion.CommitTrans
    Else
        ModificaDatosDePreVenta = False
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oPreVenta.MensajeError <> "", oPreVenta.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oPreVentaDetalle.MensajeError <> "", oPreVentaDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesPagos.MensajeError <> "", oFacturacionBienesPagos.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPreVenta = Nothing
    Set oPreVentaDetalle = Nothing
    Set oDoPreventaDetalle = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oDoFacturacionBienesPagos = Nothing
    Set oFacturacionBienesPagos = Nothing
    Set oRsLookTablas = Nothing
End Function
Function AnulaPreVenta(oDoPreventa As DoFarmPreVenta, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdReceta As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oPreVenta As New FarmPreVenta
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oPreVenta.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    
    '*********  graba tabla PreVentas  ***************
   
    If Not oPreVenta.modificar(oDoPreventa) Then
            bProcesoOK = False: GoTo TerminarAPVTA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoPreventa.IdUsuarioAuditoria, "E", oDoPreventa.idPreVenta, "FarmPreVenta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '*********  graba tabla FactOrdenesBienes  ***************
    
    oDoFactOrdenesBienes.idPreVenta = oDoPreventa.idPreVenta
    If Not oFactOrdenesBienes.SeleccionarPorIdpreventa(oDoFactOrdenesBienes) Then
       bProcesoOK = False: GoTo TerminarAPVTA
    End If
    oDoFactOrdenesBienes.IdEstadoFacturacion = 9        'anulado
    oDoFactOrdenesBienes.IdUsuarioAuditoria = oDoPreventa.idUsuario
    If Not oFactOrdenesBienes.modificar(oDoFactOrdenesBienes) Then
            bProcesoOK = False: GoTo TerminarAPVTA
    End If
    'Receta automatica debb-03/04/2012
    If lnIdReceta > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalle lnIdReceta, oDoPreventa.idCuentaAtencion, 0, oConexion
    End If
    Set oRsTmp1 = Nothing
    '
TerminarAPVTA:
    If bProcesoOK Then
        AnulaPreVenta = True
        oConexion.CommitTrans
    Else
        AnulaPreVenta = False
        ms_MensajeError = ms_MensajeError + IIf(oPreVenta.MensajeError <> "", oPreVenta.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPreVenta = Nothing
    Set oRsLookTablas = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oCommand = Nothing
    Set mo_ReglasImagenes = Nothing
End Function
Function FarmPreventasSeleccionarPorId(ByVal oDoFarmPreVenta As DoFarmPreVenta) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oPreVenta As New FarmPreVenta
    oConexion.Open sighentidades.CadenaConexion
    Set oPreVenta.Conexion = oConexion
    If oPreVenta.SeleccionarPorId(oDoFarmPreVenta) Then
        FarmPreventasSeleccionarPorId = True
    Else
        FarmPreventasSeleccionarPorId = False
        ms_MensajeError = oPreVenta.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPreVenta = Nothing
End Function
Function FarmPreVentaDetalleDevuelveTodosItems(lnIdPreventa As Long) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New FarmPreVentaDetalle
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set FarmPreVentaDetalleDevuelveTodosItems = oTabla.DevuelveTodosItems(lnIdPreventa)
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function FarmPreVentaEliminaTodosItems(oDoPreventaDetalle As DoFarmPreVentaDetalle) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New FarmPreVentaDetalle
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    FarmPreVentaEliminaTodosItems = oTabla.Eliminar(oDoPreventaDetalle)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function

'********************************   Preventa  ***************************************

'********************************   Intervencion Sanitaria  ***************************************
Function AgregaDatosDeIntervencion(ByVal oDoMovimiento As DoFarmMovimiento, _
                                   oDoMovimientoProgramas As DOfarmMovimientoProgramas, _
                                   oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                   mo_lcNombrePc As String, wxParametro347 As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoProgramas As New farmMovimientoProgramas
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String: Dim lcCorrelativoPrograma As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoProgramas.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    'SCCQ 22/10/2020 Cambio28 Inicio
    'Antes:lcCorrelativoPrograma = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(23)   'Guia Interna de Salida
    lcCorrelativoPrograma = oMovimiento.DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos(23)   'Guia Interna de Salida
    'SCCQ 22/10/2020 Cambio28 Fin
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
       .DocumentoNumero = lcCorrelativoPrograma
       .DocumentoIdtipo = 23   'Guia Interna de Salida
    End With
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarIS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloEstrategico, _
                                            wxParametro347) Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarIS
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe ITEMS" + Chr(13): GoTo TerminarIS
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
                                                  oRsDetalleProductosConLotes.Fields!idProducto, _
                                                  oRsDetalleProductosConLotes.Fields!Lote, _
                                                  oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
                                                  oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                                  oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarIS
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '*********  graba tabla MovimientosProgramas  ***************
    With oDoMovimientoProgramas
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimientoProgramas.Insertar(oDoMovimientoProgramas) Then
            bProcesoOK = False: GoTo TerminarIS
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarIS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, _
                                                            oRsDetalleProductosConLotes.Fields!idProducto, _
                                                            oRsDetalleProductosConLotes.Fields!Lote, _
                                                            oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
                                                            oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                                            oRsDetalleProductosConLotes.Fields!cantidad, _
                                                            oRsDetalleProductosConLotes.Fields!Precio, _
                                                            oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarIS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
    
TerminarIS:
    If bProcesoOK Then
        AgregaDatosDeIntervencion = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeIntervencion = False
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoProgramas = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function
Function ModificaDatosDeIntervencion(oDoMovimiento As DoFarmMovimiento, oDoMovimientoProgramas As DOfarmMovimientoProgramas, _
                                     oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                     mo_lcNombrePc As String, wxParametro347 As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oMovimientoProgramas As New farmMovimientoProgramas
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    Dim oRsLookTablas As New ADODB.Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoProgramas.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarMIS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "M", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloEstrategico, _
                                            wxParametro347) Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarMIS
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
       ms_MensajeError = "Falta ITEMS": bProcesoOK = False: GoTo TerminarMIS
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
                                                  oRsDetalleProductosConLotes.Fields!idProducto, _
                                                  oRsDetalleProductosConLotes.Fields!Lote, _
                                                  oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
                                                  oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                                  oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarMIS
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    
   '*********  graba tabla MovimientosProgramas  ***************
    If Not oMovimientoProgramas.modificar(oDoMovimientoProgramas) Then
            bProcesoOK = False: GoTo TerminarMIS
    End If
    '*********  Elimina tabla farmMovimientosDetalle y devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
           oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
           oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
           oDoMovimientoDetalle.idProducto = rs.Fields!idProducto
           oDoMovimientoDetalle.Lote = rs.Fields!Lote
           oDoMovimientoDetalle.FechaVencimiento = rs.Fields!FechaVencimiento
           oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
           oDoMovimientoDetalle.idTipoSalidaBienInsumo = rs.Fields!idTipoSalidaBienInsumo
           If Not oMovimientoDetalle.Eliminar(oDoMovimientoDetalle) Then
              bProcesoOK = False: GoTo TerminarMIS
           End If
           If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, _
                                                                rs.Fields!idProducto, rs.Fields!Lote, _
                                                                rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, _
                                                                rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
              bProcesoOK = False: GoTo TerminarMIS
           End If
           rs.MoveNext
        Loop
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarMIS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, _
                                                            oRsDetalleProductosConLotes.Fields!idProducto, _
                                                            oRsDetalleProductosConLotes.Fields!Lote, _
                                                            oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
                                                            oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                                            oRsDetalleProductosConLotes.Fields!cantidad, _
                                                            oRsDetalleProductosConLotes.Fields!Precio, _
                                                            oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarMIS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
TerminarMIS:
    If bProcesoOK Then
        ModificaDatosDeIntervencion = True
        oConexion.CommitTrans
    Else
        ModificaDatosDeIntervencion = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oMovimientoProgramas = Nothing
    Set rs = Nothing
    Set oRsLookTablas = Nothing
End Function
Function FarmMovimientoProgramaSeleccionarPorId(ByVal oDoMovimientoPrograma As DOfarmMovimientoProgramas) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimientoProg As New farmMovimientoProgramas
    oConexion.Open sighentidades.CadenaConexion
    Set oMovimientoProg.Conexion = oConexion
    If oMovimientoProg.SeleccionarPorId(oDoMovimientoPrograma) Then
        FarmMovimientoProgramaSeleccionarPorId = True
    Else
        FarmMovimientoProgramaSeleccionarPorId = False
        ms_MensajeError = oMovimientoProg.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientoProg = Nothing
End Function







'********************************   Nota de Ingreso ***************************************
Function AgregaDatosDeNotaIngreso(oDoMovimiento As DoFarmMovimiento, oDoMovimientoNotaIngreso As DOfarmMovimientoNotaIngreso, oDoProveedores As DoProveedores, oRsDetalleProductos As ADODB.Recordset, lnIdtipoFinanciamiento As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oDOCatalogoBienes As New DOCatalogoBienesInsumos
    Dim oCatalogoBienes As New CatalogoBienesInsumos
    Dim oProveedores As New Proveedores
    Dim oDoFacturacionBienesDevol As New DoFacturacionBienesDevol
    Dim oFacturacionBienesDevol As New FacturacionBienesDevol
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesFinan As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinan As New DoFacturacionBienesFinanciam
    Dim oRsDevoluciones As New Recordset
    Dim lcCorrelativo As String, lcSql As String
    Dim lnItem As Long, lnCantidadPagar As Long, lnIdOrden As Long
    Dim lnTotalPagar As Double
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oProveedores.Conexion = oConexion
    Set oMovimientoNotaIngreso.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oCatalogoBienes.Conexion = oConexion
    Set oFacturacionBienesDevol.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oFacturacionBienesFinan.Conexion = oConexion
    
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(1)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: ms_MensajeError = oMovimiento.MensajeError: GoTo TerminarNI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '************** Graba nuevo proveedor
    If oDoProveedores.RazonSocial <> "" Then
        
        If Not oProveedores.Insertar(oDoProveedores) Then
           bProcesoOK = False: ms_MensajeError = oProveedores.MensajeError: GoTo TerminarNI
        End If
        oDoMovimientoNotaIngreso.idProveedor = oDoProveedores.idProveedor
    End If
    '*********  graba tabla MovimientosNotaIngreso  ***************
    With oDoMovimientoNotaIngreso
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimientoNotaIngreso.Insertar(oDoMovimientoNotaIngreso) Then
            bProcesoOK = False: ms_MensajeError = oMovimientoNotaIngreso.MensajeError: GoTo TerminarNI
    End If
    
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle, actualiza Precio de Ultima Compra  ***************
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe productos": GoTo TerminarNI
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = IIf(IsNull(oRsDetalleProductos.Fields!RegistroSanitario), "", oRsDetalleProductos.Fields!RegistroSanitario)
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       oDoMovimientoDetalle.DocumentoNumero = IIf(IsNull(oRsDetalleProductos.Fields!NumeroDocumento), "", Trim(oRsDetalleProductos.Fields!NumeroDocumento))
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: ms_MensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNI
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenDestino, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: ms_MensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNI
       End If
       '*********  graba Devoluciones  de Pacientes***************
       If oDoMovimiento.idTipoConcepto = 21 And oDoMovimientoNotaIngreso.idCuentaAtencion > 0 Then
            With oDoFacturacionBienesDevol
                .cantidadAdevolver = oRsDetalleProductos.Fields!cantidad
                .FechaAutoriza = oDoMovimiento.FechaCreacion
                .IdComprobantePago = 0
                .idEstadoDevolucion = 2  'devuelto
                .idProducto = oRsDetalleProductos.Fields!idProducto
                .IdUsuarioAuditoria = oDoMovimiento.idUsuario
                .IdUsuarioAutoriza = oDoMovimiento.idUsuario
                .movNumero = oRsDetalleProductos.Fields!movNumeroS
                .MovTipo = "S"
                .MovNumeroE = oDoMovimiento.movNumero
                .MovTipoE = oDoMovimiento.MovTipo
            End With
            If Not oFacturacionBienesDevol.Insertar(oDoFacturacionBienesDevol) Then
                     bProcesoOK = False: ms_MensajeError = oFacturacionBienesDevol.MensajeError: GoTo TerminarNI
            End If
            If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                    oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!CantidadS
                    oDoFacturacionBienesPagos.IdOrden = oRsDetalleProductos.Fields!IdOrdenS
                    oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFacturacionBienesDevol.IdUsuarioAuditoria
                    oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!PrecioS
                    lnCantidadPagar = oDoFacturacionBienesPagos.CantidadPagar - oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oDoFacturacionBienesPagos.PrecioVenta, 2)
                    oDoFacturacionBienesPagos.CantidadPagar = lnCantidadPagar
                    oDoFacturacionBienesPagos.TotalPagar = lnTotalPagar
                    If Not oFacturacionBienesPagos.ModificarXidOrdenIdProducto(oDoFacturacionBienesPagos) Then
                       bProcesoOK = False: ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo TerminarNI
                    End If
            Else
                    lnCantidadPagar = oRsDetalleProductos.Fields!CantidadS - oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oRsDetalleProductos.Fields!PrecioS, 2)
                    oDoFacturacionBienesFinan.movNumero = oDoFacturacionBienesDevol.movNumero
                    oDoFacturacionBienesFinan.MovTipo = oDoFacturacionBienesDevol.MovTipo
                    oDoFacturacionBienesFinan.idProducto = oDoFacturacionBienesDevol.idProducto
                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = oRsDetalleProductos.Fields!IdTipoFinanciamientoS
                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = oRsDetalleProductos.Fields!IdFuenteFinanciamientoS
                    oDoFacturacionBienesFinan.FechaAutoriza = oRsDetalleProductos.Fields!FechaS
                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = oRsDetalleProductos.Fields!usuarioS
                    oDoFacturacionBienesFinan.PrecioFinanciado = oRsDetalleProductos.Fields!PrecioS
                    oDoFacturacionBienesFinan.TotalFinanciado = lnTotalPagar
                    oDoFacturacionBienesFinan.CantidadFinanciada = lnCantidadPagar
                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                    If Not oFacturacionBienesFinan.modificar(oDoFacturacionBienesFinan) Then
                       bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo TerminarNI
                    End If
                End If
       End If
       If oDoMovimiento.idTipoConcepto = 1 Or oDoMovimiento.idTipoConcepto = 2 Then
            '********* solo actualiza PRECIO ULTIMA COMPRA para CONCEPTOS=Compra/Encargo
            oDOCatalogoBienes.idProducto = oRsDetalleProductos.Fields!idProducto
            If Not oCatalogoBienes.SeleccionarPorId(oDOCatalogoBienes) Then
               bProcesoOK = False: ms_MensajeError = oCatalogoBienes.MensajeError: GoTo TerminarNI
            End If
            oDOCatalogoBienes.PrecioUltCompra = oRsDetalleProductos.Fields!Precio
            If Not oCatalogoBienes.modificar(oDOCatalogoBienes) Then
               bProcesoOK = False: ms_MensajeError = oCatalogoBienes.MensajeError: GoTo TerminarNI
            End If
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
    
TerminarNI:
    If bProcesoOK Then
        AgregaDatosDeNotaIngreso = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaIngreso = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoNotaIngreso.MensajeError <> "", oMovimientoNotaIngreso.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesDevol.MensajeError <> "", oFacturacionBienesDevol.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesFinan.MensajeError <> "", oFacturacionBienesFinan.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesPagos.MensajeError <> "", oFacturacionBienesPagos.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoNotaIngreso = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oDOCatalogoBienes = Nothing
    Set oProveedores = Nothing
    Set oFacturacionBienesDevol = Nothing
    Set oDoFacturacionBienesDevol = Nothing
End Function
'SCCQ 19/10/2020 Cambio28 Inicio
'Se copió el código de la Función AgregaDatosDeNotaIngreso y se agregó lo necesario para generar y actualizar correlativo disponible de número de documento.
Function AgregaDatosDeNotaIngreso_NumDocAutomatico(lnIdTipoLocales As String, lnIdTipoSuministro As String, lnDocumentoIdtipo As Long, oDoMovimiento As DoFarmMovimiento, oDoMovimientoNotaIngreso As DOfarmMovimientoNotaIngreso, oDoProveedores As DoProveedores, oRsDetalleProductos As ADODB.Recordset, lnIdtipoFinanciamiento As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oDOCatalogoBienes As New DOCatalogoBienesInsumos
    Dim oCatalogoBienes As New CatalogoBienesInsumos
    Dim oProveedores As New Proveedores
    Dim oDoFacturacionBienesDevol As New DoFacturacionBienesDevol
    Dim oFacturacionBienesDevol As New FacturacionBienesDevol
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesFinan As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinan As New DoFacturacionBienesFinanciam
    Dim oRsDevoluciones As New Recordset
    Dim lcCorrelativo As String, lcSql As String
    'SCCQ 19/10/2020 Cambio28 Inicio
    Dim lcCorrelativoDocumN As String
    'SCCQ 19/10/2020 Cambio28 Fin
    Dim lnItem As Long, lnCantidadPagar As Long, lnIdOrden As Long
    Dim lnTotalPagar As Double
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oProveedores.Conexion = oConexion
    Set oMovimientoNotaIngreso.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oCatalogoBienes.Conexion = oConexion
    Set oFacturacionBienesDevol.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oFacturacionBienesFinan.Conexion = oConexion
    
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(1)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    'SCCQ 19/10/2020 Cambio28 Inicio
    lcCorrelativoDocumN = oMovimiento.DevuelveYactualizaCorrelativoDisponibleFarmRelMod("E", lnIdTipoLocales, lnIdTipoSuministro, lnDocumentoIdtipo) 'GENERA Y ACTUALIZA CORRELATIVO
    oDoMovimiento.DocumentoNumero = lcCorrelativoDocumN
    'SCCQ 19/10/2020 Cambio28 Fin
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: ms_MensajeError = oMovimiento.MensajeError: GoTo TerminarNI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '************** Graba nuevo proveedor
    If oDoProveedores.RazonSocial <> "" Then
        
        If Not oProveedores.Insertar(oDoProveedores) Then
           bProcesoOK = False: ms_MensajeError = oProveedores.MensajeError: GoTo TerminarNI
        End If
        oDoMovimientoNotaIngreso.idProveedor = oDoProveedores.idProveedor
    End If
    '*********  graba tabla MovimientosNotaIngreso  ***************
    With oDoMovimientoNotaIngreso
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimientoNotaIngreso.Insertar(oDoMovimientoNotaIngreso) Then
            bProcesoOK = False: ms_MensajeError = oMovimientoNotaIngreso.MensajeError: GoTo TerminarNI
    End If
    
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle, actualiza Precio de Ultima Compra  ***************
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe productos": GoTo TerminarNI
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = IIf(IsNull(oRsDetalleProductos.Fields!RegistroSanitario), "", oRsDetalleProductos.Fields!RegistroSanitario)
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       oDoMovimientoDetalle.DocumentoNumero = IIf(IsNull(oRsDetalleProductos.Fields!NumeroDocumento), "", Trim(oRsDetalleProductos.Fields!NumeroDocumento))
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: ms_MensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNI
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenDestino, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: ms_MensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNI
       End If
       '*********  graba Devoluciones  de Pacientes***************
       If oDoMovimiento.idTipoConcepto = 21 And oDoMovimientoNotaIngreso.idCuentaAtencion > 0 Then
            With oDoFacturacionBienesDevol
                .cantidadAdevolver = oRsDetalleProductos.Fields!cantidad
                .FechaAutoriza = oDoMovimiento.FechaCreacion
                .IdComprobantePago = 0
                .idEstadoDevolucion = 2  'devuelto
                .idProducto = oRsDetalleProductos.Fields!idProducto
                .IdUsuarioAuditoria = oDoMovimiento.idUsuario
                .IdUsuarioAutoriza = oDoMovimiento.idUsuario
                .movNumero = oRsDetalleProductos.Fields!movNumeroS
                .MovTipo = "S"
                .MovNumeroE = oDoMovimiento.movNumero
                .MovTipoE = oDoMovimiento.MovTipo
            End With
            If Not oFacturacionBienesDevol.Insertar(oDoFacturacionBienesDevol) Then
                     bProcesoOK = False: ms_MensajeError = oFacturacionBienesDevol.MensajeError: GoTo TerminarNI
            End If
            If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                    oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!CantidadS
                    oDoFacturacionBienesPagos.IdOrden = oRsDetalleProductos.Fields!IdOrdenS
                    oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFacturacionBienesDevol.IdUsuarioAuditoria
                    oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!PrecioS
                    lnCantidadPagar = oDoFacturacionBienesPagos.CantidadPagar - oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oDoFacturacionBienesPagos.PrecioVenta, 2)
                    oDoFacturacionBienesPagos.CantidadPagar = lnCantidadPagar
                    oDoFacturacionBienesPagos.TotalPagar = lnTotalPagar
                    If Not oFacturacionBienesPagos.ModificarXidOrdenIdProducto(oDoFacturacionBienesPagos) Then
                       bProcesoOK = False: ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo TerminarNI
                    End If
            Else
                    lnCantidadPagar = oRsDetalleProductos.Fields!CantidadS - oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oRsDetalleProductos.Fields!PrecioS, 2)
                    oDoFacturacionBienesFinan.movNumero = oDoFacturacionBienesDevol.movNumero
                    oDoFacturacionBienesFinan.MovTipo = oDoFacturacionBienesDevol.MovTipo
                    oDoFacturacionBienesFinan.idProducto = oDoFacturacionBienesDevol.idProducto
                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = oRsDetalleProductos.Fields!IdTipoFinanciamientoS
                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = oRsDetalleProductos.Fields!IdFuenteFinanciamientoS
                    oDoFacturacionBienesFinan.FechaAutoriza = oRsDetalleProductos.Fields!FechaS
                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = oRsDetalleProductos.Fields!usuarioS
                    oDoFacturacionBienesFinan.PrecioFinanciado = oRsDetalleProductos.Fields!PrecioS
                    oDoFacturacionBienesFinan.TotalFinanciado = lnTotalPagar
                    oDoFacturacionBienesFinan.CantidadFinanciada = lnCantidadPagar
                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                    If Not oFacturacionBienesFinan.modificar(oDoFacturacionBienesFinan) Then
                       bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo TerminarNI
                    End If
                End If
       End If
       If oDoMovimiento.idTipoConcepto = 1 Or oDoMovimiento.idTipoConcepto = 2 Then
            '********* solo actualiza PRECIO ULTIMA COMPRA para CONCEPTOS=Compra/Encargo
            oDOCatalogoBienes.idProducto = oRsDetalleProductos.Fields!idProducto
            If Not oCatalogoBienes.SeleccionarPorId(oDOCatalogoBienes) Then
               bProcesoOK = False: ms_MensajeError = oCatalogoBienes.MensajeError: GoTo TerminarNI
            End If
            oDOCatalogoBienes.PrecioUltCompra = oRsDetalleProductos.Fields!Precio
            If Not oCatalogoBienes.modificar(oDOCatalogoBienes) Then
               bProcesoOK = False: ms_MensajeError = oCatalogoBienes.MensajeError: GoTo TerminarNI
            End If
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
    
TerminarNI:
    If bProcesoOK Then
        AgregaDatosDeNotaIngreso_NumDocAutomatico = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaIngreso_NumDocAutomatico = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoNotaIngreso.MensajeError <> "", oMovimientoNotaIngreso.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesDevol.MensajeError <> "", oFacturacionBienesDevol.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesFinan.MensajeError <> "", oFacturacionBienesFinan.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesPagos.MensajeError <> "", oFacturacionBienesPagos.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoNotaIngreso = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oDOCatalogoBienes = Nothing
    Set oProveedores = Nothing
    Set oFacturacionBienesDevol = Nothing
    Set oDoFacturacionBienesDevol = Nothing
End Function
'SCCQ 19/10/2020 Cambio28 Fin
Function FarmMovimientoNotaIngresoSeleccionarPorId(ByVal oDoMovimientoNotaIngreso As DOfarmMovimientoNotaIngreso) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New FarmMovimientoNotaIngreso
    oConexion.Open sighentidades.CadenaConexion
    Set oMovimiento.Conexion = oConexion
    If oMovimiento.SeleccionarPorId(oDoMovimientoNotaIngreso) Then
        FarmMovimientoNotaIngresoSeleccionarPorId = True
    Else
        FarmMovimientoNotaIngresoSeleccionarPorId = False
        ms_MensajeError = oMovimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
End Function
Function ModificaDatosDeNotaIngreso(oDoMovimiento As DoFarmMovimiento, oDoMovimientoNotaIngreso As DOfarmMovimientoNotaIngreso, oDoProveedores As DoProveedores, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oDOCatalogoBienes As New DOCatalogoBienesInsumos
    Dim oCatalogoBienes As New CatalogoBienesInsumos
    Dim oProveedores As New Proveedores
    Dim oDoFacturacionBienesDevol As New DoFacturacionBienesDevol
    Dim oFacturacionBienesDevol As New FacturacionBienesDevol
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oProveedores.Conexion = oConexion
    Set oMovimientoNotaIngreso.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oCatalogoBienes.Conexion = oConexion
    Set oFacturacionBienesDevol.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarMNI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "M", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '************** Graba nuevo proveedor
    If Trim(oDoProveedores.RazonSocial) <> "" Then
        If oDoProveedores.idProveedor = 0 Then
            
            If Not oProveedores.Insertar(oDoProveedores) Then
               bProcesoOK = False: GoTo TerminarMNI
            End If
            oDoMovimientoNotaIngreso.idProveedor = oDoProveedores.idProveedor
        End If
    End If
    '*********  graba tabla MovimientosNotaIngreso  ***************
    
    If Not oMovimientoNotaIngreso.modificar(oDoMovimientoNotaIngreso) Then
            bProcesoOK = False: GoTo TerminarMNI
    End If
    '*********  Elimina Devoluciones del Paciente
    With oDoFacturacionBienesDevol
         .IdUsuarioAuditoria = oDoMovimiento.idUsuario
         .MovNumeroE = oDoMovimiento.movNumero
         .MovTipoE = oDoMovimiento.MovTipo
    End With
    If Not oFacturacionBienesDevol.EliminarPorMovNumeroE(oDoFacturacionBienesDevol) Then
       bProcesoOK = False: GoTo TerminarMNI
    End If
    '*********  Elimina tabla farmMovimientosDetalle y devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    rs.MoveFirst
    Do While Not rs.EOF
       oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
       oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
       oDoMovimientoDetalle.idProducto = rs.Fields!idProducto
       oDoMovimientoDetalle.Lote = rs.Fields!Lote
       oDoMovimientoDetalle.FechaVencimiento = rs.Fields!FechaVencimiento
       oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = rs.Fields!idTipoSalidaBienInsumo
       If Not FarmMovimientoDetalleEliminar(oConexion, oDoMovimientoDetalle) Then
          bProcesoOK = False: GoTo TerminarMNI
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenDestino, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
          bProcesoOK = False: GoTo TerminarMNI
       End If
       rs.MoveNext
    Loop
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: GoTo TerminarMNI
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = IIf(IsNull(oRsDetalleProductos.Fields!RegistroSanitario), "", oRsDetalleProductos.Fields!RegistroSanitario)
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       oDoMovimientoDetalle.DocumentoNumero = IIf(IsNull(oRsDetalleProductos.Fields!NumeroDocumento), "", Trim(oRsDetalleProductos.Fields!NumeroDocumento)) 'Frank 07082015
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarMNI
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenDestino, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarMNI
       End If
       '*********  graba Devoluciones  de Pacientes***************
       If oDoMovimiento.idTipoConcepto = 21 And oDoMovimientoNotaIngreso.idCuentaAtencion > 0 Then
            With oDoFacturacionBienesDevol
                .cantidadAdevolver = oRsDetalleProductos.Fields!cantidad
                .FechaAutoriza = oDoMovimiento.FechaCreacion
                '.IdComprobantePago
                .idEstadoDevolucion = 2  'devuelto
                .idProducto = oRsDetalleProductos.Fields!idProducto
                .IdUsuarioAuditoria = oDoMovimiento.idUsuario
                .IdUsuarioAutoriza = oDoMovimiento.idUsuario
                .movNumero = oRsDetalleProductos.Fields!movNumeroS
                .MovTipo = "S"
                .MovNumeroE = oDoMovimiento.movNumero
                .MovTipoE = oDoMovimiento.MovTipo
            End With
            If Not oFacturacionBienesDevol.Insertar(oDoFacturacionBienesDevol) Then
                     bProcesoOK = False: GoTo TerminarMNI
            End If
       End If
       If oDoMovimiento.idTipoConcepto = 1 Or oDoMovimiento.idTipoConcepto = 2 Then
            '********* solo actualiza PRECIO ULTIMA COMPRA para CONCEPTOS=Compra/Encargo
            oDOCatalogoBienes.idProducto = oRsDetalleProductos.Fields!idProducto
            If Not oCatalogoBienes.SeleccionarPorId(oDOCatalogoBienes) Then
               bProcesoOK = False: GoTo TerminarMNI
            End If
            oDOCatalogoBienes.PrecioUltCompra = oRsDetalleProductos.Fields!Precio
            If Not oCatalogoBienes.modificar(oDOCatalogoBienes) Then
               bProcesoOK = False: GoTo TerminarMNI
            End If
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarMNI:
    If bProcesoOK Then
        ModificaDatosDeNotaIngreso = True
        oConexion.CommitTrans
    Else
        ModificaDatosDeNotaIngreso = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoNotaIngreso.MensajeError <> "", oMovimientoNotaIngreso.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesDevol.MensajeError <> "", oFacturacionBienesDevol.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoNotaIngreso = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oDOCatalogoBienes = Nothing
    Set oProveedores = Nothing
    Set oFacturacionBienesDevol = Nothing
    Set oDoFacturacionBienesDevol = Nothing
    Set rs = Nothing
End Function
Function AnulaNotaIngreso(oDoMovimiento As DoFarmMovimiento, oDOfarmMovimientoNotaIngreso As DOfarmMovimientoNotaIngreso, lnIdtipoFinanciamiento As Long, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoFacturacionBienesDevol As New DoFacturacionBienesDevol
    Dim oFacturacionBienesDevol As New FacturacionBienesDevol
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
    Dim oFacturacionBienesFinan As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinan As New DoFacturacionBienesFinanciam
    Dim oRsTmp As New Recordset, oRsDevoluciones As New Recordset
    Dim lcCorrelativo As String, lcSql As String
    Dim lnItem As Long, lnCantidadPagar As Long, lnIdOrden As Long
    Dim lnTotalPagar As Double
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oFacturacionBienesDevol.Conexion = oConexion
    Set oFacturacionBienesPagos.Conexion = oConexion
    Set oFacturacionBienesFinan.Conexion = oConexion
    
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarANI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "E", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  Elimina Devoluciones del Paciente
    If oDOfarmMovimientoNotaIngreso.idCuentaAtencion > 0 And oDoMovimiento.idTipoConcepto = 21 Then
        oRsDetalleProductos.MoveFirst
        Do While Not oRsDetalleProductos.EOF
            If ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento, oConexion) = True Then
                    oDoFacturacionBienesPagos.CantidadPagar = oRsDetalleProductos.Fields!CantidadS
                    oDoFacturacionBienesPagos.IdOrden = oRsDetalleProductos.Fields!IdOrdenS
                    oDoFacturacionBienesPagos.idProducto = oRsDetalleProductos.Fields!idProducto
                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
                    oDoFacturacionBienesPagos.PrecioVenta = oRsDetalleProductos.Fields!PrecioS
                    lnCantidadPagar = oDoFacturacionBienesPagos.CantidadPagar + oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oDoFacturacionBienesPagos.PrecioVenta, 2)
                    oDoFacturacionBienesPagos.CantidadPagar = lnCantidadPagar
                    oDoFacturacionBienesPagos.TotalPagar = lnTotalPagar
                    If Not oFacturacionBienesPagos.ModificarXidOrdenIdProducto(oDoFacturacionBienesPagos) Then
                       bProcesoOK = False: GoTo TerminarANI
                    End If
            Else
                    lnCantidadPagar = oRsDetalleProductos.Fields!CantidadS + oRsDetalleProductos.Fields!cantidad
                    lnTotalPagar = Round(lnCantidadPagar * oRsDetalleProductos.Fields!PrecioS, 2)
                    oDoFacturacionBienesFinan.movNumero = oRsDetalleProductos.Fields!movNumeroS
                    oDoFacturacionBienesFinan.MovTipo = "S"
                    oDoFacturacionBienesFinan.idProducto = oRsDetalleProductos.Fields!idProducto
                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = oRsDetalleProductos.Fields!IdTipoFinanciamientoS
                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = oRsDetalleProductos.Fields!IdFuenteFinanciamientoS
                    oDoFacturacionBienesFinan.FechaAutoriza = oRsDetalleProductos.Fields!FechaS
                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = oRsDetalleProductos.Fields!usuarioS
                    oDoFacturacionBienesFinan.PrecioFinanciado = oRsDetalleProductos.Fields!PrecioS
                    oDoFacturacionBienesFinan.TotalFinanciado = lnTotalPagar
                    oDoFacturacionBienesFinan.CantidadFinanciada = lnCantidadPagar
                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                    If Not oFacturacionBienesFinan.modificar(oDoFacturacionBienesFinan) Then
                       bProcesoOK = False: GoTo TerminarANI
                    End If
            End If
            oRsDetalleProductos.MoveNext
        Loop
        With oDoFacturacionBienesDevol
               .IdUsuarioAuditoria = oDoMovimiento.idUsuario
               .MovNumeroE = oDoMovimiento.movNumero
               .MovTipoE = oDoMovimiento.MovTipo
        End With
        If Not oFacturacionBienesDevol.EliminarPorMovNumeroE(oDoFacturacionBienesDevol) Then
             bProcesoOK = False: GoTo TerminarANI
        End If
    End If
    '*********  devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    rs.MoveFirst
    Do While Not rs.EOF
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenDestino, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
          bProcesoOK = False: GoTo TerminarANI
       End If
       rs.MoveNext
    Loop
TerminarANI:
    If bProcesoOK Then
        AnulaNotaIngreso = True
        oConexion.CommitTrans
    Else
        AnulaNotaIngreso = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesDevol.MensajeError <> "", oFacturacionBienesDevol.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oFacturacionBienesDevol = Nothing
    Set oDoFacturacionBienesDevol = Nothing
    Set rs = Nothing
End Function
'********************************   Nota de Ingreso ***************************************


'********************************   Nota de Salida  ***************************************
Function AgregaDatosDeNotaSalida(oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe items": GoTo TerminarNS
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductos.Fields!Codigo) + " " + Trim(oRsDetalleProductos.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarNS
       End If
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        AgregaDatosDeNotaSalida = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaSalida = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function
'SCCQ 20/10/2020 Cambio28 Inicio
'Se copió el código de la Función AgregaDatosDeNotaSalida y se agregó lo necesario para generar y actualizar correlativo disponible de número de documento.
Function AgregaDatosDeNotaSalida_NumDocAutomatico(lnIdTipoLocales As String, lnIdTipoSuministro As String, lnDocumentoIdtipo As Long, oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    'SCCQ 20/10/2020 Cambio28 Inicio
    Dim lcCorrelativoDocumN As String
    'SCCQ 20/10/2020 Cambio28 Fin
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    'SCCQ 20/10/2020 Cambio28 Inicio
    lcCorrelativoDocumN = oMovimiento.DevuelveYactualizaCorrelativoDisponibleFarmRelMod("S", lnIdTipoLocales, lnIdTipoSuministro, lnDocumentoIdtipo)
    oDoMovimiento.DocumentoNumero = lcCorrelativoDocumN
    'SCCQ 20/10/2020 Cambio28 Fin
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe items": GoTo TerminarNS
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductos.Fields!Codigo) + " " + Trim(oRsDetalleProductos.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarNS
       End If
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        AgregaDatosDeNotaSalida_NumDocAutomatico = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaSalida_NumDocAutomatico = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function
'SCCQ 20/10/2020 Cambio28 Fin
Function ModificaDatosDeNotaSalida(oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarMNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "M", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
'    oRsDetalleProductos.MoveFirst
'    Do While Not oRsDetalleProductos.EOF
'       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!IdProducto, oRsDetalleProductos.Fields!lote, oRsDetalleProductos.Fields!fechaVencimiento, oRsDetalleProductos.Fields!cantidad) Then
'             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
'             bProcesoOK = False: GoTo TerminarMNS
'       End If
'       oRsDetalleProductos.MoveNext
'    Loop
    '*********  Elimina tabla farmMovimientosDetalle y devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    rs.MoveFirst
    Do While Not rs.EOF
       oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
       oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
       oDoMovimientoDetalle.idProducto = rs.Fields!idProducto
       oDoMovimientoDetalle.Lote = rs.Fields!Lote
       oDoMovimientoDetalle.FechaVencimiento = rs.Fields!FechaVencimiento
       oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = rs.Fields!idTipoSalidaBienInsumo
       If Not FarmMovimientoDetalleEliminar(oConexion, oDoMovimientoDetalle) Then
          bProcesoOK = False: GoTo TerminarMNS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
          bProcesoOK = False: GoTo TerminarMNS
       End If
       rs.MoveNext
    Loop
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: GoTo TerminarMNS
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductos.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductos.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarMNS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos.Fields!idProducto, oRsDetalleProductos.Fields!Lote, oRsDetalleProductos.Fields!FechaVencimiento, oRsDetalleProductos.Fields!idTipoSalidaBienInsumo, oRsDetalleProductos.Fields!cantidad, oRsDetalleProductos.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarMNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarMNS:
    If bProcesoOK Then
        ModificaDatosDeNotaSalida = True
        oConexion.CommitTrans
    Else
        ModificaDatosDeNotaSalida = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set rs = Nothing
End Function
Function AnulaNotaSalida(oDoMovimiento As DoFarmMovimiento, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                        lnIdReceta As Long, lnIdOrdenPago As Long, Optional mo_DoFarmMovimientoVentas As DoFarmMovimientoVentas) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
    Dim oDoFacturacionBienesFinanciam As New DoFacturacionBienesFinanciam
    Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
    Dim oFactOrdenesBienes As New FactOrdenesBienes
    Dim oCommand As New ADODB.Command
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim oParameter As ADODB.Parameter
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    

    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarANS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "E", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  anula tabla FactOrdenesBienes
    oDoFactOrdenesBienes.movNumero = oDoMovimiento.movNumero
    oDoFactOrdenesBienes.MovTipo = oDoMovimiento.MovTipo
    If Not oFactOrdenesBienes.SeleccionarPorMovNumero(oDoFactOrdenesBienes) Then
            bProcesoOK = False: GoTo TerminarANS
    End If
    If oDoFactOrdenesBienes.IdEstadoFacturacion = 1 Then
        oDoFactOrdenesBienes.IdEstadoFacturacion = 9   'anulado
        If Not oFactOrdenesBienes.modificar(oDoFactOrdenesBienes) Then
                bProcesoOK = False: GoTo TerminarANS
        End If
    End If
    '
    If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacion(oDoMovimiento.movNumero, oDoMovimiento.MovTipo, 9) Then
       bProcesoOK = False: GoTo TerminarANS
    End If
    '*********  devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
           If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
              bProcesoOK = False: GoTo TerminarANS
           End If
           rs.MoveNext
        Loop
    End If
    'Receta automatica debb-03/04/2012
    If lnIdReceta > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalle lnIdReceta, mo_DoFarmMovimientoVentas.idCuentaAtencion, 0, oConexion
    End If
    '
    If lnIdOrdenPago > 0 Then
        If mo_ReglasAdmision.EPSeliminaOrdenPago(oConexion, lnIdOrdenPago, mo_DoFarmMovimientoVentas.IdUsuarioAuditoria) = False Then
        End If
    End If
    '
TerminarANS:
    If bProcesoOK Then
        AnulaNotaSalida = True
        oConexion.CommitTrans
    Else
        AnulaNotaSalida = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFacturacionBienesFinanciam.MensajeError <> "", oFacturacionBienesFinanciam.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oCommand = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set rs = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oDoFacturacionBienesFinanciam = Nothing
    Set oDoFactOrdenesBienes = Nothing
    Set oDoFacturacionBienesFinanciam = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oCommand = Nothing
    Set mo_ReglasImagenes = Nothing
End Function

Function FarmMovimientoSeleccionarPorId(ByVal oDoMovimiento As DoFarmMovimiento) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    oConexion.Open sighentidades.CadenaConexion
    Set oMovimiento.Conexion = oConexion
    If oMovimiento.SeleccionarPorId(oDoMovimiento) Then
        FarmMovimientoSeleccionarPorId = True
    Else
        FarmMovimientoSeleccionarPorId = False
        ms_MensajeError = oMovimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
End Function


Function FarmMovimientoDetalleEliminar(oConexion As ADODB.Connection, ByVal oDoMovimientoDetalle As DoFarmMovimientoDetalle) As Boolean
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Set oMovimientoDetalle.Conexion = oConexion
    If oMovimientoDetalle.Eliminar(oDoMovimientoDetalle) Then
        FarmMovimientoDetalleEliminar = True
    Else
        FarmMovimientoDetalleEliminar = False
        ms_MensajeError = oMovimientoDetalle.MensajeError
    End If
    Set oMovimientoDetalle.Conexion = Nothing
End Function

'********************************   Nota de Salida  ***************************************



'********************************   farmTipoDocumentos  ***************************************
'Modificado por Yamill Palomino

Function FarmTipoDocumentosDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmTipoDocumentosDevuelveTodos = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoDocumentosDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmTipoDocumentosDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing

    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'********************************   farmTipoDocumentos  ***************************************

'********************************   farmTipoCompra  ***************************************

'Modificado por Yamill Palomino
Function FarmTipoCompraDevuelveSegunFiltro(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Dim oParameter As ADODB.Parameter

    Set FarmTipoCompraDevuelveSegunFiltro = Nothing

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoCompraDevuelveSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set FarmTipoCompraDevuelveSegunFiltro = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'********************************   farmTipoCompra  ***************************************

'********************************   farmTipoProceso  ***************************************
'Modificado por Yamill Palomino
Function FarmTipoProcesoDevuelveSegunFiltro(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmTipoProcesoDevuelveSegunFiltro = Nothing
     
       With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoProcesoDevuelveSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   
    Set FarmTipoProcesoDevuelveSegunFiltro = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'********************************   farmTipoProceso  ***************************************


'********************************   farmTipoConceptos  ***************************************
'Modificado por Yamill Palomino
Function FarmTipoConceptosDevuelveParaRegistroDeNiNs(lcTipoAlmacen As String, lcTipoMov As String, lcTipoSuministro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmTipoConceptosDevuelveParaRegistroDeNiNs = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoConceptosDevuelveParaRegistroDeNiNs"
        Set oParameter = .CreateParameter("@TipoAlmacen", adVarChar, adParamInput, 50, lcTipoAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoMov", adVarChar, adParamInput, 50, lcTipoMov): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoSuministro", adVarChar, adParamInput, 50, lcTipoSuministro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set FarmTipoConceptosDevuelveParaRegistroDeNiNs = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function FarmTipoConceptosDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmTipoConceptosDevuelveTodos = Nothing

  
      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoConceptosDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmTipoConceptosDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
   
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'********************************   farmTipoConceptos  ***************************************


'********************************   ALMACENES  ***************************************

Function FarmAlmacenSeleccionarTodosMenosExternos() As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New farmAlmacen
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set FarmAlmacenSeleccionarTodosMenosExternos = oTabla.SeleccionarTodosMenosExternos
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
Function FarmAlmacenSeleccionarSegunFiltro(lcFiltro As String, Optional oConexion1 As Connection) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New farmAlmacen
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    Set FarmAlmacenSeleccionarSegunFiltro = oTabla.SeleccionarSegunFiltro(lcFiltro)
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
'********************************   AMACENES ***************************************

'********************************   FarmMovimientosDetalle  ***************************************
Function FarmMovimientosDetalleDevuelveTodosItems(oConexion As ADODB.Connection, lcMovTipo As String, lcMovNumero As String) As ADODB.Recordset
    Dim oTabla As New FarmMovimientoDetalle
    Set oTabla.Conexion = oConexion
    Set FarmMovimientosDetalleDevuelveTodosItems = oTabla.DevuelveTodosItems(lcMovNumero, lcMovTipo)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function
'********************************   farmMovimientosDetalle ***************************************


'********************************   INVENTARIO  ***************************************
Function FarmInventarioCabeceraDevuelveProductosPorId(lnIdInventario As Long) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New FarmInventarioCabecera
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set FarmInventarioCabeceraDevuelveProductosPorId = oTabla.DevuelveProductosPorId(lnIdInventario)
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function FarmInventarioDetalleDevuelveProductosLotesPorId(lnIdInventario As Long) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New farmInventarioDetalle
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set FarmInventarioDetalleDevuelveProductosLotesPorId = oTabla.DevuelveProductosLotesPorId(lnIdInventario)
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function


Function FarmInventarioSeleccionarPorId(ByVal oDoInventario As DOfarmInventario) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New farmInventario
    oConexion.Open sighentidades.CadenaConexion
    Set oInventario.Conexion = oConexion
    If oInventario.SeleccionarPorId(oDoInventario) Then
        FarmInventarioSeleccionarPorId = True
    Else
        FarmInventarioSeleccionarPorId = False
        ms_MensajeError = oInventario.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInventario = Nothing
End Function

Function FarmInventarioDevuelveLista() As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New farmInventario
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set FarmInventarioDevuelveLista = oTabla.DevuelveListaInventarios
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    

End Function


'********************************   INVENTARIO  ***************************************




'********************************   Correlativo  ***************************************


'********************************   Actualiza Saldos por Producto  ***************************************

Function FarmDevuelveSaldosSegunAlmacen(lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmDevuelveSaldosSegunAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmDevuelveSaldosSegunAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveSaldosSegunAlmacen = oRecordset
  
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmDevuelveMovimientos(lnIdAlmacen As Long, lcTipoMov As String, ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmDevuelveMovimientos = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmDevuelveMovimientos"
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcTipoMov): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveMovimientos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function FarmDevuelveCabeceraDeVentasOpreventa(lcTipoVenta As String, lnIdAlmacen As Long, lcTipoMov As String, ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmDevuelveCabeceraDeVentasOpreventa = Nothing
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveCabeceraDeVentasOpreventa"
        Set oParameter = .CreateParameter("@tipoVenta", adVarChar, adParamInput, 1, lcTipoVenta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveCabeceraDeVentasOpreventa = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmDevuelveSaldosSinLotesSegunAlmacen(lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmDevuelveSaldosSinLotesSegunAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosSinLotesSegunAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveSaldosSinLotesSegunAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function farmDevuelveSaldosSegunAlmacenProducto(lnIdAlmacen As Long, lnIdProducto As Long, _
                                                Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmDevuelveSaldosSegunAlmacenProducto = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosSegunAlmacenProducto"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set farmDevuelveSaldosSegunAlmacenProducto = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino
Function TiposCargoSeleccionarTodos() As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set TiposCargoSeleccionarTodos = Nothing

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposCargoSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set TiposCargoSeleccionarTodos = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
End Function


'Modificado por Yamill Palomino
Function EmpleadosDevuelvePrescriptores() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set EmpleadosDevuelvePrescriptores = Nothing
     
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDevuelvePrescriptores"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set EmpleadosDevuelvePrescriptores = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function EmpleadosDevuelveCoordinadores() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set EmpleadosDevuelveCoordinadores = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDevuelveCoordinadores"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set EmpleadosDevuelveCoordinadores = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

' Modificado por Yamill Palomino
Function EmpleadosDeImagen(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set EmpleadosDeImagen = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDeImagen"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set EmpleadosDeImagen = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function




'Modificado por Yamill Palomino
Function FarmComponenteDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmComponenteDevuelveTodos = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmComponenteDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmComponenteDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    
    
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function FarmTipoRecetaDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmTipoRecetaDevuelveTodos = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmTipoRecetaDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmTipoRecetaDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    
    
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function


'Modificado por Yamill Palomino
Function FarmComponenteSubDevuelveTodosSegunComponente(lnIdComponente As Long) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set FarmComponenteSubDevuelveTodosSegunComponente = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmComponenteSubDevuelveTodosSegunComponente"
        Set oParameter = .CreateParameter("@idComponente", adInteger, adParamInput, 0, lnIdComponente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmComponenteSubDevuelveTodosSegunComponente = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function


'Modificado por Yamill Palomino
Function farmTipoSalidaBienInsumoDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmTipoSalidaBienInsumoDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmTipoSalidaBienInsumoDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
   
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

Function farmRegeneraSaldos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmRegeneraSaldos = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 300
        .CommandText = "FarmRegeneraSaldos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmRegeneraSaldos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino
Function TipoFinanciamientosDevuelveSoloFarmacia(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set TipoFinanciamientosDevuelveSoloFarmacia = Nothing
     
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TipoFinanciamientosDevuelveSoloFarmacia"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set TipoFinanciamientosDevuelveSoloFarmacia = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
    'Resume
End Function

'Modificado por Yamill Palomino
Function TipoFinanciamientosDevuelveSoloPlanes(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set TipoFinanciamientosDevuelveSoloPlanes = Nothing

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TipoFinanciamientosDevuelveSoloPlanes"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set TipoFinanciamientosDevuelveSoloPlanes = oRecordset
    oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'Modificado por Yamill Palomino
Function TipoFinanciamientosParaCatalogoBienes() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TipoFinanciamientosParaCatalogoBienes"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set TipoFinanciamientosParaCatalogoBienes = oRecordset
    oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

Function DevuelveEstadoDocumento1(lcSql As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    DevuelveEstadoDocumento1 = ""
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = lcSql
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    DevuelveEstadoDocumento1 = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDeFacturacion(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDeFacturacion = ""

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDeFacturacion   "
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    DevuelveEstadoActualDeFacturacion = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDeEstadoCuenta(lcFiltro As String, oConexion As Connection) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    DevuelveEstadoActualDeEstadoCuenta = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDeEstadoCuenta"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
     
    DevuelveEstadoActualDeEstadoCuenta = oRecordset.Fields!lcEstado
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDeLaboratorio(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDeLaboratorio = ""
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDeLaboratorio"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    DevuelveEstadoActualDeLaboratorio = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDeImagen(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDeImagen = ""

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDeImagen"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    DevuelveEstadoActualDeImagen = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDePreventa(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDePreventa = ""

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDePreventa"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    DevuelveEstadoActualDePreventa = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino
Function DevuelveEstadoActualDelMovimiento(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDelMovimiento = ""
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDelMovimiento"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    DevuelveEstadoActualDelMovimiento = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function


'Modificado por Yamill Palomino
Function DevuelveEstadoActualDelInventario(lcFiltro As String) As String
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim lcSql As String
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    DevuelveEstadoActualDelInventario = ""
     
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveEstadoActualDelInventario"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    DevuelveEstadoActualDelInventario = oRecordset.Fields!lcEstado
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

'********************************   Dependencia Externa  ***************************************

'Modificado por YAmill Palomino
Function FarmAlmacenesDevuelveListaDependenciasExternas() As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmAlmacenesDevuelveListaDependenciasExternas = Nothing

    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmAlmacenesDevuelveListaDependenciasExternas"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FarmAlmacenesDevuelveListaDependenciasExternas = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function
Function FarmAlmacenAgregar(oDOFarmAlmacen As DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFarmAlmacen As New farmAlmacen
    FarmAlmacenAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmAlmacen.Conexion = oConexion
    If oFarmAlmacen.Insertar(oDOFarmAlmacen) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFarmAlmacen.IdUsuarioAuditoria, "A", oDOFarmAlmacen.IdAlmacen, _
                "FarmAlmacen", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Farm:" & oDOFarmAlmacen.Descripcion)
        FarmAlmacenAgregar = True
    Else
        ms_MensajeError = oFarmAlmacen.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFarmAlmacen = Nothing
End Function

Function FarmAlmacenModificar(oDOFarmAlmacen As DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFarmAlmacen As New farmAlmacen
    FarmAlmacenModificar = False
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmAlmacen.Conexion = oConexion
    If oFarmAlmacen.modificar(oDOFarmAlmacen) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFarmAlmacen.IdUsuarioAuditoria, _
                    IIf(oDOFarmAlmacen.IdEstado = 0, "E", "M"), oDOFarmAlmacen.IdAlmacen, "FarmAlmacen", oConexion, _
                        mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Farm:" & oDOFarmAlmacen.Descripcion)
        FarmAlmacenModificar = True
    Else
        ms_MensajeError = oFarmAlmacen.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFarmAlmacen = Nothing
End Function

Function FarmAlmacenEliminar(oDOFarmAlmacen As DoFarmAlmacen, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oFarmAlmacen As New farmAlmacen
    FarmAlmacenEliminar = False
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmAlmacen.Conexion = oConexion
    If oFarmAlmacen.Eliminar(oDOFarmAlmacen) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFarmAlmacen.IdUsuarioAuditoria, _
                "E", oDOFarmAlmacen.IdAlmacen, "FarmAlmacen", oConexion, _
                    mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "Farm:" & oDOFarmAlmacen.Descripcion)
        FarmAlmacenEliminar = True
    Else
        ms_MensajeError = oFarmAlmacen.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFarmAlmacen = Nothing
End Function
Function FarmAlmacenSeleccionarPorId(IdAlmacen As Long) As DoFarmAlmacen
    Dim oConexion As New ADODB.Connection
    Dim oFarmAlmacen As New farmAlmacen
    Dim oDOFarmAlmacen As New DoFarmAlmacen
    Set FarmAlmacenSeleccionarPorId = Nothing
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFarmAlmacen.Conexion = oConexion
    oDOFarmAlmacen.IdAlmacen = IdAlmacen
    If oFarmAlmacen.SeleccionarPorId(oDOFarmAlmacen) Then
        Set FarmAlmacenSeleccionarPorId = oDOFarmAlmacen
    End If
    oConexion.Close
    ms_MensajeError = oFarmAlmacen.MensajeError
    Set oConexion = Nothing
    Set oFarmAlmacen = Nothing
End Function
'********************************   Dependencia Externa  ***************************************


Function FarmMovimientoVentasSeleccionarPorIdPreventa(lnIdPreventa As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoVentasSeleccionarPorIdPreventa = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasSeleccionarPorIdPreventa"
        Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, lnIdPreventa): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasSeleccionarPorIdPreventa = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'*************Precio que se usará en: Ni, Ns, Ventas, Intervenciones, Inventario
'Modificado por YAmill Palomino
Function DevuelvePrecioSegunTipoConcepto(lnIdProducto As Long, lnIdTipoPrecioParaNiNs As sghTipoPrecioFarmacia) As Double
    Dim lcSql As String
    Dim lnPrecio As Double
    Dim oRsTmp As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    lnPrecio = 0
    
      
    If lnIdTipoPrecioParaNiNs = 5 Then
         DevuelvePrecioSegunTipoConcepto = lnPrecio
    Else
         If lnIdTipoPrecioParaNiNs > 0 Then
                      
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "DevuelvePrecioSegunTipoConcepto"
                 Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdTipoPrecioParaNiNs", adInteger, adParamInput, 0, lnIdTipoPrecioParaNiNs): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@sghPrecioCompra", adInteger, adParamInput, 0, sghPrecioCompra): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@sghPrecioDistribucion", adInteger, adParamInput, 0, sghPrecioDistribucion): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@sghPrecioVentaContado", adInteger, adParamInput, 0, sghPrecioVentaContado): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@sghPrecioDonacion", adInteger, adParamInput, 0, sghPrecioDonacion): .Parameters.Append oParameter
                 Set oRsTmp = .Execute
                 Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
   
            If oRsTmp.RecordCount > 0 Then
                'mgaray20141016
                If oRsTmp.RecordCount > 1 Then
                    MsgBox "Producto Tiene dos precios activos, Antes de Proceder a registrar ir al catalogo de Bienes e Insumos y depurar los precios", vbInformation, "Advertencia"
                    lnPrecio = IIf(IsNull(oRsTmp.Fields!lnPrecio), 0, oRsTmp.Fields!lnPrecio)
                Else
                    lnPrecio = IIf(IsNull(oRsTmp.Fields!lnPrecio), 0, oRsTmp.Fields!lnPrecio)
                End If
            End If
            oRsTmp.Close
         End If
         DevuelvePrecioSegunTipoConcepto = lnPrecio
    End If
    Set oRsTmp = Nothing
End Function

'Modificado por Yamill Palomino
Function DevuelvePrecioSegunTipoFinanciamiento(lnIdProducto As Long, lnIdtipoFinanciamiento As Long) As Double
    Dim lcSql As String
    Dim lnPrecio As Double
    Dim oRsTmp As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelvePrecioSegunTipoFinanciamiento"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
    lnPrecio = 0
    If oRsTmp.RecordCount > 0 Then
       lnPrecio = IIf(IsNull(oRsTmp.Fields!lnPrecio), 0, oRsTmp.Fields!lnPrecio)
    End If
    oRsTmp.Close
    DevuelvePrecioSegunTipoFinanciamiento = lnPrecio
    Set oRsTmp = Nothing
End Function


Function farmMovimientoSeleccionarPorTipoYnumeroDocumento(lcDocumentoNumero As String, lnDocumentoIdtipo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoSeleccionarPorTipoYnumeroDocumento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoSeleccionarPorTipoYnumeroDocumento"
        Set oParameter = .CreateParameter("@documentoNumero", adVarChar, adParamInput, 20, lcDocumentoNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@documentoIdTipo", adInteger, adParamInput, 0, lnDocumentoIdtipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoSeleccionarPorTipoYnumeroDocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function FarmMovimientoVentasDetalleSeleccionarPorCuenta(lnCuenta As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmMovimientoVentasDetalleSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmMovimientoVentasDetalleSeleccionarPorCuenta = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoVentasDetalleSeleccionarPorFechas(ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoVentasDetalleSeleccionarPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDate, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDate, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasDetalleSeleccionarPorFechas = oRecordset
  
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoNotaIngresoSeleccionarPorCuenta(lnCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoNotaIngresoSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoNotaIngresoSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoNotaIngresoSeleccionarPorCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FactOrdenesBienesSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenesBienesSeleccionarPorIdComprobantePago = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenesBienesSeleccionarPorIdComprobantePago = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesPagosSeleccionarPorCuenta(lnCuenta As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesPagosSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        
    Else
        Set oConexion = oConexion1
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FacturacionBienesPagosSeleccionarPorCuenta = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesPagosSeleccionarPorId(lnIdOrden As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesPagosSeleccionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorId"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagosSeleccionarPorId = oRecordset
  
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function AtencionesSelecionarPorCuenta(lnCuenta As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSelecionarPorCuenta = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "atencionesSelecionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set AtencionesSelecionarPorCuenta = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
   If oConexion1 Is Nothing Then
      oConexion.Close
   End If
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function FarmMovimientoNotaIngresoFiltraFuentesFinanciamiento(ldFechaInicio As Date, ldFechaFin As Date, lnidtipoConcepto As Long, lnIdAlmacen As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoNotaIngresoFiltraFuentesFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoNotaIngresoFiltraFuentesFinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoNotaIngresoFiltraFuentesFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoProgramasSeleccionarPorFechas(lnIdAlmacen As Long, ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmMovimientoProgramasSeleccionarPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoProgramasSeleccionarPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoProgramasSeleccionarPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino-ica
Function DevuelveListaDeUsuariosDelSistema() As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set DevuelveListaDeUsuariosDelSistema = Nothing
    


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveListaDeUsuariosDelSistema"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set DevuelveListaDeUsuariosDelSistema = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function

Function FarmMovimientoVentasFiltraTipoServicioHosp(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoVentasFiltraTipoServicioHosp = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasFiltraTipoServicioHosp"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasFiltraTipoServicioHosp = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function




Function FacturacionBienesPagosSeleccionarPorMovnumero(lcMovTipo As String, lcMovNumero As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FacturacionBienesPagosSeleccionarPorMovnumero = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorMovnumero"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagosSeleccionarPorMovnumero = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmDevuelveSaldosConSinLotesPorAlmacen(lnIdAlmacen As Long, lnOrden As Long, lcReporte As String, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmDevuelveSaldosConSinLotesPorAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmDevuelveSaldosConSinLotesPorAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Reporte", adVarChar, adParamInput, 50, lcReporte): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 1, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveSaldosConSinLotesPorAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino
Function BuscaProductosPorVencer(lcFechaMaxima As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL1 As String: Dim sSQL2 As String
Dim oConexion As New ADODB.Connection
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set BuscaProductosPorVencer = Nothing

     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "BuscaProductosPorVencer"
          Set oParameter = .CreateParameter("@FechaMaxima", adDBTimeStamp, adParamInput, 0, lcFechaMaxima): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
     
    Set BuscaProductosPorVencer = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto(lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
    Dim oFacturacionBienesPagos As New FacturacionBienesPagos
    Set FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto = oFacturacionBienesPagos.SeleccionarPorMovNumeroIdProducto(lcMovNumero, lcMovTipo, lnIdProducto)
    Set oFacturacionBienesPagos = Nothing
End Function

Function oFacturacionBienesFinanciamientoSeleccionarPorIdProducto(lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
    Dim oFacturacionBienesFinanc As New FacturacionBienesFinanciam
    Set oFacturacionBienesFinanciamientoSeleccionarPorIdProducto = oFacturacionBienesFinanc.SeleccionarPorIdProducto(lcMovNumero, lcMovTipo, lnIdProducto)
    Set oFacturacionBienesFinanc = Nothing
End Function

Function farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote(lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date) As ADODB.Recordset
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Set farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote = oMovimientoDetalle.farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote(lnIdAlmacen, lnIdProducto, lcLote, ldFechaVencimiento)
    Set oMovimientoDetalle = Nothing
End Function

Function FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String) As Recordset
Dim oFarmMovimientoDetalle As New FarmMovimientoDetalle
Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFarmMovimientoDetalle.Conexion = oConexion
    Set FarmDevuelveSaldosConLotesSegunAlmacen = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen, lnOrden, lcFiltro, oConexion)
    Set oFarmMovimientoDetalle = Nothing

End Function

Function AuditoriaPorFechasUsuario(lIdUsuarioSistema As Long, lcFechaIni As Date, lcFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AuditoriaPorFechasUsuario = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AuditoriaPorFechasUsuario"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, lcFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, lcFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, lIdUsuarioSistema): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AuditoriaPorFechasUsuario = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoFiltraPorMovNumeroTipo(lcMovNumero As String, lcMovTipo As String, Optional oConexion1 As Connection) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmMovimientoFiltraPorMovNumeroTipo = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoFiltraPorMovNumeroTipo"
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oRecordset = .Execute

   End With
   Set FarmMovimientoFiltraPorMovNumeroTipo = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmPreVentaFiltraPorIdPreVenta(lidPreventa As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmPreVentaFiltraPorIdPreVenta = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FarmPreVentaFiltraPorIdPreVenta"
        Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, lidPreventa): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FarmPreVentaFiltraPorIdPreVenta = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oConexion = Nothing
        Set oRecordset = Nothing
   End If
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FacturacionCuentasAtencionSeleccionarPorIdPaciente(lidPaciente As Long, oConexion As Connection, _
                                                     lbMuestraVentanaDeVariasCuentasAbiertas As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionCuentasAtencionSeleccionarPorIdPaciente = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lidPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If lbMuestraVentanaDeVariasCuentasAbiertas = True And sighentidades.Parametro556 = "S" Then
           oRecordset.Filter = "idEstado=" & sghEstadoCuenta.sghAbierto
           If oRecordset.RecordCount > 1 Then
                Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
                Dim DOPaciente As New DOPaciente
                Dim rs As New Recordset
                DOPaciente.idPaciente = lidPaciente
                Set rs = mo_AdminAdmision.AtencionesBusquedaDeAtenciones(DOPaciente)
                rs.Filter = "idEstado=" & sghEstadoCuenta.sghAbierto
                Dim oBusqueda As New PacientesCuentas
                Set oBusqueda.Atenciones = rs
                oBusqueda.Show 1
                oRecordset.Filter = "idCuentaAtencion=" & oBusqueda.idCuentaAtencion
                Set oBusqueda = Nothing
                Set mo_AdminAdmision = Nothing
                Set DOPaciente = Nothing
                Set rs = Nothing
           Else
              oRecordset.Filter = ""
           End If
        End If
   End With
   Set FacturacionCuentasAtencionSeleccionarPorIdPaciente = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function




Function DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(lidCuenta As Long, ldFecha As Date, lcHora As String, _
                  Optional oConexion1 As Connection) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnIdServicioDelPaciente As Long
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   lnIdServicioDelPaciente = 0
   If oRecordset.RecordCount > 0 Then
      '****hospitalizado
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
         If IsNull(oRecordset.Fields!fechaDesocupacion) Then
            If ldFecha >= oRecordset.Fields!fechaOcupacion Then
               lnIdServicioDelPaciente = oRecordset.Fields!IdServicio
               Exit Do
            End If
         Else
            If ldFecha >= oRecordset.Fields!fechaOcupacion And ldFecha <= oRecordset.Fields!fechaDesocupacion Then
               If ldFecha = oRecordset.Fields!fechaOcupacion And lcHora >= oRecordset.Fields!HoraOcupacion And lcHora <= oRecordset.Fields!HoraDesocupacion Then
                  lnIdServicioDelPaciente = oRecordset.Fields!IdServicio
                  Exit Do
               End If
            End If
         End If
         oRecordset.MoveNext
      Loop
   End If
   Set oCommand = Nothing
   If oRecordset.RecordCount = 0 Or lnIdServicioDelPaciente = 0 Then
        '****consultorio externo
        If oConexion1 Is Nothing Then
           Set oRecordset = AtencionesSelecionarPorCuenta(lidCuenta, oConexion)
        Else
           Set oRecordset = AtencionesSelecionarPorCuenta(lidCuenta, oConexion1)
        End If
        If oRecordset.RecordCount > 0 Then
           lnIdServicioDelPaciente = oRecordset.Fields!idServicioIngreso
        End If
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oConexion = Nothing
   End If
   DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora = lnIdServicioDelPaciente
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EmpleadosDevuelveIdCargoSegunPuntoCarga(lnIdPuntoCarga As Long) As Long
    Select Case lnIdPuntoCarga
    Case 21   'rayos X
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 7
    Case 22   'tomografia
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 6
    Case 23    'ecogr. obst
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 4
    Case 20    'ecogr. general
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 8
    Case 2, 31, 33, 34, 35, 36, 37 'Patología Clínica
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 2
    Case 3, 32 'Anatomia Patológica
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 3
    Case 11, 38 'Banco de Sangre
       EmpleadosDevuelveIdCargoSegunPuntoCarga = 11
    End Select
    
    
    
    
    
End Function

Function farmMovimientoVentasSeleccionarPorIdCuentaAtencion(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasSeleccionarPorIdCuentaAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasSeleccionarPorIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasSeleccionarPorIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmMovimientoVentasSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasSeleccionarPorIdComprobantePago = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@lnIdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasSeleccionarPorIdComprobantePago = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesDevolucionesSeleccionarPorMovNumeroE(lcMovNumeroE As String, lcMovTipoE As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesDevolucionesSeleccionarPorMovNumeroE = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesDevolucionesSeleccionarPorMovNumeroE"
        Set oParameter = .CreateParameter("@MovNumeroE", adVarChar, adParamInput, 9, lcMovNumeroE)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipoE", adVarChar, adParamInput, 1, lcMovTipoE)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesDevolucionesSeleccionarPorMovNumeroE = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmMovimientoSeleccionarPorMovNumero(lcMovNumero As String, lcMovTipo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoSeleccionarPorMovNumero = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoSeleccionarPorId"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoSeleccionarPorMovNumero = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ServiciosDondeEstuvoElPaciente(idPaciente As Long, idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar los servicios donde estuvo un paciente por atencion
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ServiciosDondeEstuvoElPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ServiciosDondeEstuvoElPaciente = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function AtencionesQueTuvoElPaciente(idPaciente As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar las atenciones que tuvo un paciente
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AtencionesQueTuvoElPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set AtencionesQueTuvoElPaciente = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function DiagnosticosSeleccionarPorIdAtencion(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar el Diagnóstico de Salida por IdAtencion
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "DiagnosticosSeleccionarPorIdAtencion"
    Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set DiagnosticosSeleccionarPorIdAtencion = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaFarmaciaPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Farmacia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaFarmaciaPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaImagenologiaPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Imageneologia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaImagenologiaPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaImagenologiaPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaInsumosPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaInsumosPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaLaboratorioPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Laboratorio de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaLaboratorioPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaLaboratorioPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaProcedimientosPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaProcedimientosPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaProcedimientosPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnNroCuenta As Long) As Double
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnTotal As Double
    '
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorNroCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   '
   lnTotal = 0
   oRecordset.Filter = "idEstadoMovimiento=1"
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
          lnTotal = lnTotal + oRecordset.Fields!TotalFinanciado
          oRecordset.MoveNext
      Loop
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta = lnTotal
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function farmMovimientoVentasExoneracionesEnFarmacia(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "farmMovimientoVentasExoneracionesEnFarmacia"
    Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set farmMovimientoVentasExoneracionesEnFarmacia = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function farmMovimientoVentasSeleccionarPorNroDocumentoTipoFinanciamiento(lcNroDocumento As String, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "farmMovimientoVentasSeleccionarPorNroDocumentoTipoFinanciamiento"
    Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, lcNroDocumento): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set farmMovimientoVentasSeleccionarPorNroDocumentoTipoFinanciamiento = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function farmMovimientoVentasDetalleSeleccionarPorId(lcMovNumero As String, lcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set farmMovimientoVentasDetalleSeleccionarPorId = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasDetalleSeleccionarPorId"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
  
   Set oCommand = Nothing
   Set farmMovimientoVentasDetalleSeleccionarPorId = oRecordset
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoVentasFiltraPorFechas(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoVentasFiltraPorFechas"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set FarmMovimientoVentasFiltraPorFechas = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function FarmMovimientoFiltraPorFechasYtipoConcepto(ldFechaIni As Date, ldFechaFin As Date, lnidtipoConcepto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoFiltraPorFechasYtipoConcepto"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set FarmMovimientoFiltraPorFechasYtipoConcepto = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'******daniel-27-12-2009-inicio
Function FarmMovimientoVentasFiltraPorFechasAlmacen(ldFechaIni As Date, ldFechaFin As Date, lnIdAlmacenOrigen As Long, lcMovNumero As String, lcMovTipo As String) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoVentasFiltraPorFechasAlmacen"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacenOrigen): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set FarmMovimientoVentasFiltraPorFechasAlmacen = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmacen(ldFechaInicio As Date, ldFechaFin As Date, lnidtipoConcepto As Long, lnIdAlmacen As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmacen"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'******daniel-27-12-2009-fin




Function FacturacionBienesFinanciamientosFiltraCuentaMovimientos(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, lnIdProductoConSinNull As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesFinanciamientosFiltraCuentaMovimientos = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosFiltraCuentaMovimientos"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ConSinNull", adInteger, adParamInput, 0, lnIdProductoConSinNull): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinanciamientosFiltraCuentaMovimientos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmMovimientoDetalleSeleccionarDevolucionesPorTipoConceptoAlmacen(ldFechaInicio As Date, ldFechaFin As Date, lnidtipoConcepto As Long, lnIdAlmacenDestino As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoDetalleSeleccionarDevolucionesPorTipoConceptoAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleSeleccionarDevolucionesPorTipoConceptoAlmacen"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoDetalleSeleccionarDevolucionesPorTipoConceptoAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmOrigen(ldFechaInicio As Date, ldFechaFin As Date, lnidtipoConcepto As Long, lnIdAlmacenOrigen As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmOrigen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmOrigen"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacenOrigen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoDetalleSeleccionaPorFechasTipoConceptoAlmOrigen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Sub FarmParteDiarioConRecalculo(ldFechaInicioSistema As Date, ldFechaInicio As Date, ldFechaFin As Date, lc_AlmacenesPICI As String, lBmb_ConsiderarRecalculo As Boolean, lBmb_ConsideraOSH As Boolean, lBmb_ConsiderarSinMovimientos As Boolean, ld_FechaVencimiento As Date)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
   oConexion.Open sighentidades.CadenaConexion
   oConexion.CursorLocation = adUseClient
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FarmParteDiarioConRecalculo"
        Set oParameter = .CreateParameter("@fechaInicioSistema", adDBTimeStamp, adParamInput, 0, ldFechaInicioSistema): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mda_FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mda_FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lc_AlmacenesParaICI", adVarChar, adParamInput, 20, lc_AlmacenesPICI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mb_ConsiderarRecalculo", adBoolean, adParamInput, 0, lBmb_ConsiderarRecalculo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mb_ConsideraOSH", adBoolean, adParamInput, 0, lBmb_ConsideraOSH): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@mb_ConsiderarSinMovimientos", adBoolean, adParamInput, 0, lBmb_ConsiderarSinMovimientos): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaVencimiento", adDBTimeStamp, adParamInput, 0, ld_FechaVencimiento): .Parameters.Append oParameter
       .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
  
   Set oCommand = Nothing
   ms_MensajeError = ""
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub


Function FacturacionBienesPagoSeleccionarPorMovNumeroProducto(mo_conexion As Connection, lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionBienesPagoSeleccionarPorMovNumeroProducto = Nothing
    ms_MensajeError = ""
    oRecordset.CursorType = adOpenForwardOnly
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagoSeleccionarPorMovNumeroProducto = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesFinancSeleccionarPorProducto(mo_conexion As Connection, lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oRecordset.CursorType = adOpenForwardOnly
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinancSeleccionarPorProducto = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmDevuelveSaldosConLotesSegunIdAlmacen(mo_conexion As Connection, lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmDevuelveSaldosConLotesSegunIdAlmacen = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosConLotesSegunAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveSaldosConLotesSegunIdAlmacen = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmMovimientoFiltrarPorFechasYtipoConcepto(mo_conexion As Connection, ldFechaIni As Date, ldFechaFin As Date, lnidtipoConcepto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_conexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoFiltraPorFechasYtipoConcepto"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
    Set oRecordset = .Execute
    'Set oRecordset.ActiveConnection = Nothing
  End With
  Set FarmMovimientoFiltrarPorFechasYtipoConcepto = oRecordset
  Set oCommand = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function FarmMovimientoVentasFiltrarPorFechas(mo_conexion As Connection, ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_conexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoVentasFiltraPorFechas"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oRecordset = .Execute
    'Set oRecordset.ActiveConnection = Nothing
  End With
  Set FarmMovimientoVentasFiltrarPorFechas = oRecordset
  Set oCommand = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function



Function FarmMovimientoVentasDetalleSeleccionarPorAlmacenFechas(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, lnPagan As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoVentasDetalleSeleccionarPorAlmacenFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorAlmacenFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPagan", adInteger, adParamInput, 0, lnPagan): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasDetalleSeleccionarPorAlmacenFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmMovimientoDetalleSeleccionarDevolucionPorTipoConceptoAlmacen(ldFechaInicio As Date, ldFechaFin As Date, lnidtipoConcepto As Long, lnIdAlmacenDestino As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoDetalleSeleccionarDevolucionPorTipoConceptoAlmacen = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleSeleccionarDevolucionPorTipoConceptoAlmacen"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, lnidtipoConcepto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoDetalleSeleccionarDevolucionPorTipoConceptoAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino
Function DevuelveMovimientosDeNotaIngresos(lnIdAlmacenDestino As Long, ldFechaInicio As String, ldFechaFin As String, lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL1 As String: Dim sSQL2 As String
Dim oConexion As New ADODB.Connection
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
        Set DevuelveMovimientosDeNotaIngresos = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveMovimientosDeNotaIngresos"
        Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 50, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 50, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set DevuelveMovimientosDeNotaIngresos = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    
Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Modificado por Yamill Palomino
Function FarmMovimientoVentasFiltraFuentesFinanciamiento(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmMovimientoVentasFiltraFuentesFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasFiltraFuentesFinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDate, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDate, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasFiltraFuentesFinanciamiento = oRecordset
  
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino
Function FarmSaldoMensualSeleccionarUltimoSaldoPorIdproducto(lnIdProducto As Long, lnIdAlmacen As Long, ldFechaMov As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmSaldoMensualSeleccionarUltimoSaldoPorIdproducto = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmSaldoMensualSeleccionarUltimoSaldoPorIdproducto"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaMov", adDate, adParamInput, 10, ldFechaMov): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmSaldoMensualSeleccionarUltimoSaldoPorIdproducto = oRecordset

   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por YAmill Palomino
Function FarmRegeneraSaldosMayorAFecha(ldFechaCreacion As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmRegeneraSaldosMayorAFecha = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        Set oParameter = .CreateParameter("@fechaCreacion", adDate, adParamInput, 10, ldFechaCreacion): .Parameters.Append oParameter
        .CommandText = "FarmRegeneraSaldosMayorAFecha"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmRegeneraSaldosMayorAFecha = oRecordset
  
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



'********************************   Despacho de Donaciones a Pacientes  ***************************************
Function AgregaDatosDeDespachoDonaciones(ByVal oDoMovimiento As DoFarmMovimiento, oDoMovimientoDespachoDonaciones As DoFarmMovimientoDonaciones, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oFarmMovimientoDonaciones As New FarmMovimientoDonaciones
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String: Dim lcCorrelativoPrograma As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim oRsLookTablas As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    
    Set oMovimiento.Conexion = oConexion
    Set oFarmMovimientoDonaciones.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    'SCCQ 23/10/2020 Cambio28 Inicio
    'Antes:lcCorrelativoPrograma = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(23)   'Guia Interna de Salida
    lcCorrelativoPrograma = oMovimiento.DevuelveYactualizaCorrelativosDisponibleFarmTipoDocumentos(23)   'Guia Interna de Salida
    'SCCQ 23/10/2020 Cambio28 Fin
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
       .DocumentoNumero = lcCorrelativoPrograma
       .DocumentoIdtipo = 23   'Guia Interna de Salida
    End With
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarIS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghDonaciones, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarIS
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe Items" + Chr(13): GoTo TerminarIS
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarIS
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    '*********  graba tabla MovimientosDonaciones  ***************
    With oDoMovimientoDespachoDonaciones
       .movNumero = lcCorrelativo
    End With
    
    If Not oFarmMovimientoDonaciones.Insertar(oDoMovimientoDespachoDonaciones) Then
            bProcesoOK = False: GoTo TerminarIS
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarIS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarIS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
    
TerminarIS:
    If bProcesoOK Then
        AgregaDatosDeDespachoDonaciones = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeDespachoDonaciones = False
        ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oFarmMovimientoDonaciones = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function
Function ModificaDatosDeDespachoDonaciones(oDoMovimiento As DoFarmMovimiento, oDoMovimientoDespachoDonaciones As DoFarmMovimientoDonaciones, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oFarmMovimientoDonaciones As New FarmMovimientoDonaciones
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    Dim rs As Recordset
    Dim oRsLookTablas As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oFarmMovimientoDonaciones.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    
    If Not oMovimiento.modificar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarMIS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "M", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*************vuelve a chequear saldos ****************
    If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, _
                                            oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghDonaciones, "") Then
        ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
        bProcesoOK = False: GoTo TerminarMIS
    End If
    If oRsDetalleProductosConLotes.RecordCount = 0 Then
       ms_MensajeError = "Falta ITEMS": bProcesoOK = False: GoTo TerminarMIS
    End If
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad) Then
             ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + Trim(oRsDetalleProductosConLotes.Fields!nombreProducto)
             bProcesoOK = False: GoTo TerminarMIS
       End If
       oRsDetalleProductosConLotes.MoveNext
    Loop
    
   '*********  graba tabla MovimientosProgramas  ***************
    If Not oFarmMovimientoDonaciones.modificar(oDoMovimientoDespachoDonaciones) Then
            bProcesoOK = False: GoTo TerminarMIS
    End If
    '*********  Elimina tabla farmMovimientosDetalle y devuelve saldos de farmSaldo,farmSaldoDetalle  ***************
    Set rs = FarmMovimientosDetalleDevuelveTodosItems(oConexion, oDoMovimiento.MovTipo, oDoMovimiento.movNumero)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
           oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
           oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
           oDoMovimientoDetalle.idProducto = rs.Fields!idProducto
           oDoMovimientoDetalle.Lote = rs.Fields!Lote
           oDoMovimientoDetalle.FechaVencimiento = rs.Fields!FechaVencimiento
           oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
           oDoMovimientoDetalle.idTipoSalidaBienInsumo = rs.Fields!idTipoSalidaBienInsumo
           If Not oMovimientoDetalle.Eliminar(oDoMovimientoDetalle) Then
              bProcesoOK = False: GoTo TerminarMIS
           End If
           If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoMovimiento.idAlmacenOrigen, rs.Fields!idProducto, rs.Fields!Lote, rs.Fields!FechaVencimiento, rs.Fields!idTipoSalidaBienInsumo, rs.Fields!cantidad, rs.Fields!Precio, oDoMovimiento.FechaCreacion) Then
              bProcesoOK = False: GoTo TerminarMIS
           End If
           rs.MoveNext
        Loop
    End If
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    oRsDetalleProductosConLotes.MoveFirst
    Do While Not oRsDetalleProductosConLotes.EOF
       oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarMIS
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoSalida, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarMIS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductosConLotes.MoveNext
    Loop
TerminarMIS:
    If bProcesoOK Then
        ModificaDatosDeDespachoDonaciones = True
        oConexion.CommitTrans
    Else
        ModificaDatosDeDespachoDonaciones = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
    Set oFarmMovimientoDonaciones = Nothing
    Set rs = Nothing
    Set oRsLookTablas = Nothing
End Function
Function FarmMovimientoDespachoDonacionesSeleccionarPorId(ByVal oDoFarmMovimientoDespachoDonaciones As DoFarmMovimientoDonaciones) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimientoDespachoDonaciones As New FarmMovimientoDonaciones
    oConexion.Open sighentidades.CadenaConexion
    Set oMovimientoDespachoDonaciones.Conexion = oConexion
    If oMovimientoDespachoDonaciones.SeleccionarPorId(oDoFarmMovimientoDespachoDonaciones) Then
        FarmMovimientoDespachoDonacionesSeleccionarPorId = True
    Else
        FarmMovimientoDespachoDonacionesSeleccionarPorId = False
        ms_MensajeError = oMovimientoDespachoDonaciones.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimientoDespachoDonaciones = Nothing
End Function


'Modificado por Yamill Palomino
Function EmpleadosDevuelveSoloServicioSocial() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set EmpleadosDevuelveSoloServicioSocial = Nothing

     
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDevuelveSoloServicioSocial"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set EmpleadosDevuelveSoloServicioSocial = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function



Function FarmMovimientoDonacionesSeleccionarPorFechas(lnIdAlmacen As Long, ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmMovimientoDonacionesSeleccionarPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDevolucionesSeleccionarPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoDonacionesSeleccionarPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function CargaUltimoCorrelativoIdAlmacen() As Long
    Dim oFarmAlmacen  As New farmAlmacen
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFarmAlmacen.Conexion = oConexion
    CargaUltimoCorrelativoIdAlmacen = oFarmAlmacen.CargaUltimoCorrelativoIdAlmacen
    
    Set oFarmAlmacen = Nothing
End Function

'Modificado por Yamill Palomino
Function DevuelveDiagnosticoSegunSubComponente(lnIdSubComponente As Long, ByRef lcTipoSOPORTE As String) As String
        Dim oRsBuscaDx As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        DevuelveDiagnosticoSegunSubComponente = ""
        lcTipoSOPORTE = ""
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "DevuelveDiagnosticoSegunSubComponente"
             Set oParameter = .CreateParameter("@IdSubComponente", adInteger, adParamInput, 0, lnIdSubComponente): .Parameters.Append oParameter
             Set oRsBuscaDx = .Execute
             Set oRsBuscaDx.ActiveConnection = Nothing
        End With
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
   
   
            If oRsBuscaDx.RecordCount > 0 Then
               If Not IsNull(oRsBuscaDx.Fields!diagnostico) Then
                  DevuelveDiagnosticoSegunSubComponente = oRsBuscaDx.Fields!diagnostico
                  If Not IsNull(oRsBuscaDx.Fields!tipoProductoSISMED) Then
                     If Trim(oRsBuscaDx.Fields!tipoProductoSISMED) <> "" Then
                        lcTipoSOPORTE = oRsBuscaDx.Fields!tipoProductoSISMED
                     End If
                  End If
               End If
            End If
            oRsBuscaDx.Close
     
        Set oRsBuscaDx = Nothing
End Function


'Modificado por Yamill Palomino
Function DevuelveSiElInventarioYaExiste(lcNroInventario As String) As Long
        Dim oRsTmp As New Recordset
        Dim lcSql As String
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveSiElInventarioYaExiste"
        Set oParameter = .CreateParameter("@NroInventario", adVarChar, adParamInput, 50, lcNroInventario): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   DevuelveSiElInventarioYaExiste = oRsTmp.RecordCount
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
        
End Function

'Modificado por Yamill Palomino
Sub FiltraSaldoParaElAjuste(lcDctoAjuste As String, lnIdAlmacenA As Long, ByRef oRsSaldosAjuste As Recordset)
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmSaldoTotalesSoloMayoresAcero"
        Set oParameter = .CreateParameter("@DctoAjuste", adVarChar, adParamInput, 50, lcDctoAjuste): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacenA", adInteger, adParamInput, 0, lnIdAlmacenA): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    If oRsTmp.RecordCount > 0 Then
       '************ya esta registrado el Documento
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!Codigo = oRsTmp.Fields!Codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!saldo = oRsTmp.Fields!cantidad
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    Else
       '************registran por  primera vez
       oRsTmp.Close
       

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultaAjusteXAlmacen"
        Set oParameter = .CreateParameter("@IdAlmacenA", adInteger, adParamInput, 0, lnIdAlmacenA): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oParameter = Nothing
   Set oCommand = Nothing
       
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!Codigo = oRsTmp.Fields!Codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Fields!saldo = oRsTmp.Fields!cantidad
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    End If
    Set oRsTmp = Nothing
   oConexion.Close
   Set oConexion = Nothing

End Sub

'Modificado por Yamill Palomino

Function FarmRelModDevuelveSegunFiltro(lcFiltro As String) As Recordset
       Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmRelModDevuelveSegunFiltro = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmRelModDevuelveSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 300, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set FarmRelModDevuelveSegunFiltro = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
       
       
End Function

Sub FarmRelModActualizaSegunFiltro(lcFiltro As String, lcDocumentoUltimoNumero As String)
       Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmRelModActualizaSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcDocumento", adVarChar, adParamInput, 20, lcDocumentoUltimoNumero): .Parameters.Append oParameter
         .Execute
   End With
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
       
       
End Sub

Function FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes(lnIdProducto As Long, lnIdAlmacen As Long, _
                                                                  ldFechaMov As Date, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaMov", adDate, adParamInput, 10, ldFechaMov): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FarmSaldoMensualSeleccionarUltimoSaldoPorIdproductoXmes = oRecordset
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub ActualizaCorrelativosPorAnio(lcAnio As String)
        On Error GoTo errACF
        Dim oRsTmp As New Recordset, lcSql As String
        Dim oConexion As New ADODB.Connection
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oConexion.BeginTrans
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "ActualizaCorrelativosPorAnio"
             Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 50, lcAnio): .Parameters.Append oParameter
             Set oRsTmp = .Execute
             Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
        If oRsTmp.RecordCount > 0 Then
            oRsTmp.Close
            If MsgBox("Ya existe MOVIMIENTOS para ese año a aperturar" & Chr(13) & Chr(13) & "Esta seguro de Aperturar ?", vbQuestion + vbYesNo, "") = vbYes Then

                With oCommand
                     .CommandType = adCmdStoredProc
                     Set .ActiveConnection = oConexion
                     .CommandTimeout = 150
                     .CommandText = "FarmRelModYParametrosModificados"
                     Set oParameter = .CreateParameter("@DocumentoUltimoNumero", adVarChar, adParamInput, 20, Right(lcAnio, 2) & "00"): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 355, lcAnio): .Parameters.Append oParameter
                     Set oRsTmp = .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
        
            End If
        Else
            oRsTmp.Close
            With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = oConexion
               .CommandTimeout = 150
                     .CommandText = "FarmTipoDocumentosModificaCorrelativo"
                     Set oRsTmp = .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = oConexion
               .CommandTimeout = 150
                     .CommandText = "FarmRelModYParametrosModificados"
                     Set oParameter = .CreateParameter("@DocumentoUltimoNumero", adVarChar, adParamInput, 20, Right(lcAnio, 2) & "00"): .Parameters.Append oParameter
                     Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 355, lcAnio): .Parameters.Append oParameter
                     Set oRsTmp = .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
                
        End If
        Set oRsTmp = Nothing
        oConexion.CommitTrans
        oConexion.Close
        Set oConexion = Nothing
        Set mo_ReglasSeguridad = Nothing
        Set oCommand = Nothing
        
        Exit Sub
errACF:
        MsgBox Err.Description
        oConexion.RollbackTrans
End Sub

' Modificado por Yamill Palomino
Function FarmMovimientoSeleccionarPorFechasCabecera(lcFechaInicial As String, lcHoraInicial As String, lcFechaFinal As String, lcHoraFinal As String, lnTodasFarmacias As Integer, lnIdAlmacen As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FarmMovimientoSeleccionarPorFechasCabecera = Nothing
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoSeleccionarPorFechasCabecera"
        Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, lcFechaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@HoraInicial", adVarChar, adParamInput, 50, lcHoraInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, lcFechaFinal): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@HoraFinal", adVarChar, adParamInput, 50, lcHoraFinal): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TodasFarmacias", adInteger, adParamInput, 0, lnTodasFarmacias): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmMovimientoSeleccionarPorFechasCabecera = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

' MODIFICADO POR YAMILL PALOMINO
Function SaldoTotalDelEstablecimiento() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset


Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set SaldoTotalDelEstablecimiento = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SaldoTotalDelEstablecimiento"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set SaldoTotalDelEstablecimiento = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function


'MODIFICADO POR YAMILL PALOMINO
Function SaldoDetalladoPorItemSeleccionarPorCodigo(lcCodigoItem As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set SaldoDetalladoPorItemSeleccionarPorCodigo = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SaldoDetalladoPorItemSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@lcCodigoItem", adVarChar, adParamInput, 50, lcCodigoItem): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set SaldoDetalladoPorItemSeleccionarPorCodigo = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
        
        
End Function

' MODIFICADO POR YAMILL PALOMINO
Function FarmMovimientoVentasSeleccionarPorCuenta(lnIdCuentaAtencion As Long) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmMovimientoVentasSeleccionarPorCuenta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasSeleccionarPorCuenta = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
        
        
End Function

Sub SaldoEliminarTodosSegunFiltro(lcWhere As String, oConexion As Connection)
'            Dim lcSql As String, oRsTmp As New Recordset
            Dim oCommand As New ADODB.Command
            Dim oParameter As ADODB.Parameter
            
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "SaldoEliminarTodosSegunFiltro"
              Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 500, lcWhere): .Parameters.Append oParameter
              .Execute
            End With
            Set oCommand = Nothing
End Sub

Sub SaldoMensualEliminarTodosSegunFiltro(lcWhere As String, oConexion As Connection)
'            Dim lcSql As String, oRsTmp As New Recordset
            Dim oCommand As New ADODB.Command
            Dim oParameter As ADODB.Parameter
            
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "SaldoMensualEliminarTodosSegunFiltro"
              Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 500, lcWhere): .Parameters.Append oParameter
              .Execute
            End With
            Set oCommand = Nothing
End Sub

Sub FarmMovimientoDetalleActualizaEspaciosDeCadaLote(oConexion As Connection)
'            Dim oCommand As New ADODB.Command
'            With oCommand
'              .CommandType = adCmdStoredProc
'              Set .ActiveConnection = oConexion
'              .CommandTimeout = 900
'              .CommandText = "FarmMovimientoDetalleActualizaEspaciosDeCadaLote"
'              .Execute
'            End With
'            Set oCommand = Nothing
End Sub

'MODIFICADO POR YAMILL PALOMINO

Function farmMovimientoVentasDetalleSeleccionarXidCuenta(ml_idCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
     
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmMovimientoVentasDetalleSeleccionarXidCuenta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasDetalleSeleccionarXidCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasDetalleSeleccionarXidCuenta = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
        
    
End Function

'Modificado por Yamill Palomino
Function FarmaciaFiltraTodosMovimientos(lcWhere As String, oConexion As Connection) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oParameter As ADODB.Parameter
    Dim oCommand As New ADODB.Command
    
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFiltraTodosMovimientos"
        Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 1000, lcWhere): .Parameters.Append oParameter
           Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmaciaFiltraTodosMovimientos = oRsTmp1
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino
Function FarmSaldoMensualDetalladoSeleccionarPorSAldoFecha(lcSaldoFecha As String, lcWhere As String, oConexion As Connection) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FarmSaldoMensualDetalladoSeleccionarPorSAldoFecha = Nothing


         With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmSaldoMensualDetalladoSeleccionarPorSAldoFecha"
        Set oParameter = .CreateParameter("@SaldoFecha", adVarChar, adParamInput, 50, lcSaldoFecha): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        If lcWhere <> "" Then
           lcWhere = Mid(lcWhere, 5)
           oRsTmp1.Filter = lcWhere
        End If
   End With
   Set FarmSaldoMensualDetalladoSeleccionarPorSAldoFecha = oRsTmp1
   Set oRsTmp1 = Nothing
   Set oCommand = Nothing
    
End Function


'Modificado por Franklin Cachay 29102013
Function FarmSaldoMensualDetalladoSeleccionarPorFiltro(lcWhere As String, oConexion As Connection) As Recordset
   
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsTmp1 As New Recordset
    Set FarmSaldoMensualDetalladoSeleccionarPorFiltro = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmSaldoMensualDetalladoSeleccionarPorFiltro"
        If Right(Trim(lcWhere), 3) = "and" Then
           lcWhere = Left(lcWhere, Len(lcWhere) - 4)
        End If
        Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 1500, lcWhere): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
   End With
   Set FarmSaldoMensualDetalladoSeleccionarPorFiltro = oRsTmp1
   Set oCommand = Nothing
'   Set oRsTmp1 = Nothing
   
End Function


Sub FarmaciaActualizaFechaDeVencimiento(ldfechaVencimientoNueva As Date, lnIdProducto As Long, lcLote As String, lcFechaVencimientoActual As Date)

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    
    With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FarmaciaActualizaFechaDeVencimiento"
                Set oParameter = .CreateParameter("@fechaVencimientoNueva", adDBTimeStamp, adParamInput, 0, Format(ldfechaVencimientoNueva, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaVencimientoActual", adDBTimeStamp, adParamInput, 0, Format(lcFechaVencimientoActual, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                .Execute
    End With
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    
End Sub


'MODIFICADO POR YAMILL PALOMINO
Function FarmaciaMuestraSaldosPorCodigo(lcCodigo As String) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmaciaMuestraSaldosPorCodigo = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaMuestraSaldosPorCodigo"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 50, lcCodigo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmaciaMuestraSaldosPorCodigo = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

'modificado por yamill palomino
Function FacturacionCuentasAtencionFiltrarXcuenta(lnIdCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionFiltrarXcuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FacturacionCuentasAtencionFiltrarXcuenta = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing

End Function


' MODIFICADO POR YAMILL PALOMINO
Function SaldosSegunOrden(lnOrdenadoPor As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SaldosSegunOrden"
        Set oParameter = .CreateParameter("@OrdenadoPor", adInteger, adParamInput, 0, lnOrdenadoPor): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set SaldosSegunOrden = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

'MODIFICADO POR YAMILL PALOMINO

Function FarmMovimientoVentasFiltrarMovnumeroIdTipoFinanc(lcMovNumero As String, lcMovTipo As String, _
                                                      lnIdtipoFinanciamiento As Long, Optional oConexion1 As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset

    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasFiltrarMovnumeroIdTipoFinanc"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@LcMovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRs = .Execute
        If oConexion1 Is Nothing Then
           Set oRs.ActiveConnection = Nothing
        End If
   End With
   Set FarmMovimientoVentasFiltrarMovnumeroIdTipoFinanc = oRs
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRs = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function

'MODIFICADO POR YAMILL PALOMINO
Function farmMovimientoVentasFiltrarMovnumero(lcMovNumero As String) As Recordset
    Dim lcSql As String, oRs As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmMovimientoVentasFiltrarMovnumero = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasFiltrarMovnumero"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasFiltrarMovnumero = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function

'Modificado por Yamill Palomino

Function FarmMovimientoVentasFiltrarFechasAlmacIdTipoServ1(mda_FechaInicio As Date, mda_FechaFin As Date, _
                             lc_TipoServicioHosp As String, lnIdAlmacen As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    On Error GoTo errFMV
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasFiltrarFechasAlmacIdTipoServ"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoServicioHosp", adInteger, adParamInput, 0, Val(lc_TipoServicioHosp)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FarmMovimientoVentasFiltrarFechasAlmacIdTipoServ1 = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
errFMV:
  MsgBox Err.Description
End Function

'modificado por yamill palomino

Function FarmInventarioSeleccionarXdocumento(ml_Documento As String, lnIdAlmacenDestino As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmInventarioSeleccionarXdocumento = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmInventarioSeleccionarXdocumento"
        Set oParameter = .CreateParameter("@Documento", adVarChar, adParamInput, 50, ml_Documento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FarmInventarioSeleccionarXdocumento = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

'Modificado por Yamill PAlomino
Sub ActualizaImporteDeCabeceraMovimientos(mda_FechaInicio As Date, mda_FechaFin As Date)
    Dim oRsAct1 As New Recordset
    Dim oRsAct2 As New Recordset
    Dim oConexion As New Connection
    Dim lnidtipoConcepto As Long
    Dim LcTexto1 As String, lcTexto2 As String, lcTexto3 As String
    Dim lcNombre As String, lnPrecio As Double, ml_Importe As Double, lbPrimeraVez As Boolean
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ActualizaImporteDeCabeceraMovimientos"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(mda_FechaInicio, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(mda_FechaFin, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)): .Parameters.Append oParameter
        Set oRsAct1 = .Execute
        Set oRsAct1.ActiveConnection = Nothing
   End With
   Set oParameter = Nothing
   Set oCommand = Nothing
            
               
            If oRsAct1.RecordCount > 0 Then
       oRsAct1.MoveFirst
       Do While Not oRsAct1.EOF
          lnidtipoConcepto = oRsAct1.Fields!idTipoConcepto
          LcTexto1 = oRsAct1.Fields!MovTipo
          lcTexto2 = oRsAct1.Fields!movNumero
          lcNombre = Format(oRsAct1.Fields!FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY) & " - " & IIf(IsNull(oRsAct1.Fields!DocumentoNumero), "", oRsAct1.Fields!DocumentoNumero)
          lnPrecio = oRsAct1.Fields!totalMovimiento
          ml_Importe = 0
          Do While Not oRsAct1.EOF And LcTexto1 = oRsAct1.Fields!MovTipo And lcTexto2 = oRsAct1.Fields!movNumero
             ml_Importe = ml_Importe + Round(oRsAct1.Fields!cantidad * oRsAct1.Fields!Precio, 2)
             oRsAct1.MoveNext
             If oRsAct1.EOF Then
                Exit Do
             End If
          Loop
          If ml_Importe <> lnPrecio Then
If LcTexto1 = "S" Then
lnPrecio = 0
End If

' Modificado por Yamill Palomino
             lbPrimeraVez = True
             If lnidtipoConcepto = 10 Then

  
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultaPreVentaXfarmMovimientoVentas"
        Set oParameter = .CreateParameter("@Texto2", adVarChar, adParamInput, 9, lcTexto2): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Texto1", adVarChar, adParamInput, 9, LcTexto1): .Parameters.Append oParameter
        Set oRsAct2 = .Execute
        Set oRsAct2.ActiveConnection = Nothing
   End With
   Set oParameter = Nothing
   Set oCommand = Nothing
                
                       
                If oRsAct2.RecordCount > 0 Then
                   If oRsAct2.Fields!idPreVenta > 0 Then
                      lbPrimeraVez = False
                   End If
                End If
                oRsAct2.Close
             Else
If LcTexto1 = "S" Then
lnPrecio = 0
End If
             
             End If
             '
             If lbPrimeraVez = True Then
             
                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "FarmMovimientoModificarTotal"
                  Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, FormatCurrency(ml_Importe, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcTexto2): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcTexto1): .Parameters.Append oParameter
                  .Execute
                End With
                Set oParameter = Nothing
                Set oCommand = Nothing
  
             End If
          End If
       Loop
    End If
    oRsAct1.Close
    Set oRsAct1 = Nothing
    Set oRsAct2 = Nothing
    Set oCommand = Nothing
    oConexion.Close
    Set oConexion = Nothing
End Sub

Sub PreparaTablasDBF(ByRef oRsFox As Recordset, ByRef oRsFox1 As Recordset, ByRef oRsFox2 As Recordset, _
                      lc_CodigoFarmacia As String, lcAnioMes As String, oConexionFox As Connection, lcAnioMes1 As String, _
                      lbEliminaDatosDeTablaFormato As Boolean)
        Dim LcTexto1 As String
        
        If lbEliminaDatosDeTablaFormato = True Then
            LcTexto1 = "delete from formato where codigo_pre='" & lc_CodigoFarmacia & "' and (annomes>='" & lcAnioMes & "' and annomes<='" & lcAnioMes1 & "')"
            oRsFox.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        End If
        
        LcTexto1 = "update formdet set sit='9' where codigo_pre='" & lc_CodigoFarmacia & "'"
        oRsFox.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        LcTexto1 = "delete from formdet where codigo_pre='" & lc_CodigoFarmacia & "' and (annomes>='" & lcAnioMes & "' and annomes<='" & lcAnioMes1 & "')"
        oRsFox.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox.Open "select * from formDet", oConexionFox, adOpenKeyset, adLockOptimistic
        
        LcTexto1 = "update formdetL set sit='9' where codigo_pre='" & lc_CodigoFarmacia & "'"
        oRsFox1.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        LcTexto1 = "delete from formdetL where codigo_pre='" & lc_CodigoFarmacia & "' and (annomes>='" & lcAnioMes & "' and annomes<='" & lcAnioMes1 & "')"
        oRsFox1.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox1.Open "select * from formDetL", oConexionFox, adOpenKeyset, adLockOptimistic
        
        LcTexto1 = "update formdetM set sit='9' where codigo_pre='" & lc_CodigoFarmacia & "'"
        oRsFox2.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        LcTexto1 = "delete from formdetM where codigo_pre='" & lc_CodigoFarmacia & "' and (annomes>='" & lcAnioMes & "' and annomes<='" & lcAnioMes1 & "')"
        oRsFox2.Open LcTexto1, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox2.Open "select * from formDetM", oConexionFox, adOpenKeyset, adLockOptimistic
        


  
End Sub


'modificado por yamill palomino

Function farmSaldoMensualFiltrarFechaYalmacen(ldFechaHistoricoXmes As Date, lnIdAlmacen As Long) As Recordset

    Dim lcSql As String, oRs As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmSaldoMensualFiltrarFechaYalmacen = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmSaldoMensualFiltrarFechaYalmacen"
        Set oParameter = .CreateParameter("@FechaHistoricoXmes", adDBTimeStamp, adParamInput, 0, ldFechaHistoricoXmes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set farmSaldoMensualFiltrarFechaYalmacen = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function
' Modificado por Yamill Palomino
'debb-15/09/2015
Function farmSaldoMensualFiltrarFecha(ldFechaHistoricoXmes As Date, lc_AlmacenesParaICI As String, lnIdAlmacenRep As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
    Dim lcWhere As String
    lcWhere = IIf(InStr(lc_AlmacenesParaICI, "/") > 0, "", "and dbo.farmSaldoMensual.idAlmacen=" & lnIdAlmacenRep)
    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmSaldoMensualFiltrarFecha"
        Set oParameter = .CreateParameter("@Where", adVarChar, adParamInput, 2500, lcWhere): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   lcSql = ""
   If lcWhere = "" Then
      'Todas las farmacias
      lcSql = "saldoFecha='" & Format(ldFechaHistoricoXmes, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) & "'" & _
              " and idAlmacen<>2"
   Else
      'Una farmacia
      lcSql = "saldoFecha='" & Format(ldFechaHistoricoXmes, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) & "'"
   End If
   oRs.Filter = lcSql
   
   Set farmSaldoMensualFiltrarFecha = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

Sub dbfFormatoSeleccionarTodos(ByRef oRsTmp As Recordset, oConexionFox As Connection)


    Dim lcSql As String, oRs As New Recordset
    lcSql = "select * from formato"
    If oRsTmp.State = 1 Then oRsTmp.Close
    oRsTmp.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
End Sub

'MODIFICADO POR YAMILL PALOMINO

Function farmMovimientoFiltrarXfechas(mda_FechaInicio As Date, mda_FechaFin As Date, oConexion As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset

    Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set farmMovimientoFiltrarXfechas = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoFiltrarXfechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set farmMovimientoFiltrarXfechas = oRs
   Set oCommand = Nothing
   Set oRs = Nothing
    
    
    
End Function

'MODIFICADO POR YAMILL PALOMINO

Function farmMovimientoFiltrarIntervSanitaria(mda_FechaInicio As Date, mda_FechaFin As Date, oConexion As Connection) As Recordset
Dim lcSql As String, oRs As New Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set farmMovimientoFiltrarIntervSanitaria = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoFiltrarIntervSanitaria"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set farmMovimientoFiltrarIntervSanitaria = oRs
   Set oCommand = Nothing
   Set oRs = Nothing
End Function

'modificado por yamill palomino

Function farmMovimientoDetalleSeleccionarSalidasXitem(lcMovNumeroS As String, lnIdProducto As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmMovimientoDetalleSeleccionarSalidasXitem = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleSeleccionarSalidasXitem"
        Set oParameter = .CreateParameter("@MovNumeroS", adVarChar, adParamInput, 50, lcMovNumeroS): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set farmMovimientoDetalleSeleccionarSalidasXitem = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function


'modificado por yamill palomino
Function SaldosNegativos() As Recordset
    Dim lcSql As String, oRs As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set SaldosNegativos = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SaldosNegativos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set SaldosNegativos = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
    
End Function

'modificado por yamill palomino

Function farmMovimientoVentasXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set farmMovimientoVentasXcuenta = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasXcuenta"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set farmMovimientoVentasXcuenta = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
        
End Function

'MODIFICADO POR YAMILL PALOMINO

Function DonacionesXcuenta(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim lcSql As String, oRs As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set DonacionesXcuenta = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DonacionesXcuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set DonacionesXcuenta = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function
    

'MODIFICADO POR YAMILL PALOMINO

Function FacturacionCuentasAtencionXidCuenta(ml_idCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRs As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FacturacionCuentasAtencionXidCuenta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionXidCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter '        Set oRsTmp2 = .Execute
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FacturacionCuentasAtencionXidCuenta = oRs
   oConexion.Close
   Set oRs = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function

'MODIFICADO POR YAMILL PALOMINO

Function farmMovimientoVentasDetalleXcuenta(ml_idCuentaAtencion As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmMovimientoVentasDetalleXcuenta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasDetalleXcuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter '        Set oRsTmp2 = .Execute
         Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasDetalleXcuenta = oRsTmp2
   oConexion.Close
   Set oRsTmp2 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
    
End Function


'MODIFICADO POR YAMILL

Function FacturacionBienesFinanciamientosXcuenta(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionBienesFinanciamientosXcuenta = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosXcuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
   End With
   Set FacturacionBienesFinanciamientosXcuenta = oRsTmp2
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
End Function

'MODIFICADO POR YAMILL PALOMINO

Function FacturacionBienesPagosXCuenta(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set FacturacionBienesPagosXCuenta = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosXCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
   End With
   Set FacturacionBienesPagosXCuenta = oRsTmp2
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing

End Function


'MODIFICADO POR YAMILL PALOMINO
Function FacturacionBienesDevolucionesXcuenta(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set FacturacionBienesDevolucionesXcuenta = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesDevolucionesXcuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
   End With
   Set FacturacionBienesDevolucionesXcuenta = oRsTmp2
   Set oRsTmp2 = Nothing
   Set oCommand = Nothing
    
End Function


'MODIFICADO POR YAMILL PALOMINO
Function farmPreVentaXidComprobantePago(lnIdComprobantePago As Long, Optional oConexion1 As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set farmPreVentaXidComprobantePago = Nothing
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "farmPreVentaXidComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp2.ActiveConnection = Nothing
        End If
   End With
   Set farmPreVentaXidComprobantePago = oRsTmp2
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp2 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

'MODIFICADO POR YAMILL PALOMINO

Function farmSaldoSoloFarmaciasSismed(lnIdProducto As Long, Optional oConexion1 As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "farmSaldoSoloFarmaciasSismed"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp2.ActiveConnection = Nothing
        End If
   End With
   Set farmSaldoSoloFarmaciasSismed = oRsTmp2
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp2 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function


' yamill falta
Function farmSaldoTotalesSoloMayoresAcero(lcFiltro As String, lnOrdenadoPor As sghTipoOrden) As Recordset
   Dim sSql As String, oRsTmp2 As New Recordset
    


Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmSaldoTotalesSoloMayoresAcero = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmSaldoTotalesSoloMayoresAcero"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, IIf(lcFiltro = "", "", " and " & lcFiltro)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Order", adVarChar, adParamInput, 50, IIf(lnOrdenadoPor = sghPorCodigo, " ORDER BY dbo.FactCatalogoBienesInsumos.Codigo", " ORDER BY dbo.FactCatalogoBienesInsumos.Nombre")): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
        oRsTmp2.Filter = "idTipoLocales='F'"
   End With
   Set farmSaldoTotalesSoloMayoresAcero = oRsTmp2
   oConexion.Close
   Set oRsTmp2 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function



Function ChequeaQueSaldosSeaPositivo(oConexion As ADODB.Connection, lnIdAlmacen As Long, lnIdProducto As Long, _
                                     lnIdTipoSalidaBienInsumo As Long, cantidad As Long) As Boolean
    Dim oRsTmp As New Recordset
    ChequeaQueSaldosSeaPositivo = False
    Set oRsTmp = farmDevuelveSaldosSegunAlmacenProducto(lnIdAlmacen, lnIdProducto, oConexion)
    oRsTmp.Filter = "IdTipoSalidaBienInsumo=" & lnIdTipoSalidaBienInsumo
    If oRsTmp.RecordCount > 0 Then
       If oRsTmp.Fields!cantidad >= cantidad Then
          ChequeaQueSaldosSeaPositivo = True
       End If
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function

Function FarmAlmacenConsultaXtipoFarmacia() As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Set FarmAlmacenConsultaXtipoFarmacia = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmAlmacenConsultaXtipoFarmacia"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FarmAlmacenConsultaXtipoFarmacia = oRs
   oConexion.Close
   Set oRs = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino
Function FarmAlmacenObtieneCodigoSismed(lcidTipoSuministro As String) As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FarmAlmacenObtieneCodigoSismed = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmAlmacenObtieneCodigoSismed"
        Set oParameter = .CreateParameter("@idTipoSuministro", adVarChar, adParamInput, 2, lcidTipoSuministro): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FarmAlmacenObtieneCodigoSismed = oRs
   oConexion.Close
   Set oRs = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function



Function ExportaPreciosSismedRegSant(lnIdProducto As Long, oConexion As Connection)
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ExportaPreciosSismedRegSant"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        'Set oRs.ActiveConnection = Nothing
   End With
   Set ExportaPreciosSismedRegSant = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function

Function ExportaPreciosSismedFecha(lnIdProducto As Long, oConexion As Connection)
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ExportaPreciosSismedFecha"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        'Set oRs.ActiveConnection = Nothing
   End With
   Set ExportaPreciosSismedFecha = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function

Function ExportaPreciosSismed(oConexion As Connection)
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ExportaPreciosSismed"
        Set oRs = .Execute
   End With
   Set ExportaPreciosSismed = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function


Function farmTipoProductosSismedDevuelveTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmTipoProductosSismedDevuelveTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmTipoProductosSismedDevuelveTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
   
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function


Function farmDevuelveSaldosSinLotesSegunAlmacenSOPORTE(lnIdAlmacen As Long, lnOrden As Long, _
                                                       lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosSinLotesSegunAlmacenSOPORTE"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmDevuelveSaldosSinLotesSegunAlmacenSOPORTE = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmMovimientosDetalleDevuelveTodosItemsSOPORTE(oConexion As ADODB.Connection, lcMovTipo As String, lcMovNumero As String) As ADODB.Recordset
    Dim oTabla As New FarmMovimientoDetalle
    Set oTabla.Conexion = oConexion
    Set FarmMovimientosDetalleDevuelveTodosItemsSOPORTE = oTabla.DevuelveTodosItemsSOPORTE(lcMovNumero, lcMovTipo)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'***debb-Setiembre2014
Function AcumulaRedondeoEnBoletas(ldFechaInicial As Date, ldFechaFinal As Date, oConexion As Connection) As Double
    Dim mrs_Tmp33 As New Recordset, lnRedondeo As Double, mrs_Tmp34 As New Recordset
    Dim lcSql As String, mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    lnRedondeo = 0
    Set mrs_Tmp33 = farmMovimientoDetalleSumaTotales(ldFechaInicial, ldFechaFinal, oConexion)
    If mrs_Tmp33.RecordCount > 0 Then
       mrs_Tmp33.MoveFirst
       Do While Not mrs_Tmp33.EOF
          Set mrs_Tmp34 = mo_ReglasFarmacia.farmMovimientoSeleccionarPorMovNumero(mrs_Tmp33!movNumero, "S")
          If mrs_Tmp34.RecordCount > 0 Then
             lnRedondeo = lnRedondeo + (mrs_Tmp33!totalDetalle - mrs_Tmp34!total)
          End If
          mrs_Tmp34.Close
          mrs_Tmp33.MoveNext
       Loop
    End If
    mrs_Tmp33.Close
    Set mrs_Tmp33 = Nothing
    AcumulaRedondeoEnBoletas = lnRedondeo
End Function
Function farmMovimientoDetalleSumaTotales(mda_FechaInicio As Date, mda_FechaFin As Date, oConexion As Connection) As Recordset
Dim lcSql As String, oRs As New Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleSumaTotales"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRs = .Execute
   End With
   Set farmMovimientoDetalleSumaTotales = oRs
   Set oCommand = Nothing
   Set oRs = Nothing
End Function

'***debb-Setiembre2014

Function farmMovimientoVentasSeleccionarXMovimiento(lcMovTipo As String, lcMovNumero As String, oConexion As Connection) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "farmMovimientoVentasSeleccionarPorId"
     Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, lcMovNumero)
     .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, lcMovTipo)
   .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   Set farmMovimientoVentasSeleccionarXMovimiento = oRecordset
   Set oCommand = Nothing
End Function

Function FarmMovimientoDetalleSeleccionarXcodigo(lcCodigo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleSeleccionarXcodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmMovimientoDetalleSeleccionarXcodigo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub FarmaciaActualizaTipoSalida(lnIdProducto As Long, lnNuevoTipo As Long)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaActualizaTipoSalida"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nuevoTipo", adInteger, adParamInput, 0, lnNuevoTipo): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub


Function farmMovimientoVentasDetalleXidProducto(lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasDetalleXidProducto"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasDetalleXidProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'DEBB2014a
Function farmTipoInventarioSeleccionarTodos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmTipoInventarioSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmTipoInventarioSeleccionarTodos = oRecordset
        oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
   
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function
'DEBB2014a
Function farmInventarioSeleccionarPorIdAlmacen(lnIdAlmacen As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmInventarioSeleccionarPorIdAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmInventarioSeleccionarPorIdAlmacen = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'DEBB2014a
Function LaFarmaciaEstaRegenerandoSaldos(lnIdFarmacia As Long) As Boolean
     LaFarmaciaEstaRegenerandoSaldos = False
     Dim oRsTmp As New Recordset
     Dim lnDiaSemana As Integer
     Set oRsTmp = FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & lnIdFarmacia)
     If oRsTmp.RecordCount > 0 Then
        lnDiaSemana = Weekday(CDate(lcBuscaParametro.RetornaFechaServidorSQL))
        If Mid(oRsTmp!regenerarEstado, lnDiaSemana, 1) = "P" Then
           LaFarmaciaEstaRegenerandoSaldos = True
           MsgBox "En este momento se está REGENERANDO SALDOS", vbInformation, "FARMACIA"
        End If
     End If
     oRsTmp.Close
     Set oRsTmp = Nothing
End Function

'DEBB2014b
Function FarmHistPrecioSeleccionarPorCodigo(lcCodigo As String, lcNombre As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmHistPrecioSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 100, lcNombre): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmHistPrecioSeleccionarPorCodigo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb2014b
Function farmMovimientoSeleccionarPorIdAlmacenDestino(lnIdAlmacenDestino As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoSeleccionarPorIdAlmacenDestino"
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoSeleccionarPorIdAlmacenDestino = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'debb2014b
Function farmHistPrecioSeleccionarPorId(lnIdHistPrecio As Long) As DoFarmHistPrecio
    Dim oConexion As New ADODB.Connection
    Dim oFarmHistPrecio As New FarmHistPrecio
    Dim oDoFarmHistPrecio As New DoFarmHistPrecio
    Set farmHistPrecioSeleccionarPorId = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmHistPrecio.Conexion = oConexion
    oDoFarmHistPrecio.idHistPrecio = lnIdHistPrecio
    If oFarmHistPrecio.SeleccionarPorId(oDoFarmHistPrecio) Then
        Set farmHistPrecioSeleccionarPorId = oDoFarmHistPrecio
    End If
    oConexion.Close
    ms_MensajeError = oFarmHistPrecio.MensajeError
    Set oConexion = Nothing
    Set oFarmHistPrecio = Nothing
End Function
'debb2014b
Function FarmHistPrecioAgregar(oDoFarmHistPrecio As DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS As Long, _
                               mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oFarmHistPrecio As New FarmHistPrecio
    Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos, oCatalogoBienesInsumos As New CatalogoBienesInsumos
    On Error GoTo ErrAdd
    FarmHistPrecioAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oFarmHistPrecio.Conexion = oConexion
    Set oCatalogoBienesInsumos.Conexion = oConexion
    If oFarmHistPrecio.Insertar(oDoFarmHistPrecio) Then
        oDOCatalogoBienesInsumos.idProducto = oDoFarmHistPrecio.idProducto
        If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) = True Then
            oDOCatalogoBienesInsumos.IdUsuarioAuditoria = oDoFarmHistPrecio.IdUsuarioAuditoria
            oDOCatalogoBienesInsumos.PrecioCompra = oDoFarmHistPrecio.PrecioCompra
            oDOCatalogoBienesInsumos.PrecioDistribucion = oDoFarmHistPrecio.PrecioDistribucion
            oDOCatalogoBienesInsumos.PrecioDonacion = oDoFarmHistPrecio.PrecioDonacion
            
            If oCatalogoBienesInsumos.modificar(oDOCatalogoBienesInsumos) = True Then
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFarmHistPrecio.IdUsuarioAuditoria, "A", _
                    oDoFarmHistPrecio.idHistPrecio, "FarmHistPrecio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                    "IdMedicamento/Insumo. " & Trim(Str(oDoFarmHistPrecio.idProducto)))           'ListBarItems.idListItem
                ActualizaPreciosEnFuentesFinanciamiento oDoFarmHistPrecio, oConexion
                FarmHistPrecioAgregar = True
            End If
        End If
    Else
        ms_MensajeError = oFarmHistPrecio.MensajeError
    End If
    oConexion.CommitTrans
    oConexion.Close
    FarmHistPrecioAgregar = True
ErrAdd:
    If FarmHistPrecioAgregar = False Then
       MsgBox Err.Description
       oConexion.RollbackTrans
    End If
    Set oConexion = Nothing
    Set oFarmHistPrecio = Nothing
    Set oDOCatalogoBienesInsumos = Nothing
    Set oCatalogoBienesInsumos = Nothing
End Function
'debb2014b
Function FarmHistPrecioModificar(oDoFarmHistPrecio As DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS As Long, _
                                 mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oFarmHistPrecio As New FarmHistPrecio
    Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos, oCatalogoBienesInsumos As New CatalogoBienesInsumos
    On Error GoTo ErrMod
    FarmHistPrecioModificar = False
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmHistPrecio.Conexion = oConexion
    Set oCatalogoBienesInsumos.Conexion = oConexion
    If oFarmHistPrecio.modificar(oDoFarmHistPrecio) = True Then
        oDOCatalogoBienesInsumos.idProducto = oDoFarmHistPrecio.idProducto
        If oCatalogoBienesInsumos.SeleccionarPorId(oDOCatalogoBienesInsumos) = True Then
            oDOCatalogoBienesInsumos.IdUsuarioAuditoria = oDoFarmHistPrecio.IdUsuarioAuditoria
            oDOCatalogoBienesInsumos.PrecioCompra = oDoFarmHistPrecio.PrecioCompra
            oDOCatalogoBienesInsumos.PrecioDistribucion = oDoFarmHistPrecio.PrecioDistribucion
            oDOCatalogoBienesInsumos.PrecioDonacion = oDoFarmHistPrecio.PrecioDonacion
            If oCatalogoBienesInsumos.modificar(oDOCatalogoBienesInsumos) = True Then
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFarmHistPrecio.IdUsuarioAuditoria, "M", _
                    oDoFarmHistPrecio.idHistPrecio, "FarmHistPrecio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                    "IdMedicamento/Insumo. " & Trim(Str(oDoFarmHistPrecio.idProducto)))
                ActualizaPreciosEnFuentesFinanciamiento oDoFarmHistPrecio, oConexion
                FarmHistPrecioModificar = True
            End If
        End If
    Else
        ms_MensajeError = oFarmHistPrecio.MensajeError
    End If
    oConexion.CommitTrans
    oConexion.Close
    FarmHistPrecioModificar = True
ErrMod:

    If FarmHistPrecioModificar = False Then
       MsgBox Err.Description
       oConexion.RollbackTrans
    End If
    Set oConexion = Nothing
    Set oFarmHistPrecio = Nothing
    Set oDOCatalogoBienesInsumos = Nothing
    Set oCatalogoBienesInsumos = Nothing
End Function
'debb2014b
Function FarmHistPrecioEliminar(oDoFarmHistPrecio As DoFarmHistPrecio, mo_lnIdTablaLISTBARITEMS As Long, _
                                mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oFarmHistPrecio As New FarmHistPrecio
    On Error GoTo ErrElim
    FarmHistPrecioEliminar = False
    oConexion.Open sighentidades.CadenaConexion
    Set oFarmHistPrecio.Conexion = oConexion
    If oFarmHistPrecio.Eliminar(oDoFarmHistPrecio) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFarmHistPrecio.IdUsuarioAuditoria, "E", oDoFarmHistPrecio.idHistPrecio, _
            "FarmHistPrecio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                "IdMedicamento/Insumo. " & Trim(Str(oDoFarmHistPrecio.idProducto)))
        FarmHistPrecioEliminar = True
    Else
        ms_MensajeError = oFarmHistPrecio.MensajeError
    End If
    oConexion.CommitTrans
    oConexion.Close
ErrElim:
    If Err.Number <> 0 Then
       MsgBox Err.Description
       oConexion.RollbackTrans
    End If
    Set oConexion = Nothing
    Set oFarmHistPrecio = Nothing
End Function
'debb2014b
Sub ActualizaPreciosEnFuentesFinanciamiento(oDoFarmHistPrecio As DoFarmHistPrecio, oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp As New Recordset
    Dim lcSql As String
    Dim lbNuevo As Boolean
    Dim oDoFinanciamientoCatalogoBien As New DoFinanciamientoCatalogoBien, oFinanciamientoCatalogoBien As New FinanciamientoCatalogoBien
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
    'Agrega Insumo en otros Tipo de Financiamiento y actualiza Precio
    Set oRsTmp = mo_ReglasFacturacion.CatalogoBienesInsumosHospSeleccionarXIdProducto(oDoFarmHistPrecio.idProducto)
    Set oRsTmp1 = mo_AdminComun.TiposFinanciamientoSegunFiltro("SeIngresPrecios=1")
    If oRsTmp1.RecordCount > 0 Then
       Set oFinanciamientoCatalogoBien.Conexion = oConexion
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lbNuevo = True
          If oRsTmp.RecordCount > 0 Then
             oRsTmp.MoveFirst
             oRsTmp.Find "idTipoFinanciamiento=" & oRsTmp1.Fields!IdTipoFinanciamiento
             If Not oRsTmp.EOF Then
                lbNuevo = False
             End If
          End If

          oDoFinanciamientoCatalogoBien.idProducto = oDoFarmHistPrecio.idProducto
          oDoFinanciamientoCatalogoBien.IdTipoFinanciamiento = oRsTmp1.Fields!IdTipoFinanciamiento
          oDoFinanciamientoCatalogoBien.Activo = 1
          oDoFinanciamientoCatalogoBien.PrecioUnitario = oDoFarmHistPrecio.PrecioVenta
          oDoFinanciamientoCatalogoBien.IdUsuarioAuditoria = oDoFarmHistPrecio.IdUsuarioAuditoria
          If lbNuevo = True Then
             If oFinanciamientoCatalogoBien.Insertar(oDoFinanciamientoCatalogoBien) = False Then
             End If
          Else
              oDoFinanciamientoCatalogoBien.IdPlanCatalogo = oRsTmp.Fields!IdPlanCatalogo
              If oFinanciamientoCatalogoBien.modificar(oDoFinanciamientoCatalogoBien) = False Then
              End If
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oDoFinanciamientoCatalogoBien = Nothing
    Set oFinanciamientoCatalogoBien = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_AdminComun = Nothing
End Sub
'debb2014b
Function FarmHistPrecioPorIdProducto(lnIdProducto As Long, oConexion As Connection) As Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmHistPrecioPorIdProducto"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRs = .Execute
        'Set oRs.ActiveConnection = Nothing
   End With
   Set FarmHistPrecioPorIdProducto = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
End Function

'debb2014b
Function FactCatalogoBienesInsumosSeleccionarXDescripYcodigo(lcCodigo As String, lcDescripcion As String) As Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarXDescripYcodigo"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcDescripcion", adVarChar, adParamInput, 100, lcDescripcion): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
   End With
   Set FactCatalogoBienesInsumosSeleccionarXDescripYcodigo = oRs
   Set oRs = Nothing
   Set oCommand = Nothing
   Set oConexion = Nothing
End Function
'AGREGADO MARIANO 07112014   debb2014b
Function FarmMovimientoVentasDetalleSeleccionarPorCuentaXmes(lnCuenta As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmMovimientoVentasDetalleSeleccionarPorCuentaXmes = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorCuentaXMes"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@idMes", adInteger, adParamInput, 0, lnmes): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmMovimientoVentasDetalleSeleccionarPorCuentaXmes = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'mariano 19112014  debb2014b
Function FarmMovimientoSeleccionPorAlmacenProductosIngresados(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacenDestino As Long, lnIdAlmacenOrigen As Long, oConexion As Connection) As ADODB.Recordset
'On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmMovimientoSeleccionPorAlmacenProductosIngresados = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoSeleccionPorAlmacenProductosIngresados"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacenOrigen): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmMovimientoSeleccionPorAlmacenProductosIngresados = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
'Exit Function
'ManejadorDeError:
 '  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
End Function

'mariano reporte con proveedores 19112014  debb2014b
Function FarmMovimientoSeleccionPorProveedorProductosIngresados(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacenDestino As Long, lnIdAlmacenOrigen As Long, lnIdProveedor As Long, oConexion As Connection) As ADODB.Recordset
'On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmMovimientoSeleccionPorProveedorProductosIngresados = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoSeleccionPorProveedorProductosIngresados"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacenOrigen", adInteger, adParamInput, 0, lnIdAlmacenOrigen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, lnIdProveedor): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmMovimientoSeleccionPorProveedorProductosIngresados = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
'Exit Function
'ManejadorDeError:
 '  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
End Function

'proveedores mariano   debb2014b
Function FarmProveedorSeleccionarSegunFiltro(lcFiltro As String, Optional oConexion1 As Connection) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New Proveedores
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    Set FarmProveedorSeleccionarSegunFiltro = oTabla.SeleccionarSegunFiltro(lcFiltro)
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'AGREGADO POR MARIANO 11112014  debb2014b
Function FarmInventarioSeleccionarXdocumentoAlmacen(ml_Documento As String, lnIdAlmacenDestino As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmInventarioSeleccionarXdocumentoAlmacen = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmInventarioSeleccionarXdocumentoGeneral"
        Set oParameter = .CreateParameter("@Documento", adVarChar, adParamInput, 50, ml_Documento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAlmacenDestino", adInteger, adParamInput, 0, lnIdAlmacenDestino): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmInventarioSeleccionarXdocumentoAlmacen = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function AgregaDatosDeInventario(oDoInventario As DOfarmInventario, mo_rsInventarioCabecera As Recordset, _
                                 mo_rsInventarioDetalle As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                 mo_lcNombrePc As String, lcInventarioTemporal As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New farmInventario
    Dim oInventarioCabecera As New FarmInventarioCabecera
    Dim oInventarioDetalle As New farmInventarioDetalle
    Dim oDoInventarioCabecera As New DoFarmInventarioCabecera
    Dim oDoInventarioDetalle As New DoFarmInventarioDetalle
    Dim oReglasCaja As New ReglasCaja
    Dim bProcesoOK As Boolean
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    'Bloqueo
    Set oInventario.Conexion = oConexion
    Set oInventarioCabecera.Conexion = oConexion
    Set oInventarioDetalle.Conexion = oConexion
    
    '*********  graba tabla farmInventario  ***************
    
    If Not oInventario.Insertar(oDoInventario) Then
        bProcesoOK = False: GoTo TerminarADI
    End If
    If oDoInventario.numeroInventario = "" Then
        If lcInventarioTemporal <> "" Then
           oDoInventario.numeroInventario = lcInventarioTemporal
        Else
           oDoInventario.numeroInventario = Right(Trim(Str(Year(oDoInventario.FechaCreacion))), 2) & Right("0" & oDoInventario.idInventario, 2)
        End If
        If Not oInventario.modificar(oDoInventario) Then
           bProcesoOK = False: GoTo TerminarADI
        End If
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoInventario.IdUsuarioAuditoria, "A", oDoInventario.idInventario, "FarmInventario", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")             'ListBarItems.idListItem
    '*********  graba tabla farmInventarioCabecera  ***************
    oDoInventarioCabecera.idInventario = oDoInventario.idInventario
    oDoInventarioCabecera.IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
    With mo_rsInventarioCabecera
        .MoveFirst
        Do While Not .EOF
           oDoInventarioCabecera.cantidad = .Fields!cantidad
           oDoInventarioCabecera.idProducto = .Fields!idProducto
           oDoInventarioCabecera.Precio = .Fields!Precio
           oDoInventarioCabecera.total = .Fields!total
           oDoInventarioCabecera.cantidadSaldo = .Fields!cantidadSaldo
           oDoInventarioCabecera.cantidadFaltante = .Fields!cantidadFaltante
           oDoInventarioCabecera.cantidadSobrante = .Fields!cantidadSobrante
           If Not oInventarioCabecera.Insertar(oDoInventarioCabecera) Then
                bProcesoOK = False: GoTo TerminarADI
           End If
           .MoveNext
        Loop
    End With
    '*********  graba tabla farmInventarioDetalle  ***************
    oDoInventarioDetalle.idInventario = oDoInventario.idInventario
    oDoInventarioDetalle.IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
    With mo_rsInventarioDetalle
        .MoveFirst
        Do While Not .EOF
           oDoInventarioDetalle.cantidad = .Fields!cantidad
           oDoInventarioDetalle.FechaVencimiento = .Fields!FechaVencimiento
           oDoInventarioDetalle.idProducto = .Fields!idProducto
           oDoInventarioDetalle.Lote = .Fields!Lote
           oDoInventarioDetalle.Precio = .Fields!Precio
           oDoInventarioDetalle.RegistroSanitario = .Fields!RegistroSanitario
           oDoInventarioDetalle.idTipoSalidaBienInsumo = .Fields!idTipoSalidaBienInsumo
           oDoInventarioDetalle.cantidadSaldo = .Fields!cantidadSaldo
           oDoInventarioDetalle.cantidadFaltante = .Fields!cantidadFaltante
           oDoInventarioDetalle.cantidadSobrante = .Fields!cantidadSobrante
           oDoInventarioDetalle.EsHistoricoSaldo = .Fields!EsHistoricoSaldo
           If Not oInventarioDetalle.Insertar(oDoInventarioDetalle) Then
                bProcesoOK = False: GoTo TerminarADI
           End If
           .MoveNext
        Loop
    End With
TerminarADI:
    If bProcesoOK Then
        AgregaDatosDeInventario = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeInventario = False
        ms_MensajeError = ""
        ms_MensajeError = ms_MensajeError + IIf(oInventario.MensajeError <> "", oInventario.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oInventarioCabecera.MensajeError <> "", oInventarioCabecera.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oInventarioDetalle.MensajeError <> "", oInventarioDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInventario = Nothing
    Set oInventarioCabecera = Nothing
    Set oInventarioDetalle = Nothing
    Set oDoInventarioCabecera = Nothing
    Set oDoInventarioDetalle = Nothing
    Set oReglasCaja = Nothing
End Function

Function ModificaDatosDeInventario(oDoInventario As DOfarmInventario, mo_rsInventarioCabecera As Recordset, mo_rsInventarioDetalle As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New farmInventario
    Dim oInventarioCabecera As New FarmInventarioCabecera
    Dim oInventarioDetalle As New farmInventarioDetalle
    Dim oDoInventarioCabecera As New DoFarmInventarioCabecera
    Dim oDoInventarioDetalle As New DoFarmInventarioDetalle
    Dim bProcesoOK As Boolean
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    Set oInventario.Conexion = oConexion
    Set oInventarioCabecera.Conexion = oConexion
    Set oInventarioDetalle.Conexion = oConexion
    '*********  graba tabla farmInventario  ***************
   
    If Not oInventario.modificar(oDoInventario) Then
        bProcesoOK = False: GoTo TerminarMDI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoInventario.IdUsuarioAuditoria, "M", oDoInventario.idInventario, "FarmInventario", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")             'ListBarItems.idListItem
    '*********  graba tabla farmInventarioCabecera  ***************
    
    oDoInventarioCabecera.idInventario = oDoInventario.idInventario
    oDoInventarioCabecera.IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
    If Not oInventarioCabecera.Eliminar(oDoInventarioCabecera) Then
        bProcesoOK = False: GoTo TerminarMDI
    End If
    With mo_rsInventarioCabecera
        .MoveFirst
        Do While Not .EOF
           oDoInventarioCabecera.cantidad = .Fields!cantidad
           oDoInventarioCabecera.idInventario = oDoInventario.idInventario
           oDoInventarioCabecera.idProducto = .Fields!idProducto
           oDoInventarioCabecera.IdUsuarioAuditoria = oDoInventario.idUsuario
           oDoInventarioCabecera.Precio = .Fields!Precio
           oDoInventarioCabecera.total = .Fields!total
           oDoInventarioCabecera.cantidadSaldo = .Fields!cantidadSaldo
           oDoInventarioCabecera.cantidadFaltante = .Fields!cantidadFaltante
           oDoInventarioCabecera.cantidadSobrante = .Fields!cantidadSobrante
           If Not oInventarioCabecera.Insertar(oDoInventarioCabecera) Then
                bProcesoOK = False: GoTo TerminarMDI
           End If
           .MoveNext
        Loop
    End With
    '*********  graba tabla farmInventarioDetalle  ***************
    
    oDoInventarioDetalle.idInventario = oDoInventario.idInventario
    oDoInventarioDetalle.IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
    If Not oInventarioDetalle.Eliminar(oDoInventarioDetalle) Then
        bProcesoOK = False: GoTo TerminarMDI
    End If
    With mo_rsInventarioDetalle
        .MoveFirst
        Do While Not .EOF
           oDoInventarioDetalle.cantidad = .Fields!cantidad
           oDoInventarioDetalle.FechaVencimiento = .Fields!FechaVencimiento
           oDoInventarioDetalle.idInventario = oDoInventario.idInventario
           oDoInventarioDetalle.idProducto = .Fields!idProducto
           oDoInventarioDetalle.Lote = .Fields!Lote
           oDoInventarioDetalle.Precio = .Fields!Precio
           oDoInventarioDetalle.RegistroSanitario = .Fields!RegistroSanitario
           oDoInventarioDetalle.idTipoSalidaBienInsumo = .Fields!idTipoSalidaBienInsumo
           oDoInventarioDetalle.cantidadSaldo = .Fields!cantidadSaldo
           oDoInventarioDetalle.cantidadFaltante = .Fields!cantidadFaltante
           oDoInventarioDetalle.cantidadSobrante = .Fields!cantidadSobrante
           oDoInventarioDetalle.EsHistoricoSaldo = .Fields!EsHistoricoSaldo
           If Not oInventarioDetalle.Insertar(oDoInventarioDetalle) Then
                bProcesoOK = False: GoTo TerminarMDI
           End If
           .MoveNext
        Loop
    End With
TerminarMDI:
    If bProcesoOK Then
        ModificaDatosDeInventario = True
        oConexion.CommitTrans
    Else
        ModificaDatosDeInventario = False
        ms_MensajeError = ""
        ms_MensajeError = ms_MensajeError + IIf(oInventario.MensajeError <> "", oInventario.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oInventarioCabecera.MensajeError <> "", oInventarioCabecera.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oInventarioDetalle.MensajeError <> "", oInventarioDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInventario = Nothing
    Set oInventarioCabecera = Nothing
    Set oInventarioDetalle = Nothing
    Set oDoInventarioCabecera = Nothing
    Set oDoInventarioDetalle = Nothing
End Function

Function AnulaInventario(oDoInventario As DOfarmInventario, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New farmInventario
    AnulaInventario = False
    oConexion.Open sighentidades.CadenaConexion
    Set oInventario.Conexion = oConexion
    If oInventario.modificar(oDoInventario) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoInventario.IdUsuarioAuditoria, "E", oDoInventario.idInventario, "FarmInventario", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")             'ListBarItems.idListItem
        AnulaInventario = True
    Else
        ms_MensajeError = oInventario.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oInventario = Nothing
End Function

Function CierraInventario(oDoInventario As DOfarmInventario, lnTotalInventario As Double, _
                          mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oInventario As New farmInventario
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimiento As New DoFarmMovimiento
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim oDOfarmMovimientoNotaIngreso As New DOfarmMovimientoNotaIngreso
    Dim oFarmMovimientoNotaIngreso As New FarmMovimientoNotaIngreso
    Dim lcCorrelativo As String
    Dim oRsInventarioDetalle As New ADODB.Recordset
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    Set oFarmMovimientoNotaIngreso.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(1)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .DocumentoIdtipo = 10   'farmTipoDocumentos.nombre=inventario
       .DocumentoNumero = oDoInventario.numeroInventario
       .idAlmacenDestino = oDoInventario.IdAlmacen
       .idAlmacenOrigen = 0   'farmAlmacen.descripcion=<<NINGUNO>>
       .idEstadoMovimiento = 1
       .idTipoConcepto = 19     'farmTipoConceptos.concepto=inventario
       .idUsuario = oDoInventario.IdUsuarioAuditoria
       .IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
       .movNumero = lcCorrelativo
       .MovTipo = lcConstanteMovimientoEntrada
       .total = lnTotalInventario
       .FechaCreacion = oDoInventario.FechaCierre
    End With
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarCI
    End If
    oDOfarmMovimientoNotaIngreso.movNumero = oDoMovimiento.movNumero
    oDOfarmMovimientoNotaIngreso.MovTipo = oDoMovimiento.MovTipo
    If Not oFarmMovimientoNotaIngreso.Insertar(oDOfarmMovimientoNotaIngreso) Then
            bProcesoOK = False: GoTo TerminarCI
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.MovTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoInventario.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
    lnItem = 1
    Set oRsInventarioDetalle = FarmInventarioDetalleDevuelveProductosLotesPorId(oDoInventario.idInventario)
    Do While Not oRsInventarioDetalle.EOF
       oDoMovimientoDetalle.cantidad = oRsInventarioDetalle.Fields!cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsInventarioDetalle.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsInventarioDetalle.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsInventarioDetalle.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsInventarioDetalle.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = IIf(IsNull(oRsInventarioDetalle.Fields!RegistroSanitario), "", oRsInventarioDetalle.Fields!RegistroSanitario)
       oDoMovimientoDetalle.total = Round(oRsInventarioDetalle.Fields!cantidad * oRsInventarioDetalle.Fields!Precio, 2)
       oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsInventarioDetalle.Fields!idTipoSalidaBienInsumo
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarCI
       End If
       If Not oMovimientoDetalle.ActualizaSaldosPorProducto(lcConstanteMovimientoEntrada, oDoInventario.IdAlmacen, oRsInventarioDetalle.Fields!idProducto, oRsInventarioDetalle.Fields!Lote, oRsInventarioDetalle.Fields!FechaVencimiento, oRsInventarioDetalle.Fields!idTipoSalidaBienInsumo, oRsInventarioDetalle.Fields!cantidad, oRsInventarioDetalle.Fields!Precio, oDoMovimiento.FechaCreacion) Then
                bProcesoOK = False: GoTo TerminarCI
       End If
       lnItem = lnItem + 1
       oRsInventarioDetalle.MoveNext
    Loop
    '*********  graba tabla farmInventario  ***************
    Set oInventario.Conexion = oConexion
    If Not oInventario.modificar(oDoInventario) Then
       bProcesoOK = False: GoTo TerminarCI
    End If
TerminarCI:
    If bProcesoOK Then
        CierraInventario = True
        oConexion.CommitTrans
    Else
        CierraInventario = False
        ms_MensajeError = ms_MensajeError + IIf(oInventario.MensajeError <> "", oInventario.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oFarmMovimientoNotaIngreso.MensajeError <> "", oFarmMovimientoNotaIngreso.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
End Function

'debb2014-d
Function farmMovimientoNotaIngresoSeleccionarXmovimiento(lcMovNumero As String, lcMovTipo As String, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set farmMovimientoNotaIngresoSeleccionarXmovimiento = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoNotaIngresoSeleccionarXmovimiento"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set farmMovimientoNotaIngresoSeleccionarXmovimiento = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-09/07/2015
Function farmMovimientosVentasDetalleDevuelveTodosItems(oConexion As ADODB.Connection, lcMovTipo As String, lcMovNumero As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientosVentasDetalleDevuelveTodosItems"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set farmMovimientosVentasDetalleDevuelveTodosItems = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-27/08/2015
Function DevuelveTipoProducto(lnIdProducto As Long, Optional oConexion1 As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim mo_reglasComunes  As New SIGHNegocios.ReglasComunes
Dim ms_MensajeError As String
Const lcInsumo As String = "Insumo": Const lcMedicamento As String = "Medicamento"
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oRecordset = mo_reglasComunes.CatalogoBienesInsumosSeleccionarXid(lnIdProducto, oConexion)
        DevuelveTipoProducto = IIf(oRecordset!TipoProducto = 1, lcInsumo, lcMedicamento)
        oRecordset.Close
        oConexion.Close
        Set oConexion = Nothing
    Else
        Set oRecordset = mo_reglasComunes.CatalogoBienesInsumosSeleccionarXid(lnIdProducto, oConexion1)
        DevuelveTipoProducto = IIf(oRecordset!TipoProducto = 1, lcInsumo, lcMedicamento)
        oRecordset.Close
    End If
    Set oRecordset = Nothing
    Set mo_reglasComunes = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-03/09/2015
Function RecetaChequeaSiFechaVigenciaEsCorrecta(oRsFarmacia As Recordset) As Boolean
   RecetaChequeaSiFechaVigenciaEsCorrecta = True
   If oRsFarmacia.RecordCount > 0 Then
        oRsFarmacia.MoveFirst
        If oRsFarmacia!FechaVigencia <= CDate(lcBuscaParametro.RetornaFechaServidorSQL) Then
           MsgBox "La FECHA DE VIGENCIA debe ser mayor a HOY", vbInformation, "Receta"
           RecetaChequeaSiFechaVigenciaEsCorrecta = False
        End If
   End If
End Function



'debb-04/05/2016
Function FacturacionBienesFinanciamientosXfechasExoneracion(ldFexoInicio As Date, ldFexoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesFinanciamientosXfechasExoneracion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosXfechasExoneracion"
        Set oParameter = .CreateParameter("@FexoInicio", adDBTimeStamp, adParamInput, 0, ldFexoInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FexoFin", adDBTimeStamp, adParamInput, 0, ldFexoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinanciamientosXfechasExoneracion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-04/05/2016
Function farmMovimientoVentasPacientesSegurosConCredito(ldFmovInicio As Date, ldFmovFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasPacientesSegurosConCredito = Nothing
    ms_MensajeError = ""
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 900
        .CommandText = "farmMovimientoVentasPacientesSegurosConCredito"
        Set oParameter = .CreateParameter("@FmovInicio", adDBTimeStamp, adParamInput, 0, ldFmovInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FmovFin", adDBTimeStamp, adParamInput, 0, ldFmovFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasPacientesSegurosConCredito = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'debb-04/05/2016
Function farmMovimientoVentasPacientesPagantesConCredito(ldFmovInicio As Date, ldFmovFin As Date, _
                                                        ldFcancInicio As Date, ldFcancFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasPacientesPagantesConCredito = Nothing
    ms_MensajeError = ""
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasPacientesPagantesConCredito"
        Set oParameter = .CreateParameter("@FmovInicio", adDBTimeStamp, adParamInput, 0, ldFmovInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FmovFin", adDBTimeStamp, adParamInput, 0, ldFmovFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FcancInicio", adDBTimeStamp, adParamInput, 0, IIf(ldFcancInicio = 0, Null, ldFcancInicio)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FcancFin", adDBTimeStamp, adParamInput, 0, ldFcancFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasPacientesPagantesConCredito = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'debb-08/11/2016
Function CatalogoDIGEMIDAsignaCodigoPaquete() As String
       Dim mo_ReglasFacturacion As New ReglasFacturacion
       Dim oRsTmp1 As New Recordset
       CatalogoDIGEMIDAsignaCodigoPaquete = lcLetrasCodigoPaquete & "001"
       Set oRsTmp1 = mo_ReglasFacturacion.FactCatalogoPaqueteSeleccionarPorFiltro(" tipoPaquete=2  order by codigo desc")
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          If Not IsNull(oRsTmp1!esItemFarmacia) Then
             CatalogoDIGEMIDAsignaCodigoPaquete = lcLetrasCodigoPaquete & Right("0000" & Trim(Str(Val(Mid(oRsTmp1!Codigo, 3, 3)) + 1)), 3)
          End If
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Set mo_ReglasFacturacion = Nothing
End Function
'debb-08/11/2016
Function CatalogoDIGEMIDesCodigoPaquete(lcCodigoPaquete As String, Optional oConexion1 As Connection) As Boolean
    CatalogoDIGEMIDesCodigoPaquete = False
    Dim rs As New Recordset
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigoDEBB1(lcCodigoPaquete, 1, 5, True, oConexion1)
    If rs.RecordCount > 0 Then
       If Not IsNull(rs!esPaquete) Then
       If rs!esPaquete = True Then
          CatalogoDIGEMIDesCodigoPaquete = True
       End If
       End If
    End If
    Set rs = Nothing
    Set mo_ReglasFacturacion = Nothing
End Function

'debb-08/11/2016
Function CatalogoDIGEMIDdevuelveITEMS(lcCodigoPaquete As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set CatalogoDIGEMIDdevuelveITEMS = Nothing
        ms_MensajeError = ""
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionPaquetesFarmaciaXcodigo"
            Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 20, lcCodigoPaquete): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set CatalogoDIGEMIDdevuelveITEMS = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       
        Exit Function
ManejadorDeError:
           ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function


Function LaNIoNSesUnARMADO_PAQUETE(lnIdFarmacia As Long, lcIdTipoConcepto As Long, lcNroDocumento As String, _
                                   lbSeBuscaDesdeNotaIngreso As Boolean) As Boolean
    LaNIoNSesUnARMADO_PAQUETE = False
    If lnIdFarmacia = 999 And lcIdTipoConcepto = 20 Then
        Dim oRsTmp1 As New Recordset
        Dim oConexion As New Connection
        Dim lbExisteItemFueraDelPaquete As Boolean
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        If lbSeBuscaDesdeNotaIngreso = True Then
           Set oRsTmp1 = FarmaciaFiltraTodosMovimientos("and dbo.farmMovimiento.idTipoConcepto=20 and " & _
                                                        "dbo.farmMovimiento.documentoNumero='" & lcNroDocumento & "' and " & _
                                                        "dbo.farmMovimiento.movTipo='S'", oConexion)
        Else
           Set oRsTmp1 = FarmaciaFiltraTodosMovimientos("and dbo.farmMovimiento.idTipoConcepto=20 and " & _
                                                        "dbo.farmMovimiento.documentoNumero='" & lcNroDocumento & "' and " & _
                                                        "dbo.farmMovimiento.movTipo='E'", oConexion)
        End If
        If oRsTmp1.RecordCount > 0 Then
           lbExisteItemFueraDelPaquete = False
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If Not IsNull(oRsTmp1!esPaquete) Then
                If lbSeBuscaDesdeNotaIngreso = True Then
                   If oRsTmp1!esPaquete = True Then
                      lbExisteItemFueraDelPaquete = True
                   End If
                Else
                   If oRsTmp1!esPaquete = False Then
                      lbExisteItemFueraDelPaquete = True
                   End If
                End If
              End If
              oRsTmp1.MoveNext
           Loop
           If lbExisteItemFueraDelPaquete = False Then
              LaNIoNSesUnARMADO_PAQUETE = True
           End If
        End If
        oConexion.Close
        Set oRsTmp1 = Nothing
        Set oConexion = Nothing
    End If
End Function

Function FarmMovimientoVentasDetalleSeleccionarPorNBoleta(lcDocumentoNumero As String, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmMovimientoVentasDetalleSeleccionarPorNBoleta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorNBoleta"
        Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, lcDocumentoNumero): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FarmMovimientoVentasDetalleSeleccionarPorNBoleta = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function




Function Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes(lcCodigoProducto As String, lcAlmacenSismed As String, _
                                                                  ldFechaMov As Date, oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion1
        .CommandTimeout = 150
        .CommandText = "Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes"
        Set oParameter = .CreateParameter("@Producto", adVarChar, adParamInput, 7, lcCodigoProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, lcAlmacenSismed): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaMov", adVarChar, adParamInput, 6, Format(ldFechaMov, "YYYYMM")): .Parameters.Append oParameter
        Set oRecordset = .Execute
   
   End With
   Set Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-10/12/2018
Sub Farm_formDetActualizar(oRsFormDet As Recordset, lcAnioMes As String, lcAlmacen As String, oConexion As Connection, oRsFormDetL As Recordset)
    oRsFormDet.Filter = "annomes='" & lcAnioMes & "' and sit='1'"
    If oRsFormDet.RecordCount > 0 Then
        oRsFormDet.MoveFirst
        Dim oDofarm_formDet As New Dofarm_formDet
        Dim Farm_FormDet As New Farm_FormDet
        Dim oDoFarm_FormDetL As New DoFarm_FormDetL
        Dim oFarm_FormDetL As New Farm_FormDetL
        '
        Set Farm_FormDet.Conexion = oConexion
        oDofarm_formDet.CODIGO_EJE = oRsFormDet!CODIGO_EJE
        oDofarm_formDet.CODIGO_PRE = oRsFormDet!CODIGO_PRE
        oDofarm_formDet.TIPSUM = oRsFormDet!TIPSUM
        oDofarm_formDet.ANNOMES = lcAnioMes
        oDofarm_formDet.IdUsuarioAuditoria = sighentidades.Usuario
        If Farm_FormDet.Eliminar(oDofarm_formDet) Then
           Do While Not oRsFormDet.EOF
If Val(oRsFormDet!CODIGO_MED) = 808 Then
ms_MensajeError = ""
End If
                oDofarm_formDet.CODIGO_MED = oRsFormDet!CODIGO_MED
                oDofarm_formDet.credHosp = oRsFormDet!credHosp
                oDofarm_formDet.defnac = oRsFormDet!defnac
                oDofarm_formDet.dev_merma = oRsFormDet!dev_merma
                oDofarm_formDet.dev_ven = oRsFormDet!dev_ven
                oDofarm_formDet.devol = oRsFormDet!devol
                oDofarm_formDet.distri = oRsFormDet!distri
                oDofarm_formDet.do_con = oRsFormDet!do_con
                oDofarm_formDet.do_fecExp = IIf(IsNull(oRsFormDet!do_fecExp), 0, oRsFormDet!do_fecExp)
                oDofarm_formDet.do_ingre = oRsFormDet!do_ingre
                oDofarm_formDet.do_otr = oRsFormDet!do_otr
                oDofarm_formDet.do_saldo = oRsFormDet!do_saldo
                oDofarm_formDet.do_stk = oRsFormDet!do_stk
                oDofarm_formDet.do_tot = oRsFormDet!do_tot
                oDofarm_formDet.dstkCero = oRsFormDet!dstkCero
                oDofarm_formDet.exo = oRsFormDet!exo
                oDofarm_formDet.fac_perd = oRsFormDet!fac_perd
                oDofarm_formDet.fec_exp = IIf(IsNull(oRsFormDet!fec_exp), 0, oRsFormDet!fec_exp)
                oDofarm_formDet.fecha = IIf(IsNull(oRsFormDet!fecha), 0, oRsFormDet!fecha)
                oDofarm_formDet.indiProc = oRsFormDet!indiProc
                oDofarm_formDet.indiSiga = oRsFormDet!indiSiga
                oDofarm_formDet.ingre = oRsFormDet!ingre
                oDofarm_formDet.intersan = oRsFormDet!intersan
                oDofarm_formDet.merma = oRsFormDet!merma
                oDofarm_formDet.mptoRepo = oRsFormDet!mptoRepo
                oDofarm_formDet.otr_conv = oRsFormDet!otr_conv
                oDofarm_formDet.otras_sal = oRsFormDet!otras_sal
                oDofarm_formDet.Precio = oRsFormDet!Precio
                oDofarm_formDet.reIngre = oRsFormDet!reIngre
                oDofarm_formDet.req = oRsFormDet!total
                oDofarm_formDet.saldo = oRsFormDet!saldo
                oDofarm_formDet.SIS = oRsFormDet!SIS
                oDofarm_formDet.SIT = oRsFormDet!SIT
                oDofarm_formDet.soat = oRsFormDet!soat
                oDofarm_formDet.stock_fin = oRsFormDet!stock_fin
                oDofarm_formDet.stock_fin1 = oRsFormDet!stock_fin1
                oDofarm_formDet.total = oRsFormDet!total
                oDofarm_formDet.transf = oRsFormDet!transf
                oDofarm_formDet.Usuario = oRsFormDet!Usuario
                oDofarm_formDet.vencido = oRsFormDet!vencido
                oDofarm_formDet.venta = oRsFormDet!venta
                oDofarm_formDet.ventaInst = oRsFormDet!ventaInst
                If Farm_FormDet.Insertar(oDofarm_formDet) Then
                End If
                oRsFormDet.MoveNext
           Loop
           '
           Set oFarm_FormDetL.Conexion = oConexion
           oRsFormDetL.Filter = "annomes='" & lcAnioMes & "' and sit='1'"
           oDoFarm_FormDetL.CODIGO_EJE = oRsFormDetL!CODIGO_EJE
           oDoFarm_FormDetL.CODIGO_PRE = oRsFormDetL!CODIGO_PRE
           oDoFarm_FormDetL.TIPSUM = oRsFormDetL!TIPSUM
           oDoFarm_FormDetL.ANNOMES = lcAnioMes
           oDoFarm_FormDetL.IdUsuarioAuditoria = sighentidades.Usuario
           If oFarm_FormDetL.Eliminar(oDoFarm_FormDetL) Then
               
               oRsFormDetL.MoveFirst
               Do While Not oRsFormDetL.EOF
                  oDoFarm_FormDetL.CODIGO_MED = oRsFormDetL!CODIGO_MED
                  oDoFarm_FormDetL.Lote = oRsFormDetL!Lote
                  oDoFarm_FormDetL.FECHVTO = oRsFormDetL!FECHVTO
                  oDoFarm_FormDetL.saldo = oRsFormDetL!saldo
                  oDoFarm_FormDetL.SIT = oRsFormDetL!SIT
                  If oFarm_FormDetL.Insertar(oDoFarm_FormDetL) Then
                  End If
                  oRsFormDetL.MoveNext
               Loop
           End If
           '
        End If
        Set oDofarm_formDet = Nothing
        Set Farm_FormDet = Nothing
        Set oDoFarm_FormDetL = Nothing
        Set oFarm_FormDetL = Nothing
    End If
End Sub




Function farm_formDetXrangoMeses(lcAnioMesInicial As String, lcAnioMesFinal As String, lcAlmacen As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set farm_formDetXrangoMeses = Nothing
        ms_MensajeError = ""
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "farm_formDetXrangoMeses"
            Set oParameter = .CreateParameter("@AnioMes1", adVarChar, adParamInput, 6, lcAnioMesInicial): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@AnioMes2", adVarChar, adParamInput, 6, lcAnioMesFinal): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, lcAlmacen): .Parameters.Append oParameter
            Set oRecordset = .Execute
sighentidades.ParaAuditoria = "paso farm_formdetxrangomeses"
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set farm_formDetXrangoMeses = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
           ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function


Function farm_formDetSeleccionaSoloGrabados(lcAnioMesInicial As String, lcAnioMesFinal As String, _
                                            lcAlmacen As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String, lcMes1 As String
        Set farm_formDetSeleccionaSoloGrabados = Nothing
        ms_MensajeError = ""
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "farm_formDetSeleccionaSoloGrabados"
            Set oParameter = .CreateParameter("@AnioMes1", adVarChar, adParamInput, 6, lcAnioMesInicial): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@AnioMes2", adVarChar, adParamInput, 6, lcAnioMesFinal): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, lcAlmacen): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set farm_formDetSeleccionaSoloGrabados = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
           ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function

Function farm_formDetChequeaICIqueFaltanGrabar(lcAnioMesInicial As String, lcAnioMesFinal As String, _
                                               lcAlmacen As String, ByRef ldDia2 As Date) As Boolean
       On Error GoTo ManejadorDeError
       Dim oRecordset As New ADODB.Recordset
       Dim ms_MensajeError As String, lcMes1 As String
       farm_formDetChequeaICIqueFaltanGrabar = True
       ms_MensajeError = ""
       ms_MensajeError = ""
       Set oRecordset = farm_formDetSeleccionaSoloGrabados(lcAnioMesInicial, lcAnioMesFinal, lcAlmacen)
       If oRecordset.RecordCount > 0 Then
            oRecordset.MoveFirst
            lcMes1 = oRecordset!ANNOMES
            If lcAnioMesInicial <> lcMes1 Then
               ms_MensajeError = ms_MensajeError & "Falta procesar ICI del Año: " & Left(lcMes1, 4) & " y mes: " & Right(lcMes1, 2) & " (tabla farm_FormDet)" & Chr(13)
            End If
            Do While Not oRecordset.EOF
               If lcMes1 <> oRecordset!ANNOMES Then
                  ms_MensajeError = ms_MensajeError & "Falta procesar ICI del  Año: " & Left(lcMes1, 4) & " y mes: " & Right(lcMes1, 2) & " (tabla farm_FormDet)" & Chr(13)
                  lcMes1 = oRecordset!ANNOMES
               End If
               If Right(lcMes1, 2) = "12" Then
                  lcMes1 = Trim(Str(Val(Left(lcMes1, 4)) + 1)) & "01"
               Else
                  lcMes1 = Left(lcMes1, 4) + Right("0" & Trim(Str(Val(Right(lcMes1, 2)) + 1)), 2)
               End If
               'ldDia2 = CDate("01/" & Right(oRecordset!annomes, 2) & "/" & Left(oRecordset!annomes, 4))
               ldDia2 = CDate(sighentidades.UltimaFechaDDMMYYDelMesActual1(CDate("01/" & Right(oRecordset!ANNOMES, 2) & "/" & Left(oRecordset!ANNOMES, 4))))
               oRecordset.MoveNext
            Loop
            If ms_MensajeError <> "" Then
               MsgBox ms_MensajeError, vbInformation, ""
               farm_formDetChequeaICIqueFaltanGrabar = False
            End If
       End If
       Set oRecordset = Nothing
       Exit Function
ManejadorDeError:
           ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function



Function farmSaldoSoloAlmacenSismed(lnIdProducto As Long, Optional oConexion1 As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "farmSaldoSoloAlmacenSismed"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp2.ActiveConnection = Nothing
        End If
   End With
   Set farmSaldoSoloAlmacenSismed = oRsTmp2
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRsTmp2 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function


Function FarmSaldoDetalladoPorIdProductoIdAlmacen(lnIdAlmacen As Long, lnIdProducto As Long) As Recordset
Dim lcSql As String, oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmSaldoDetalladoPorIdProductoIdAlmacen = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmSaldoDetalladoPorIdProductoIdAlmacen"
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmSaldoDetalladoPorIdProductoIdAlmacen = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function

Function SaldoTotalPorAlmacen(lnIdAlmacen As Long) As Recordset
Dim lcSql As String, oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set SaldoTotalPorAlmacen = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SaldoTotalPorAlmacen"
        Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set SaldoTotalPorAlmacen = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function

Function FarmPreventaFiltrar(lcFiltro As String) As Recordset
Dim lcSql As String, oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FarmPreventaFiltrar = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmPreventaFiltrar"
        Set oParameter = .CreateParameter("@filtro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmPreventaFiltrar = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function

Function farmMovimientoDetalleSeleccionarPorMovNumeroTipo(lcMovNumero As String, lcMovTipo As String, oConexion As Connection) As Recordset
Dim lcSql As String, oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set farmMovimientoDetalleSeleccionarPorMovNumeroTipo = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleSeleccionarPorMovNumeroTipo"
        Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@movtipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
   End With
   Set farmMovimientoDetalleSeleccionarPorMovNumeroTipo = oRsTmp1
   Set oCommand = Nothing
End Function


Function farmInventarioParaExportar() As Recordset
Dim lcSql As String, oRsTmp1 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set farmInventarioParaExportar = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmInventarioParaExportar"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set farmInventarioParaExportar = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
End Function


Function farmInventarioDetallePorFiltro(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Dim oParameter As ADODB.Parameter

    Set farmInventarioDetallePorFiltro = Nothing

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmInventarioDetallePorFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set farmInventarioDetallePorFiltro = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function


Function EmpleadosDevuelveId(lcDNI As String) As Long
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_reglasComunes.EmpleadosSeleccionarPorDNI(lcDNI)
    If oRsTmp1.RecordCount > 0 Then
       EmpleadosDevuelveId = oRsTmp1!idEmpleado
    Else
       EmpleadosDevuelveId = 0
    End If
    Set mo_reglasComunes = Nothing
    Set oRsTmp1 = Nothing
End Function

Function farmUnidosisSeleccionarTodos(Optional oConexion1 As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    Set farmUnidosisSeleccionarTodos = Nothing

  With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "farmUnidosisSeleccionarTodos"
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
    Set farmUnidosisSeleccionarTodos = oRecordset
    Set oCommand = Nothing
    If oConexion1 Is Nothing Then
       oConexion.Close
       Set oConexion = Nothing
       Set oRecordset = Nothing
    End If
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Description

End Function


Function FarmaciaEsUnidosis(lnIdFarmacia As Long) As Boolean
    Dim oRsFarmacias As New Recordset
    Set oRsFarmacias = FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & lnIdFarmacia)
    If oRsFarmacias.RecordCount > 0 Then
       If oRsFarmacias!esUnidosis = 1 Then
          FarmaciaEsUnidosis = True
       End If
    End If
    oRsFarmacias.Close
    Set oRsFarmacias = Nothing
End Function


Function DevuelveSiSonItemsDeUNIDOSIS(lbLaFarmaciaEsUnidosis As Boolean, mRs_Productos As Recordset, oRsItemsUnidosis As Recordset) As String
   Dim ms_MensajeError As String
   Dim lcCodigoSinPunto As String
   ms_MensajeError = ""
   If ms_MensajeError = "" And lbLaFarmaciaEsUnidosis = True Then
        If oRsItemsUnidosis.RecordCount = 0 Then
           ms_MensajeError = ms_MensajeError + "No hay ITEMS en la FARMACIA UNIDOSIS" + Chr(13)
        Else
           mRs_Productos.MoveFirst
           Do While Not mRs_Productos.EOF
              lcCodigoSinPunto = Left(mRs_Productos!Codigo, Len(Trim(mRs_Productos!Codigo)) - 1)
              oRsItemsUnidosis.MoveFirst
              oRsItemsUnidosis.Find "codigo='" & lcCodigoSinPunto & "'"
              If oRsItemsUnidosis.EOF Then
                 ms_MensajeError = ms_MensajeError + "El ITEM " + Trim(mRs_Productos.Fields!Codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  no pertenece a FARMACIA UNIDOSIS" + Chr(13)
              End If
              mRs_Productos.MoveNext
           Loop
        End If
   ElseIf ms_MensajeError = "" And lbLaFarmaciaEsUnidosis = False Then
        If oRsItemsUnidosis.RecordCount > 0 Then
           mRs_Productos.MoveFirst
           Do While Not mRs_Productos.EOF
              lcCodigoSinPunto = Left(mRs_Productos!Codigo, Len(Trim(mRs_Productos!Codigo)) - 1)
              oRsItemsUnidosis.MoveFirst
              oRsItemsUnidosis.Find "codigo='" & lcCodigoSinPunto & "'"
              If Not oRsItemsUnidosis.EOF Then
                 ms_MensajeError = ms_MensajeError + "El ITEM " + Trim(mRs_Productos.Fields!Codigo) + " - " + Trim(mRs_Productos.Fields!nombreProducto) + "  solo es usado en FARMACIA UNIDOSIS" + Chr(13)
              End If
              mRs_Productos.MoveNext
           Loop
        End If
   End If
   DevuelveSiSonItemsDeUNIDOSIS = ms_MensajeError
End Function


Function farm_formDetLXmes(lcAnioMes As String, lcCodigoEje As String, lcTipSum As String, lcAlmacen As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set farm_formDetLXmes = Nothing
        ms_MensajeError = ""
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "farm_formdetlSeleccionarPorId"
            Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, lcCodigoEje): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, lcAlmacen): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, lcTipSum): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, lcAnioMes): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, ""): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set farm_formDetLXmes = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
           ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function



Function farmMovimientoVentasSeleccionarXId(lcMovNumero As String, lcMovTipo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasSeleccionarXId = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasSeleccionarPorId"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasSeleccionarXId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farm_psicotropicosSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farm_psicotropicosSeleccionarTodos = Nothing
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farm_psicotropicosSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farm_psicotropicosSeleccionarTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function farm_psicotropTipoSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farm_psicotropTipoSeleccionarTodos = Nothing
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farm_psicotropTipoSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farm_psicotropTipoSeleccionarTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmaciaItemsPorDebajoStockMinimo() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaItemsPorDebajoStockMinimo = Nothing
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaItemsPorDebajoStockMinimo"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaItemsPorDebajoStockMinimo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SECCIÓN RAUL
Function usp_catalogoserviciosagregar(oRsDetalleProductos As ADODB.Recordset) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New FarmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    ms_MensajeError = ""
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
              'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
   
    lnItem = 1
    If oRsDetalleProductos.RecordCount = 0 Then
       bProcesoOK = False: ms_MensajeError = "No existe items": GoTo TerminarNS
    End If
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
      ' oDoMovimientoDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
       oDoMovimientoDetalle.Id = oRsDetalleProductos.Fields!Id
       oDoMovimientoDetalle.Id_Atencion = oRsDetalleProductos.Fields!idAtencion
       oDoMovimientoDetalle.Descripcion_Tipo_Item = oRsDetalleProductos.Fields!Descripcion_Tipo_Item
        oDoMovimientoDetalle.Fg_tipo = oRsDetalleProductos.Fields!Fg_tipo
       oDoMovimientoDetalle.Codigo = oRsDetalleProductos.Fields!Codigo
       oDoMovimientoDetalle.nombreProducto = oRsDetalleProductos.Fields!nombreProducto
       oDoMovimientoDetalle.IdSubClasificacion = oRsDetalleProductos.Fields!IdSubClasificacion
       oDoMovimientoDetalle.labConfHIS = oRsDetalleProductos.Fields!labConfHIS
            
      If Not oMovimientoDetalle.usp_catalogoserviciosagregar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
            lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        usp_catalogoserviciosagregar = True
        oConexion.CommitTrans
    Else
        usp_catalogoserviciosagregar = False
        ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
        ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function


Function usp_listarcatalogoidat(oConexion As ADODB.Connection, lcIdAtencion As String) As ADODB.Recordset
    Dim oTabla As New FarmMovimientoDetalle
    Set oTabla.Conexion = oConexion
    Set usp_listarcatalogoidat = oTabla.usp_listarcatalogoidat(lcIdAtencion)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function usp_eliminacatalogoidat(oConexion As ADODB.Connection, Id As Long) As ADODB.Recordset
    Dim oTabla As New FarmMovimientoDetalle
    Set oTabla.Conexion = oConexion
    Set usp_eliminacatalogoidat = oTabla.usp_eliminacatalogoidat(Id)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function usp_listarcatalogoservicio(lcFiltro As String, lcFiltro1 As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set usp_listarcatalogoservicio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listarcatalogoservicio"
                Set oParameter = .CreateParameter("@Descripcion_Item", adVarChar, adParamInput, 300, lcFiltro): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Descripcion_Tipo_Item", adVarChar, adParamInput, 100, lcFiltro1): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set usp_listarcatalogoservicio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function usp_listarcatalogoservicioc(lcFiltro As String, lcFiltro1 As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set usp_listarcatalogoservicioc = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listarcatalogoservicioc"
                Set oParameter = .CreateParameter("@Codigo_Item", adVarChar, adParamInput, 300, lcFiltro): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Descripcion_Tipo_Item", adVarChar, adParamInput, 100, lcFiltro1): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set usp_listarcatalogoservicioc = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function usp_listarsubclasdx(lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set usp_listarsubclasdx = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listarsubclasdx"
                Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set usp_listarsubclasdx = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function usp_listarhissituacio(lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set usp_listarhissituacio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listarhissituacio"
                Set oParameter = .CreateParameter("@Descripcio", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set usp_listarhissituacio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function usp_listartipoitem(lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set usp_listartipoitem = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listartipoitem"
                Set oParameter = .CreateParameter("@Descripcion_Tipo_Item", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set usp_listartipoitem = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function





