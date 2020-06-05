VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dllFactUCGestionCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para CAJA
'        Programado por: Barrantes D
'        Fecha: Enero 2011
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oRsDetalleProductosConLotes As New Recordset
Dim ms_MensajeError As String

'Carga los Medicamentos que el Paciente tiene que Pagar
Sub CargaDetalleBoletaB(ByRef mrs_FacturacionProductos As ADODB.Recordset, ml_idCuentaAtencion As Long, ByRef lnTotalExonerado As Double)
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim Rs As Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lnMaxItemsBoleta As Integer
    Dim lnContinua As Boolean, lnItems As Integer
    Dim lcSql As String
    
    Dim lnIdOrden As Long: Dim lnIdOrdenNoBoleta As Long
    
    
    lnMaxItemsBoleta = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
    
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    If lcBuscaParametro.SeleccionaFilaParametro(231) = "N" Then
        mo_ReglasComunes.AtencionesActualizaCreditoHospitalario ml_idCuentaAtencion
    End If
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    
    Set Rs = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorCuenta(ml_idCuentaAtencion)
    If Rs.RecordCount > 0 Then
       'calcula hasta que IdOrden va a emitir boleta
       lnItems = 0
       lnTotalExonerado = 0
       Rs.MoveFirst
       Do While Not Rs.EOF
          lnIdOrden = Rs.Fields!IdOrden
          lnTotalExonerado = lnTotalExonerado + Rs!ImporteExonerado
          Do While Not Rs.EOF And lnIdOrden = Rs.Fields!IdOrden
                If Not IsNull(Rs!MovNumero) Then
                    mrs_FacturacionProductos.AddNew
                    mrs_FacturacionProductos!idProducto = Rs!idProducto
                    mrs_FacturacionProductos!Codigo = Rs!Codigo
                    mrs_FacturacionProductos!nombreProducto = Rs!nombreProducto
                    mrs_FacturacionProductos!cantidad = Rs!cantidad
                    mrs_FacturacionProductos!PrecioUnitario = Rs!PrecioUnitario
                    mrs_FacturacionProductos!TotalPorPagar = Rs!TotalPorPagar
                    mrs_FacturacionProductos!IdEstadoFacturacion = Rs!IdEstadoFacturacion
                    mrs_FacturacionProductos!idPuntoCarga = Rs!idPuntoCarga
                    mrs_FacturacionProductos!IdOrden = Rs!IdOrden
                    mrs_FacturacionProductos!EstadoLocal = "L" 'Estado Leido de la BD
                    mrs_FacturacionProductos.Update
                End If
                Rs.MoveNext
                If Rs.EOF Then
                   Exit Do
                End If
          Loop
       Loop
    End If
    
End Sub
'Carga los Servicios que el Paciente tiene que Pagar
Sub CargaDetalleBoletaS(ByRef mrs_FacturacionProductos As ADODB.Recordset, ml_idCuentaAtencion As Long, ByRef lnTotalExonerado As Double)
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasComunes As New ReglasComunes
    Dim Rs As Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros: Dim lnMaxItemsBoleta As Integer
    Dim lnContinua As Boolean, lnItems As Integer
    Dim lcSql As String
    Dim lnIdOrden As Long: Dim lnIdOrdenNoBoleta As Long
    
    
    lnMaxItemsBoleta = Val(lcBuscaParametro.SeleccionaFilaParametro(102))
    
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    If lcBuscaParametro.SeleccionaFilaParametro(231) = "N" Then
        mo_ReglasComunes.AtencionesActualizaCreditoHospitalario ml_idCuentaAtencion
    End If
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    
    Set Rs = mo_ReglasFacturacion.FacturacionServicioPagosPorCuenta(ml_idCuentaAtencion)
    Rs.Filter = "EstadoFacOrdenServicio=1"
    If Rs.RecordCount > 0 Then
       'calcula hasta que IdOrden va a emitir boleta
       lnItems = 0
       Rs.MoveFirst
       lnTotalExonerado = 0
       Do While Not Rs.EOF
          lnIdOrden = Rs.Fields!idOrdenPago
          lnTotalExonerado = lnTotalExonerado + Rs!ImporteExonerado
          Do While Not Rs.EOF And lnIdOrden = Rs.Fields!idOrdenPago
                mrs_FacturacionProductos.AddNew
                mrs_FacturacionProductos!idProducto = Rs!idProducto
                mrs_FacturacionProductos!Codigo = Rs!Codigo
                mrs_FacturacionProductos!nombreProducto = Rs!nombreProducto
                mrs_FacturacionProductos!IdTipoFinanciamiento = Rs!IdTipoFinanciamiento
                mrs_FacturacionProductos!cantidad = Rs!cantidad
                mrs_FacturacionProductos!PrecioUnitario = Rs!PrecioUnitario
                mrs_FacturacionProductos!TotalPorPagar = Rs!TotalPorPagar
                mrs_FacturacionProductos!IdEstadoFacturacion = Rs!IdEstadoFacturacion
                mrs_FacturacionProductos!idPuntoCarga = Rs!idPuntoCarga
                mrs_FacturacionProductos!IdOrden = Rs!idOrdenPago
                mrs_FacturacionProductos!EstadoLocal = "L" 'Estado Leido de la BD
                mrs_FacturacionProductos.Update
                Rs.MoveNext
                If Rs.EOF Then
                   Exit Do
                End If
          Loop
       Loop
    End If
End Sub

Function GenerarRecordsetProductos() As ADODB.Recordset
    Dim mrs_FacturacionProductos As New ADODB.Recordset
    With mrs_FacturacionProductos
          .Fields.Append "IdFacturacionProducto", adInteger
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdTipoFinanciamiento", adInteger
          .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
          .Fields.Append "Poliza", adVarChar, 255
          .Fields.Append "TipoFinanciamiento", adVarChar, 255
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "PrecioUnitario", adCurrency
          .Fields.Append "TotalPorPagar", adCurrency
          .Fields.Append "ImporteSIS", adCurrency, , adFldIsNullable
          .Fields.Append "ImporteSOAT", adCurrency, , adFldIsNullable
          .Fields.Append "ImporteEXO", adCurrency, , adFldIsNullable
          .Fields.Append "IdEstadoFacturacion", adInteger
          .Fields.Append "IdPuntoCarga", adInteger
          .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
          .Fields.Append "IdCajero", adInteger
          .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "FechaCajero", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
          .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
          .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
          .Fields.Append "IdServicioInternamiento", adInteger, , adFldIsNullable
          .Fields.Append "IdUsuarioAuditoria", adInteger, , adFldIsNullable
          .Fields.Append "EstadoLocal", adVarChar, 1
          .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
          .Fields.Append "IdComprobantePagoDevolucion", adInteger, , adFldIsNullable
          .Fields.Append "IdOrden", adInteger
          .Fields.Append "idOrdenPagoNew", adDouble, , adFldIsNullable                        'debb-18/052016
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set GenerarRecordsetProductos = mrs_FacturacionProductos
End Function




Function CajaComprobantePagoCuentaAtencionAgregarB(oDOCajaComprobantePago As DOCajaComprobantesPago, _
                                                    rsFacturacionProductosBien As Recordset, _
                                                    lnIdCuentaAtencionAcerrar As Long, mo_lnIdTablaLISTBARITEMS As Long, _
                                                    mo_lcNombrePc As String) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oFactOrdenesBienes As New FactOrdenesBienes
      Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
      Dim oCajaComprobantePago As New CajaComprobantesPago
      Dim oCajaNroDocumento As New CajaNroDocumento
      Dim oDoCajaNroDocumento As New DOCajaNroDocumento
      Dim oDOCuentaAtencion As New DOCuentaAtencion
      Dim oCuentasAtencion As New CuentasAtencion
      Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
      Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
      Dim bProcesoOK As Boolean
      Dim sResultado As String
      Dim lnIdOrden As Long


10        CajaComprobantePagoCuentaAtencionAgregarB = False
          
20      oConexion.CursorLocation = adUseClient
30      oConexion.CommandTimeout = 300
40      oConexion.Open sighentidades.CadenaConexion
        
50        ms_MensajeError = ""
60        oConexion.BeginTrans
70        Set oFactOrdenesBienes.Conexion = oConexion
80        Set oCajaComprobantePago.Conexion = oConexion
90        Set oCajaNroDocumento.Conexion = oConexion
100       Set oCuentasAtencion.Conexion = oConexion
          
110       oDOCajaComprobantePago.idEstadoComprobante = 4  'PAGADO
120       If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
130           bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
140       End If
          '
150       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
          
          '*********************************************************************************
          '                             PROCESAR BIENES E INSUMOS
          '*********************************************************************************
160       If rsFacturacionProductosBien.RecordCount > 0 Then
170           rsFacturacionProductosBien.MoveFirst
180           Do While Not rsFacturacionProductosBien.EOF
190              lnIdOrden = rsFacturacionProductosBien.Fields!IdOrden
200              Do While Not rsFacturacionProductosBien.EOF And lnIdOrden = rsFacturacionProductosBien.Fields!IdOrden
210                 rsFacturacionProductosBien.MoveNext
220                 If rsFacturacionProductosBien.EOF Then
230                    Exit Do
240                 End If
250              Loop
260              oDoFactOrdenesBienes.IdOrden = lnIdOrden
270              If Not oFactOrdenesBienes.SeleccionarPorId(oDoFactOrdenesBienes) Then
280                 bProcesoOK = False: ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrorManager
290              End If
300              oDoFactOrdenesBienes.IdEstadoFacturacion = 4   'pagado
310              oDoFactOrdenesBienes.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
320              oDoFactOrdenesBienes.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
330              If Not oFactOrdenesBienes.Modificar(oDoFactOrdenesBienes) Then
340                 bProcesoOK = False: ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrorManager
350              End If
360           Loop
          Else
              bProcesoOK = False: ms_MensajeError = "no existe items para " & oDOCajaComprobantePago.NroSerie & "-" & oDOCajaComprobantePago.NroDocumento: GoTo ErrorManager
370       End If
          '********************************************************************************
          '                               COMPROBANTE DE PAGO
          '********************************************************************************
          'ACTUALIZAR NRO DE DOCUMENTO
380       oDoCajaNroDocumento.idCaja = oDOCajaComprobantePago.idCaja
390       oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
          Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOCajaComprobantePago.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
400       oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
410       oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
420       oDoCajaNroDocumento.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
430       If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
440           bProcesoOK = False: ms_MensajeError = oCajaNroDocumento.MensajeError: GoTo ErrorManager
450       End If
          'Cierra Cuenta si se CANCELA BOLETAS DE FARMACIA Y SERVICIOS
460       If lnIdCuentaAtencionAcerrar > 0 Then
470           oDOCuentaAtencion.idCuentaAtencion = lnIdCuentaAtencionAcerrar
480           If Not oCuentasAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
490              bProcesoOK = False: ms_MensajeError = oCuentasAtencion.MensajeError: GoTo ErrorManager
500           End If
510           oDOCuentaAtencion.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
520           If Not mo_ReglasFacturacion.CuentasAtencionPagada(oDOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") = True Then
530              bProcesoOK = False: ms_MensajeError = mo_ReglasFacturacion.MensajeError: GoTo ErrorManager
540           End If
550       End If
          '
560       oConexion.CommitTrans
570       oConexion.Close
580       ms_MensajeError = ""
590       CajaComprobantePagoCuentaAtencionAgregarB = True
ErrorManager:
600       If CajaComprobantePagoCuentaAtencionAgregarB = False Then
610           oConexion.RollbackTrans
620           oConexion.Close
630           ms_MensajeError = ms_MensajeError & _
                             sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                             "Function CajaComprobantePagoCuentaAtencionAgregarB", "dllFactUCGestionCaja.cls")  'debb-02/05/2016
                                   
640       End If
650       Set oConexion = Nothing
660       Set oFactOrdenesBienes = Nothing
670       Set oDoFactOrdenesBienes = Nothing
680       Set oCajaComprobantePago = Nothing
690       Set oCajaNroDocumento = Nothing
700       Set oDoCajaNroDocumento = Nothing
710       Set oDOCuentaAtencion = Nothing
720       Set oCuentasAtencion = Nothing
730       Set mo_ReglasFacturacion = Nothing
740       Set mo_ReglasSeguridad = Nothing
End Function





Function CajaComprobantePagoCuentaAtencionAgregarS(oDOCajaComprobantePago As DOCajaComprobantesPago, _
                                        oDOGestionCaja As doCajaGestion, rsFacturacionProductosServ As Recordset, _
                                        lnIdCuentaAtencionAcerrar As Long, mo_lnIdTablaLISTBARITEMS As Long, _
                                        mo_lcNombrePc As String) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
      Dim oDOSeguros As New DOFacturacionSeguros
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oFacturacionSeguros As New FacturacionSeguros
      Dim oFactOrdenesServicio As New FactOrdenesServicio
      Dim oCajaComprobantePago As New CajaComprobantesPago
      Dim oCajaNroDocumento As New CajaNroDocumento
      Dim oDOCuentaAtencion As New DOCuentaAtencion
      Dim oCuentasAtencion As New CuentasAtencion
      Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
      Dim bProcesoOK As Boolean
      Dim sResultado As String
      Dim sIdOrdenesServProcesadas  As String
      Dim sIdOrdenesServProcesadasMemo As String
      Dim sIdOrdenesBienProcesadas  As String
      Dim sIdOrdenesBienProcesadasMemo As String
10    Dim sSql As String: Dim lnDctosEnCajaComprobantesPago  As Double
      Dim rsActualizar As New ADODB.Recordset
      Dim oFactOrdenServicio As New FactOrdenServicio
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oFactOrdenServicioPagos As New FactOrdenServicioPagos
      Dim oDoFactOrdenServPagos As New DoFactOrdenServPagos
      Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
      Dim lnIdOrdenPago As Long

20        CajaComprobantePagoCuentaAtencionAgregarS = False
          
          oConexion.CommandTimeout = 900
          oConexion.CursorLocation = adUseClient
30        oConexion.Open sighentidades.CadenaConexion
          
50        oConexion.BeginTrans
60        Set oFacturacionSeguros.Conexion = oConexion
70        Set oCajaComprobantePago.Conexion = oConexion
80        Set oCajaNroDocumento.Conexion = oConexion
90        Set oCuentasAtencion.Conexion = oConexion
100       Set oFactOrdenServicioPagos.Conexion = oConexion
110       Set oFactOrdenServicio.Conexion = oConexion
          
120       ms_MensajeError = ""
          
          'AGREGAR EL COMPROBANTE DE PAGO
130       oDOCajaComprobantePago.idEstadoComprobante = 4  'PAGADO
140       If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
150           bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
160       End If
          '
170       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
          
          '*********************************************************************************
          '                                  PROCESAR SERVICIOS
          '*********************************************************************************
          'rsFacturacionProductosServ.RecordCount VERIFICA CUANTOS SERVICIOS VA A PAGAR
180       If rsFacturacionProductosServ.RecordCount > 0 Then ''Buscar donde pasa valor rsFacturacionProductosServ
190           rsFacturacionProductosServ.MoveFirst
200           Do While Not rsFacturacionProductosServ.EOF
210              lnIdOrdenPago = rsFacturacionProductosServ.Fields!IdOrden
220              oDoFactOrdenServPagos.idOrdenPago = lnIdOrdenPago
230              If Not oFactOrdenServicioPagos.SeleccionarPorId(oDoFactOrdenServPagos) Then
240                 bProcesoOK = False: ms_MensajeError = oFactOrdenServicioPagos.MensajeError: GoTo ErrorManager
250              End If
260              oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'pagado
270              oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
280              oDoFactOrdenServPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
290              If Not oFactOrdenServicioPagos.Modificar(oDoFactOrdenServPagos) Then 'Si no se actualiza el pago
300                 bProcesoOK = False: ms_MensajeError = oFactOrdenServicioPagos.MensajeError: GoTo ErrorManager
310              End If
320              oDoFactOrdenServ.IdOrden = oDoFactOrdenServPagos.IdOrden
                    'oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Cargar valores al objero oDoFactOrdenServ
330              If Not oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Then 'Si no cargó valores
340                 bProcesoOK = False: ms_MensajeError = oFactOrdenServicio.MensajeError: GoTo ErrorManager
350              End If
360              oDoFactOrdenServ.IdEstadoFacturacion = 4  'pagado
370              oDoFactOrdenServ.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
380              If Not oFactOrdenServicio.Modificar(oDoFactOrdenServ) Then
390                 bProcesoOK = False: ms_MensajeError = oFactOrdenServicio.MensajeError: GoTo ErrorManager
400              End If
410              Do While Not rsFacturacionProductosServ.EOF And lnIdOrdenPago = rsFacturacionProductosServ.Fields!IdOrden
420                 rsFacturacionProductosServ.MoveNext
430                 If rsFacturacionProductosServ.EOF Then
440                    Exit Do
450                 End If
460              Loop
470           Loop
          Else
              bProcesoOK = False: ms_MensajeError = "no existe items para " & oDOCajaComprobantePago.NroSerie & "-" & oDOCajaComprobantePago.NroDocumento: GoTo ErrorManager
480       End If
          
          
          '********************************************************************************
          '                               COMPROBANTE DE PAGO
          '********************************************************************************
          'ACTUALIZAR NRO DE DOCUMENTO
          Dim oDoCajaNroDocumento As New DOCajaNroDocumento
490       oDoCajaNroDocumento.idCaja = oDOGestionCaja.idCaja
500       oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
          Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOGestionCaja.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
510       oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
520       oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
530       If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
540           bProcesoOK = False: ms_MensajeError = oCajaNroDocumento.MensajeError: GoTo ErrorManager
550       End If
          'Cierra Cuenta si se CANCELA BOLETAS DE FARMACIA Y SERVICIOS
560       If lnIdCuentaAtencionAcerrar > 0 Then
570           oDOCuentaAtencion.idCuentaAtencion = lnIdCuentaAtencionAcerrar
580           If Not oCuentasAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
590              bProcesoOK = False: ms_MensajeError = oCuentasAtencion.MensajeError: GoTo ErrorManager
600           End If
610           oDOCuentaAtencion.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
620           If Not mo_ReglasFacturacion.CuentasAtencionPagada(oDOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") = True Then
630              bProcesoOK = False: ms_MensajeError = mo_ReglasFacturacion.MensajeError: GoTo ErrorManager
640           End If
650       End If

                  
660       oConexion.CommitTrans
670       oConexion.Close
680       ms_MensajeError = ""
690       CajaComprobantePagoCuentaAtencionAgregarS = True
ErrorManager:
700       If CajaComprobantePagoCuentaAtencionAgregarS = False Then
710           oConexion.RollbackTrans
720           oConexion.Close
730           ms_MensajeError = ms_MensajeError & _
                              sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                             "Function CajaComprobantePagoCuentaAtencionAgregarS", "dllFactUCGestionCaja.cls")  'debb-02/05/2016
                 
740       End If
750       Set oConexion = Nothing
760       Set oDOFacturacionServicios = Nothing
770       Set oDoFacturacionBienes = Nothing
780       Set oDOSeguros = Nothing
790       Set oFacturacionServicios = Nothing
800       Set oFacturacionSeguros = Nothing
810       Set oFactOrdenesServicio = Nothing
820       Set oCajaComprobantePago = Nothing
830       Set oCajaNroDocumento = Nothing
840       Set oDOCuentaAtencion = Nothing
850       Set oCuentasAtencion = Nothing
860       Set mo_ReglasFacturacion = Nothing
870       Set mo_ReglasSeguridad = Nothing
End Function




'********************************  CAJA   ************************************
Function CajaComprobantePagoBienesRegistraBoleta(oDOCajaComprobantePago As DOCajaComprobantesPago, _
                                                 oDOGestionCaja As doCajaGestion, oDOFactOrdenBien As DoFactOrdenesBienes, _
                                                 oRsDetalleProductos As Recordset, ml_IdUsuarioAuditoria As Long, _
                                                 mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                                                 lcPaciente As String) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oFactOrdenesBienes As New FactOrdenesBienes
      Dim oCajaComprobantePago As New CajaComprobantesPago
      Dim oCajaNroDocumento As New CajaNroDocumento
      Dim oMovimiento As New farmMovimiento
      Dim oDoMovimiento As New DoFarmMovimiento
      Dim oFarmMovimientoVentas As New farmMovimientoVentas
      Dim oDoFarmMovimientoVentas As New DoFarmMovimientoVentas
      Dim oMovimientoVentasDetalle As New FarmMovimientoVentasDetalle
      Dim oDoMovimientoVentasDetalle As New DoFarmMovimientoVentasDetalle
      Dim oMovimientoDetalle As New FarmMovimientoDetalle
      Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
      Dim oDoFarmPreVenta As New DoFarmPreVenta
      Dim oFarmPreventa As New FarmPreVenta
      Dim oDoRecetaDetalle As New DoRecetaDetalle, oRecetaDetalle As New RecetaDetalle
      Dim oDoRecetaCabecera As New DoRecetaCabecera, oRecetaCabecera As New RecetaCabecera
      Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
      Dim lcCorrelativo As String
      Dim bProcesoOK As Boolean
      Dim sResultado As String
      Dim oDoCajaNroDocumento As New DOCajaNroDocumento
      Dim lcTicket As String
      Dim lcBuscaParametro As New SIGHDatos.Parametros
      Dim lnItem As Long
      Dim ml_IdFuenteFinanciamiento As Long
      Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
      Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
      Dim lnIdReceta As Long
      Dim lcSql As String
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
10        CajaComprobantePagoBienesRegistraBoleta = False
          
          'MODIFICADO POR YAMILL PALOMINO
          
20        oConexion.CursorLocation = adUseClient
30        oConexion.CommandTimeout = 300
40        oConexion.Open sighentidades.CadenaConexion

50        oConexion.BeginTrans
60        Set oFactOrdenesBienes.Conexion = oConexion
70        Set oCajaComprobantePago.Conexion = oConexion
80        Set oMovimientoDetalle.Conexion = oConexion
90        Set oMovimiento.Conexion = oConexion
100       Set oFarmPreventa.Conexion = oConexion
110       Set oFarmMovimientoVentas.Conexion = oConexion
120       Set oMovimientoVentasDetalle.Conexion = oConexion
130       Set oCajaNroDocumento.Conexion = oConexion
140       Set oRecetaCabecera.Conexion = oConexion
150       Set oRecetaDetalleItem.Conexion = oConexion
          
          '*********comprobante pago
          
160       oDOCajaComprobantePago.IdTipoOrden = 2  'Bienes e Insumos
170       oDOCajaComprobantePago.idEstadoComprobante = 4    'pagado
180       If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
190           bProcesoOK = False: GoTo ErrorManager1
200       End If
          '
210       oDoFarmPreVenta.idPreVenta = oDOFactOrdenBien.idPreVenta
220       If Not oFarmPreventa.SeleccionarPorId(oDoFarmPreVenta) Then
230           bProcesoOK = False: GoTo ErrorManager1
240       End If
          '
          
250       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)        'ListBarItems.idListItem
          '*********  graba tabla correlativos farmTipoDocumentos  ***************
260       lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
          '*********  graba tabla Movimientos  ***************
270       lcTicket = Trim(oDOCajaComprobantePago.NroSerie) & "-" & Trim(oDOCajaComprobantePago.NroDocumento)
280       With oDoMovimiento
290          .FechaCreacion = oDOCajaComprobantePago.FechaCobranza
300          .idAlmacenDestino = 0  'Ninguno
310          .idAlmacenOrigen = oDOCajaComprobantePago.idFarmacia
320          .idEstadoMovimiento = 1
330          .idTipoConcepto = 10
340          .idUsuario = oDOCajaComprobantePago.IdUsuarioAuditoria
350          .MovTipo = "S"
360          .MovNumero = lcCorrelativo
370          .DocumentoNumero = lcTicket
380          .DocumentoIdtipo = 26
390          .total = oDOCajaComprobantePago.total
400          .idUsuario = oDoFarmPreVenta.idVendedor
410          .IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
420       End With
          
430       If Not oMovimiento.Insertar(oDoMovimiento) Then
440           bProcesoOK = False: GoTo ErrorManager1
450       End If
          '*************vuelve a chequear saldos ****************
460       If Not DevuelveDetalleProductosConLotes(oConexion, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductos, oDoMovimiento.FechaCreacion, sghTipoSalidaItemFarmacia.sghSoloVenta) Then
470           ms_MensajeError = ms_MensajeError + "Hubo problemas al Devolver Saldos con Lotes/FechasVencimiento "
480            bProcesoOK = False: GoTo ErrorManager1
490       End If
500       oRsDetalleProductosConLotes.MoveFirst
510       Do While Not oRsDetalleProductosConLotes.EOF
520          If Not ChequeaQueSaldosConLotesSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
                                          oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, _
                                          oRsDetalleProductosConLotes.Fields!FechaVencimiento, _
                                          oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                          oRsDetalleProductosConLotes.Fields!cantidad) Then
530                ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + _
                                     Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + _
                                     Trim(oRsDetalleProductosConLotes.Fields!nombreProducto) + Chr(13) + _
                                     "(Farmacia debe modificar PREVENTA y controlar mejor su STOCK MINIMO)" + Chr(13)
540                bProcesoOK = False: GoTo ErrorManager1
550          ElseIf ChequeaQueSaldosSeaPositivo(oConexion, oDoMovimiento.idAlmacenOrigen, _
                                                oRsDetalleProductosConLotes.Fields!idProducto, _
                                                oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, _
                                                oRsDetalleProductosConLotes.Fields!cantidad) = False Then
560                ms_MensajeError = ms_MensajeError + "Ya no hay Saldos para el Producto: " + _
                                     Trim(oRsDetalleProductosConLotes.Fields!Codigo) + " " + _
                                     Trim(oRsDetalleProductosConLotes.Fields!nombreProducto) + Chr(13) + _
                                     "(Farmacia debe modificar PREVENTA y controlar mejor su STOCK MINIMO)" + Chr(13)
570                bProcesoOK = False: GoTo ErrorManager1
580          End If
590          oRsDetalleProductosConLotes.MoveNext
600       Loop
          '********* preventas
610       oDoFarmPreVenta.idEstadoPreventa = 2   'con documento emitido
620       oDoFarmPreVenta.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
630       If Not oFarmPreventa.Modificar(oDoFarmPreVenta) Then
640           bProcesoOK = False: GoTo ErrorManager1
650       End If
660       ml_IdFuenteFinanciamiento = 0    'Particular= tenga o  no "Nro CUENTA", Historia
          '*********  graba tabla MovimientosVentas  ***************
670       With oDoFarmMovimientoVentas
680          .MovTipo = oDoMovimiento.MovTipo
690          .MovNumero = oDoMovimiento.MovNumero
700          .idCuentaAtencion = oDoFarmPreVenta.idCuentaAtencion
710          .IdDiagnostico = oDoFarmPreVenta.IdDiagnostico
720          .idPaciente = oDoFarmPreVenta.idPaciente
730          .idPrescriptor = oDoFarmPreVenta.idPrescriptor
740          .idPreVenta = oDoFarmPreVenta.idPreVenta
750          .IdTipoFinanciamiento = oDoFarmPreVenta.IdTipoFinanciamiento
760          .idTipoReceta = oDoFarmPreVenta.idTipoReceta
770          .IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
780          .tipoVenta = "P"
790          .idFuenteFinanciamiento = ml_IdFuenteFinanciamiento
800          If .idCuentaAtencion > 0 Then
810             .IdServicioPaciente = DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(.idCuentaAtencion, _
                            Format(oDoMovimiento.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), _
                            Format(oDoMovimiento.FechaCreacion, "hh:mm"), oConexion)
820          Else
830             .IdServicioPaciente = 0
840          End If
850          .FechaHoraPrescribe = oDoFarmPreVenta.FechaHoraPrescribe
860       End With
870       If Not oFarmMovimientoVentas.Insertar(oDoFarmMovimientoVentas) Then
880                bProcesoOK = False: GoTo ErrorManager1
890       End If
          '************ FactOrdenBien
          
900       oDOFactOrdenBien.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
910       oDOFactOrdenBien.IdEstadoFacturacion = 4  'pagada
920       oDOFactOrdenBien.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
930       oDOFactOrdenBien.MovNumero = oDoMovimiento.MovNumero
940       oDOFactOrdenBien.MovTipo = oDoMovimiento.MovTipo
950       If Not oFactOrdenesBienes.Modificar(oDOFactOrdenBien) Then
960           bProcesoOK = False: GoTo ErrorManager1
970       End If
          '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
          
980       oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
990       oDoMovimientoDetalle.MovNumero = oDoMovimiento.MovNumero
1000      oDoMovimientoDetalle.MovTipo = oDoMovimiento.MovTipo
1010      lnItem = 1
1020      oRsDetalleProductosConLotes.MoveFirst
1030      Do While Not oRsDetalleProductosConLotes.EOF
1040         oDoMovimientoDetalle.cantidad = oRsDetalleProductosConLotes.Fields!cantidad
1050         oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductosConLotes.Fields!FechaVencimiento
1060         oDoMovimientoDetalle.idProducto = oRsDetalleProductosConLotes.Fields!idProducto
1070         oDoMovimientoDetalle.Item = lnItem
1080         oDoMovimientoDetalle.Lote = oRsDetalleProductosConLotes.Fields!Lote
1090         oDoMovimientoDetalle.Precio = oRsDetalleProductosConLotes.Fields!Precio
1100         oDoMovimientoDetalle.RegistroSanitario = ""
1110         oDoMovimientoDetalle.total = oRsDetalleProductosConLotes.Fields!total
1120         oDoMovimientoDetalle.idTipoSalidaBienInsumo = oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo
1130         If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
1140             bProcesoOK = False: GoTo ErrorManager1
1150         End If
1160         If Not oMovimientoDetalle.ActualizaSaldosPorProducto(oDoMovimiento.MovTipo, oDoMovimiento.idAlmacenOrigen, oRsDetalleProductosConLotes.Fields!idProducto, oRsDetalleProductosConLotes.Fields!Lote, oRsDetalleProductosConLotes.Fields!FechaVencimiento, oRsDetalleProductosConLotes.Fields!idTipoSalidaBienInsumo, oRsDetalleProductosConLotes.Fields!cantidad, oRsDetalleProductosConLotes.Fields!Precio, oDoMovimiento.FechaCreacion) Then
1170             bProcesoOK = False: GoTo ErrorManager1
1180         End If
1190         lnItem = lnItem + 1
1200         oRsDetalleProductosConLotes.MoveNext
1210      Loop
          '*********  graba tabla farmMovimientosVentasDetalle, FacturacionBienesFinanciamientos ***************
          
1220      oDoMovimientoVentasDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
1230      oDoMovimientoVentasDetalle.MovNumero = oDoMovimiento.MovNumero
1240      oDoMovimientoVentasDetalle.MovTipo = oDoMovimiento.MovTipo
1250      oRsDetalleProductos.MoveFirst
1260      Do While Not oRsDetalleProductos.EOF
1270         oDoMovimientoVentasDetalle.cantidad = oRsDetalleProductos.Fields!cantidad
1280         oDoMovimientoVentasDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
1290         oDoMovimientoVentasDetalle.Precio = oRsDetalleProductos.Fields!Precio
1300         oDoMovimientoVentasDetalle.total = oRsDetalleProductos.Fields!total
1310         If Not oMovimientoVentasDetalle.Insertar(oDoMovimientoVentasDetalle) Then
1320             bProcesoOK = False: GoTo ErrorManager1
1330         End If
1340         oRsDetalleProductos.MoveNext
1350      Loop
          '************* Actualiza Nro Serie y Documento
          
1360      oDoCajaNroDocumento.idCaja = oDOGestionCaja.idCaja
1370      oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
          Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOGestionCaja.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
1380        oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
1390        oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
1400        If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
1410            bProcesoOK = False: GoTo ErrorManager1
1420        End If
       
          'Receta automatica
          'debb-18/05/2016
1430      If oDoFarmPreVenta.idPreVenta > 0 Then
1440          If oRecetaCabecera.ConsultaRecetaCabeceraFarmacia(oDoRecetaCabecera, Trim(Str(oDoFarmPreVenta.idPreVenta))) = True Then
1450               lnIdReceta = oDoRecetaCabecera.IdReceta
1460               oDoRecetaCabecera.idEstado = 3
1470               oDoRecetaCabecera.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
1490               If oRecetaCabecera.Modificar(oDoRecetaCabecera) = False Then
1500                  bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo ErrorManager1
1510               End If
1520               Set oRecetaDetalle.Conexion = oConexion
1530               Set oRsTmp1 = oRecetaDetalle.SeleccionarPorIdReceta(lnIdReceta)
                   
'                   'debb-18/05/2016 (inicio)
'1540               oDoRecetaDetalleItem.IdReceta = lnIdReceta
'1550               oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoFarmPreVenta.IdUsuarioAuditoria
'1560               If oRecetaDetalleItem.Eliminar(oDoRecetaDetalleItem) = False Then
'1570                  bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo ErrorManager1
'1580               End If
'                   'debb-18/05/2016 (fin)
                   
1590               oRsDetalleProductos.MoveFirst
1600               Do While Not oRsDetalleProductos.EOF
1610                 oRsTmp1.MoveFirst
1620                 oRsTmp1.Find "idItem=" & oRsDetalleProductos.Fields!idProducto
1630                 If Not oRsTmp1.EOF Then
1640                       oDoRecetaDetalle.IdReceta = lnIdReceta
1650                       oDoRecetaDetalle.idItem = oRsDetalleProductos.Fields!idProducto
1660                       oDoRecetaDetalle.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
1670                       oDoRecetaDetalle.idEstadoDetalle = 3
1680                       If oRecetaDetalle.RecetaDetalleActualizaCantDespachada(oDoRecetaDetalle) = False Then
1690                           bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo ErrorManager1
1700                       End If
1710                       oDoRecetaDetalleItem.IdReceta = lnIdReceta
1720                       oDoRecetaDetalleItem.CantidadDespachada = oRsDetalleProductos.Fields!cantidad
1730                       oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDoFarmPreVenta.idPreVenta))
1740                       oDoRecetaDetalleItem.idItem = oRsDetalleProductos.Fields!idProducto
                           oDoRecetaDetalleItem.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
1750                       oDoRecetaDetalleItem.IdUsuarioAuditoria = oDoFarmPreVenta.IdUsuarioAuditoria
                           oDoRecetaDetalleItem.idEstado = sghRecetaEstados.sighRecetaConBoleta
1760                       If oRecetaDetalleItem.Modificar(oDoRecetaDetalleItem) = False Then
1770                          bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo ErrorManager1
1780                       End If
1790                 End If
1800                 oRsDetalleProductos.MoveNext
1810               Loop
1820               oRsTmp1.Close
1830          End If
1840      End If
          '
1850      oConexion.CommitTrans
1860      oConexion.Close
1870      Set oConexion = Nothing
1880      Set oCommand = Nothing
1890      Set mo_ReglasSeguridad = Nothing
1900      ms_MensajeError = ""
1910      CajaComprobantePagoBienesRegistraBoleta = True
1920      Exit Function
ErrorManager1:
1930      oConexion.RollbackTrans
1940      oConexion.Close
1950      Set oConexion = Nothing
1960      ms_MensajeError = ms_MensajeError + IIf(Err.Description <> "", Err.Description + Chr(13), "")
1970      ms_MensajeError = ms_MensajeError + IIf(oFactOrdenesBienes.MensajeError <> "", oFactOrdenesBienes.MensajeError + Chr(13), "")
1980      ms_MensajeError = ms_MensajeError + IIf(oCajaComprobantePago.MensajeError <> "", oCajaComprobantePago.MensajeError + Chr(13), "")
1990      ms_MensajeError = ms_MensajeError + IIf(oCajaNroDocumento.MensajeError <> "", oCajaNroDocumento.MensajeError + Chr(13), "")
2000      ms_MensajeError = ms_MensajeError + IIf(oMovimiento.MensajeError <> "", oMovimiento.MensajeError + Chr(13), "")
2010      ms_MensajeError = ms_MensajeError + IIf(oFarmMovimientoVentas.MensajeError <> "", oFarmMovimientoVentas.MensajeError + Chr(13), "")
2020      ms_MensajeError = ms_MensajeError + IIf(oMovimientoVentasDetalle.MensajeError <> "", oMovimientoVentasDetalle.MensajeError + Chr(13), "")
2030      ms_MensajeError = ms_MensajeError + IIf(oMovimientoDetalle.MensajeError <> "", oMovimientoDetalle.MensajeError + Chr(13), "")
2040      ms_MensajeError = ms_MensajeError + IIf(oFarmPreventa.MensajeError <> "", oFarmPreventa.MensajeError + Chr(13), "")
2050      ms_MensajeError = ms_MensajeError & _
                            sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                            "Function CajaComprobantePagoBienesRegistraBoleta", "dllFactUCGestionCaja.cls")  'debb-02/05/2016
                                
2060      CajaComprobantePagoBienesRegistraBoleta = False
2070      Set oFactOrdenesBienes = Nothing
2080      Set oCajaComprobantePago = Nothing
2090      Set oCajaNroDocumento = Nothing
2100      Set oMovimiento = Nothing
2110      Set oFarmMovimientoVentas = Nothing
2120      Set oDoFarmMovimientoVentas = Nothing
2130      Set oMovimientoVentasDetalle = Nothing
2140      Set oDoMovimientoDetalle = Nothing
2150      Set oDoFarmPreVenta = Nothing
2160      Set oFarmPreventa = Nothing
2170      Set oDoCajaNroDocumento = Nothing
2180      Set lcBuscaParametro = Nothing
2190      Set oRsTmp1 = Nothing
2200      Set oDoRecetaDetalle = Nothing
2210      Set oRecetaDetalle = Nothing
2220      Set oDoRecetaCabecera = Nothing
2230      Set oRecetaCabecera = Nothing
2240      Set oRecetaDetalleItem = Nothing
2250      Set oDoRecetaDetalleItem = Nothing
2260      Set mo_ReglasSeguridad = Nothing
End Function
Function DevuelveDetalleProductosConLotes(oConexion As ADODB.Connection, lnIdAlmacen As Long, oRsConsolidado As Recordset, ldFechaActual As Date, lnIdTipoSalidaBienInsumo As Long) As Boolean
    On Error Resume Next
    Dim lcBuscaParametro As New SIGHDatos.Parametros
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
             Set oRsSaldosEnEsteMomento = oFarmMovimientoDetalle.FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen, 0, Trim(oRsConsolidado.Fields!Codigo), oConexion)
             oRsSaldosEnEsteMomento.Filter = "idTipoSalidaBienInsumoSaldo=" & lnIdTipoSalidaBienInsumo
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
    Set lcBuscaParametro = Nothing
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
        Set oRecordset = AtencionesSelecionarPorCuenta(lidCuenta, oConexion)
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
Function AtencionesSelecionarPorCuenta(lnCuenta As Long, Optional oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set AtencionesSelecionarPorCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSelecionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesSelecionarPorCuenta = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function







Function CajaComprobantePagoServicioAgregar(oDOCajaComprobantePago As DOCajaComprobantesPago, _
                                 oDOGestionCaja As doCajaGestion, oDoFactOrdenServPagos As DoFactOrdenServPagos, _
                                 rsFacturacionProductos As Recordset, ml_IdUsuarioAuditoria As Long, _
                                 oDoAtencion As DOAtencion, lnIdPuntoCarga As Long, mo_lnIdTablaLISTBARITEMS As Long, _
                                 mo_lcNombrePc As String, lnIdReceta As Long, lbEsUnaRecetaOtrosCpt As Boolean) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
      Dim oDOSeguros As New DOFacturacionSeguros
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oFacturacionBienes As New FacturacionBienesInsumos
      Dim oFacturacionSeguros As New FacturacionSeguros
      Dim oFactOrdenesServicio As New FactOrdenesServicio
      Dim oCajaComprobantePago As New CajaComprobantesPago
      Dim oCajaNroDocumento As New CajaNroDocumento
      Dim oFactOrdenServPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oDoFacturacionServicioDespacho As New DoFacturacionServicioDespacho
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oFactOrdenServ As New FactOrdenServicio
      Dim oCitas As New Citas
      Dim bProcesoOK As Boolean
      Dim sResultado As String
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
      Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad

10        CajaComprobantePagoServicioAgregar = False
          
20        oConexion.CursorLocation = adUseClient
30        oConexion.CommandTimeout = 300
40        oConexion.Open sighentidades.CadenaConexion
            
50        oConexion.BeginTrans
60        Set oFacturacionServicios.Conexion = oConexion
70        Set oFacturacionBienes.Conexion = oConexion
80        Set oFacturacionSeguros.Conexion = oConexion
90        Set oFactOrdenesServicio.Conexion = oConexion
100       Set oCajaComprobantePago.Conexion = oConexion
110       Set oCajaNroDocumento.Conexion = oConexion
120       Set oCitas.Conexion = oConexion
130       Set oFactOrdenServPagos.Conexion = oConexion
140       Set oFacturacionServicioPagos.Conexion = oConexion
150       Set oFacturacionServicioDespacho.Conexion = oConexion
160       Set oFactOrdenServ.Conexion = oConexion
          
170       oDOCajaComprobantePago.idEstadoComprobante = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
180       oDOCajaComprobantePago.IdTipoOrden = 1  'Servicio
          If lbEsUnaRecetaOtrosCpt = True Then
             oDOCajaComprobantePago.idCuentaAtencion = 0
          End If
190       If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
200           bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
210       End If
          '
220       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
          
          'Crear orden de servicio
230       oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'Pagado
240       oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
250       If oDoFactOrdenServPagos.idOrdenPago = 0 Then
              'Se ha creado desde CAJA
260           oDoFactOrdenServ.FechaCreacion = oDoFactOrdenServPagos.FechaCreacion
270           oDoFactOrdenServ.fechaDespacho = oDoFactOrdenServPagos.FechaCreacion
              If lbEsUnaRecetaOtrosCpt = True Then
                 oDoFactOrdenServ.idCuentaAtencion = 0
              Else
                 oDoFactOrdenServ.idCuentaAtencion = oDOCajaComprobantePago.idCuentaAtencion
              End If
290           oDoFactOrdenServ.IdEstadoFacturacion = 4 'PAGADO
300           oDoFactOrdenServ.idFuenteFinanciamiento = 1 'Contado
310           oDoFactOrdenServ.idPaciente = oDOCajaComprobantePago.idPaciente
320           oDoFactOrdenServ.idPuntoCarga = lnIdPuntoCarga
330           oDoFactOrdenServ.IdTipoFinanciamiento = oDOCajaComprobantePago.IdTipoFinanciamiento
              
340           oDoFactOrdenServ.idUsuario = oDoFactOrdenServPagos.IdUsuarioAuditoria
350           oDoFactOrdenServ.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
360           oDoFactOrdenServ.IdUsuarioDespacho = oDoFactOrdenServPagos.IdUsuarioAuditoria
370           If oDoFactOrdenServ.idCuentaAtencion > 0 Then
380              oDoFactOrdenServ.IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFactOrdenServ.idCuentaAtencion, Format(oDoFactOrdenServ.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), Format(Now, "hh:mm"))
390           Else
400              oDoFactOrdenServ.IdServicioPaciente = 0
410           End If
420           If Not oFactOrdenServ.Insertar(oDoFactOrdenServ) Then
430               bProcesoOK = False: GoTo ErrorManager
440           End If
450           oDoFactOrdenServPagos.IdOrden = oDoFactOrdenServ.IdOrden
460           If Not oFactOrdenServPagos.Insertar(oDoFactOrdenServPagos) Then
470               bProcesoOK = False: GoTo ErrorManager
480           End If
490           rsFacturacionProductos.MoveFirst
500           Do While Not rsFacturacionProductos.EOF
510              oDoFacturacionServicioDespacho.cantidad = rsFacturacionProductos!cantidad
520              oDoFacturacionServicioDespacho.IdOrden = oDoFactOrdenServ.IdOrden
530              oDoFacturacionServicioDespacho.idProducto = rsFacturacionProductos!idProducto
540              oDoFacturacionServicioDespacho.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
550              oDoFacturacionServicioDespacho.Precio = rsFacturacionProductos!PrecioUnitario
560              oDoFacturacionServicioDespacho.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
570              If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDespacho) Then
580                   bProcesoOK = False: GoTo ErrorManager
590              End If
600              oDoFacturacionServicioPagos.cantidad = rsFacturacionProductos!cantidad
610              oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
620              oDoFacturacionServicioPagos.idProducto = rsFacturacionProductos!idProducto
630              oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
640              oDoFacturacionServicioPagos.Precio = rsFacturacionProductos!PrecioUnitario
650              oDoFacturacionServicioPagos.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
660              If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
670                   bProcesoOK = False: GoTo ErrorManager
680              End If
690              rsFacturacionProductos.MoveNext
700           Loop
710       Else
              'Se ha creado en FACTURACION, Es una PREVENTA (SERVICIOS)
720           If Not oFactOrdenServPagos.Modificar(oDoFactOrdenServPagos) Then
730               bProcesoOK = False: GoTo ErrorManager
740           End If
750           oDoFactOrdenServ.IdOrden = oDoFactOrdenServPagos.IdOrden
760           If Not oFactOrdenServ.SeleccionarPorId(oDoFactOrdenServ) Then
770               bProcesoOK = False: GoTo ErrorManager
780           End If
790           oDoFactOrdenServ.IdEstadoFacturacion = 4 'Pagado
800           If Not oFactOrdenServ.Modificar(oDoFactOrdenServ) Then
810               bProcesoOK = False: GoTo ErrorManager
820           End If
830       End If
                  
          'Si la orden es de Admsión de Consulta Externa
840       If lnIdPuntoCarga = 6 Then
              'Generar una solicitud automatica, actualizar el estado de la cita a pagado
850           If Not oCitas.ActualizarCitaPagadaDEBB(oDoAtencion.idAtencion, ml_IdUsuarioAuditoria) Then
860               bProcesoOK = False: GoTo ErrorManager
870           End If
880       End If
                  
          Dim lbContinnn As Boolean
          Dim lcBuscaParametro As New SIGHDatos.Parametros
          Dim oDoCajaNroDocumento As New DOCajaNroDocumento
          lbContinnn = True
          If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghReembolsos Then
             If lcBuscaParametro.SeleccionaFilaParametro(584) <> "S" Then
                lbContinnn = False
             End If
          End If
          Set lcBuscaParametro = Nothing
          If lbContinnn = True Then
890       oDoCajaNroDocumento.idCaja = oDOCajaComprobantePago.idCaja
900       oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
          Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOCajaComprobantePago.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
910       oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
920       oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
930       If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
940           bProcesoOK = False: GoTo ErrorManager
950       End If
          End If
          'Receta automatica debb-03/04/2012
960       If lnIdReceta > 0 Then
970           With oCommand
980             .CommandType = adCmdStoredProc
990             Set .ActiveConnection = oConexion
1000            .CommandTimeout = 150
1010            .CommandText = "ActualizaRecetaCabecerarXEstadoYComprobantePago"
1020            Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oDOCajaComprobantePago.IdComprobantePago)
1030            .Parameters.Append oParameter
1040            Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta)
1050            .Parameters.Append oParameter
1060            .Execute
1070          End With
1080          Set oCommand = Nothing
              'debb-18/05/2016 (inicio)
              Dim oDoRecetaDetalleItem As New DoRecetaDetalleItem, oRecetaDetalleItem As New RecetaDetalleItem
1090          Set oRecetaDetalleItem.Conexion = oConexion
              oDoRecetaDetalleItem.IdReceta = lnIdReceta
              oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
1100          rsFacturacionProductos.MoveFirst
1110          Do While Not rsFacturacionProductos.EOF
1120             oDoRecetaDetalleItem.CantidadDespachada = 1
1130             oDoRecetaDetalleItem.DocumentoDespacho = " "
1140             oDoRecetaDetalleItem.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
                 oDoRecetaDetalleItem.idItem = rsFacturacionProductos!idProducto
                 oDoRecetaDetalleItem.idEstado = sghRecetaEstados.sighRecetaConBoleta
1170             If Not oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) Then
1180                bProcesoOK = False: GoTo ErrorManager
1190             End If
1200             rsFacturacionProductos.MoveNext
1210          Loop
1220          Set oDoRecetaDetalleItem = Nothing
1230          Set oRecetaDetalleItem = Nothing
              'debb-18/05/2016 (fin)
1240      End If    '
1250      oConexion.CommitTrans
1260      oConexion.Close
1270      Set oConexion = Nothing
1280      ms_MensajeError = ""
1290      CajaComprobantePagoServicioAgregar = True
          
ErrorManager:

1300      If CajaComprobantePagoServicioAgregar = False Then

1310          ms_MensajeError = Err.Description
1320          oConexion.RollbackTrans
1330          oConexion.Close
1340          Set oConexion = Nothing
1350          ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                             "Function CajaComprobantePagoServicioAgregar", "dllFactUCGestionCaja.cls")  'debb-02/05/2016
                  
1360      End If
1370      Set oDOFacturacionServicios = Nothing
1380      Set oDoFacturacionBienes = Nothing
1390      Set oDOSeguros = Nothing
1400      Set oFacturacionServicios = Nothing
1410      Set oFacturacionBienes = Nothing
1420      Set oFacturacionSeguros = Nothing
1430      Set oFactOrdenesServicio = Nothing
1440      Set oCajaComprobantePago = Nothing
1450      Set oCajaNroDocumento = Nothing
1460      Set oFactOrdenServPagos = Nothing
1470      Set oDoFacturacionServicioPagos = Nothing
1480      Set oFacturacionServicioPagos = Nothing
1490      Set oFacturacionServicioDespacho = Nothing
1500      Set oDoFacturacionServicioDespacho = Nothing
1510      Set oDoFactOrdenServ = Nothing
1520      Set oFactOrdenServ = Nothing
1530      Set oCitas = Nothing
1540      Set mo_ReglasSeguridad = Nothing
End Function

'debb-21/08/2015 (inicio)
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'debb-21/08/2015 (fin)


Function CajaComprobantePagoServicioAgregarHISTORICO(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, oDoFactOrdenServPagos As DoFactOrdenServPagos, rsFacturacionProductos As Recordset, ml_IdUsuarioAuditoria As Long, oDoAtencion As DOAtencion, lnIdPuntoCarga As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdReceta As Long) As Boolean
      Dim oConexion As New ADODB.Connection
      Dim oDOFacturacionServicios As New DOFacturacionServicios
      Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
      Dim oDOSeguros As New DOFacturacionSeguros
      Dim oFacturacionServicios As New FacturacionServicios
      Dim oFacturacionBienes As New FacturacionBienesInsumos
      Dim oFacturacionSeguros As New FacturacionSeguros
      Dim oFactOrdenesServicio As New FactOrdenesServicio
      Dim oCajaComprobantePago As New CajaComprobantesPago
      Dim oCajaNroDocumento As New CajaNroDocumento
      Dim oFactOrdenServPagos As New FactOrdenServicioPagos
      Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
      Dim oFacturacionServicioPagos As New FacturacionServicioPagos
      Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
      Dim oDoFacturacionServicioDespacho As New DoFacturacionServicioDespacho
      Dim oDoFactOrdenServ As New DoFactOrdenServ
      Dim oFactOrdenServ As New FactOrdenServicio
      Dim oCitas As New Citas
      Dim bProcesoOK As Boolean
      Dim sResultado As String
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
      Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad

10        CajaComprobantePagoServicioAgregarHISTORICO = False
          
20        oConexion.CursorLocation = adUseClient
30        oConexion.CommandTimeout = 300
40        oConexion.Open sighentidades.CadenaConexion
            
50        oConexion.BeginTrans
60        Set oFacturacionServicios.Conexion = oConexion
70        Set oFacturacionBienes.Conexion = oConexion
80        Set oFacturacionSeguros.Conexion = oConexion
90        Set oFactOrdenesServicio.Conexion = oConexion
100       Set oCajaComprobantePago.Conexion = oConexion
110       Set oCajaNroDocumento.Conexion = oConexion
120       Set oCitas.Conexion = oConexion
130       Set oFactOrdenServPagos.Conexion = oConexion
140       Set oFacturacionServicioPagos.Conexion = oConexion
150       Set oFacturacionServicioDespacho.Conexion = oConexion
160       Set oFactOrdenServ.Conexion = oConexion
          
170       oDOCajaComprobantePago.idEstadoComprobante = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
180       oDOCajaComprobantePago.IdTipoOrden = 1  'Servicio
190       If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
200           bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
210       End If
          '
220       Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
          
          'Crear orden de servicio
230       oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'Pagado
240       oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
250       If oDoFactOrdenServPagos.idOrdenPago = 0 Then
              'Se ha creado desde CAJA
260           oDoFactOrdenServ.FechaCreacion = oDoFactOrdenServPagos.FechaCreacion
270           oDoFactOrdenServ.fechaDespacho = oDoFactOrdenServPagos.FechaCreacion
280           oDoFactOrdenServ.idCuentaAtencion = oDOCajaComprobantePago.idCuentaAtencion
290           oDoFactOrdenServ.IdEstadoFacturacion = 4 'PAGADO
300           oDoFactOrdenServ.idFuenteFinanciamiento = 1 'Contado
310           oDoFactOrdenServ.idPaciente = oDOCajaComprobantePago.idPaciente
320           oDoFactOrdenServ.idPuntoCarga = lnIdPuntoCarga
330           oDoFactOrdenServ.IdTipoFinanciamiento = oDOCajaComprobantePago.IdTipoFinanciamiento
              
340           oDoFactOrdenServ.idUsuario = oDoFactOrdenServPagos.IdUsuarioAuditoria
350           oDoFactOrdenServ.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
360           oDoFactOrdenServ.IdUsuarioDespacho = oDoFactOrdenServPagos.IdUsuarioAuditoria
370           If oDoFactOrdenServ.idCuentaAtencion > 0 Then
380              oDoFactOrdenServ.IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFactOrdenServ.idCuentaAtencion, Format(oDoFactOrdenServ.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), Format(Now, "hh:mm"))
390           Else
400              oDoFactOrdenServ.IdServicioPaciente = 0
410           End If
420           If Not oFactOrdenServ.Insertar(oDoFactOrdenServ) Then
430               bProcesoOK = False: GoTo ErrorManager
440           End If
450           oDoFactOrdenServPagos.IdOrden = oDoFactOrdenServ.IdOrden
460           If Not oFactOrdenServPagos.Insertar(oDoFactOrdenServPagos) Then
470               bProcesoOK = False: GoTo ErrorManager
480           End If
490           rsFacturacionProductos.MoveFirst
500           Do While Not rsFacturacionProductos.EOF
510              oDoFacturacionServicioDespacho.cantidad = rsFacturacionProductos!cantidad
520              oDoFacturacionServicioDespacho.IdOrden = oDoFactOrdenServ.IdOrden
530              oDoFacturacionServicioDespacho.idProducto = rsFacturacionProductos!idProducto
540              oDoFacturacionServicioDespacho.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
550              oDoFacturacionServicioDespacho.Precio = rsFacturacionProductos!PrecioUnitario
560              oDoFacturacionServicioDespacho.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
570              If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDespacho) Then
580                   bProcesoOK = False: GoTo ErrorManager
590              End If
600              oDoFacturacionServicioPagos.cantidad = rsFacturacionProductos!cantidad
610              oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
620              oDoFacturacionServicioPagos.idProducto = rsFacturacionProductos!idProducto
630              oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
640              oDoFacturacionServicioPagos.Precio = rsFacturacionProductos!PrecioUnitario
650              oDoFacturacionServicioPagos.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
660              If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
670                   bProcesoOK = False: GoTo ErrorManager
680              End If
690              rsFacturacionProductos.MoveNext
700           Loop
710       Else
              'Se ha creado en FACTURACION, Es una PREVENTA (SERVICIOS)
720           If Not oFactOrdenServPagos.Modificar(oDoFactOrdenServPagos) Then
730               bProcesoOK = False: GoTo ErrorManager
740           End If
750           oDoFactOrdenServ.IdOrden = oDoFactOrdenServPagos.IdOrden
760           If Not oFactOrdenServ.SeleccionarPorId(oDoFactOrdenServ) Then
770               bProcesoOK = False: GoTo ErrorManager
780           End If
790           oDoFactOrdenServ.IdEstadoFacturacion = 4 'Pagado
800           If Not oFactOrdenServ.Modificar(oDoFactOrdenServ) Then
810               bProcesoOK = False: GoTo ErrorManager
820           End If
830       End If
                  
'          'Si la orden es de Admsión de Consulta Externa
'       If lnIdPuntoCarga = 6 Then
'              'Generar una solicitud automatica, actualizar el estado de la cita a pagado
'           If Not oCitas.ActualizarCitaPagadaDEBB(oDoAtencion.idAtencion, ml_IdUsuarioAuditoria) Then
'               bProcesoOK = False: GoTo ErrorManager
'           End If
'       End If
                  
'       Dim oDoCajaNroDocumento As New DOCajaNroDocumento
'       oDoCajaNroDocumento.idCaja = oDOCajaComprobantePago.idCaja
'       oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
'          Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOCajaComprobantePago.idCaja, _
'                                                                   oDOCajaComprobantePago.idTipoComprobante)
'       oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
'       oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
'       If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
'           bProcesoOK = False: GoTo ErrorManager
'       End If
          'Receta automatica debb-03/04/2012
960       If lnIdReceta > 0 Then
970           With oCommand
980             .CommandType = adCmdStoredProc
990             Set .ActiveConnection = oConexion
1000            .CommandTimeout = 150
1010            .CommandText = "ActualizaRecetaCabecerarXEstadoYComprobantePago"
1020            Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oDOCajaComprobantePago.IdComprobantePago)
1030            .Parameters.Append oParameter
1040            Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta)
1050            .Parameters.Append oParameter
1060            .Execute
1070          End With
1080          Set oCommand = Nothing
              'debb-18/05/2016 (inicio)
              Dim oDoRecetaDetalleItem As New DoRecetaDetalleItem, oRecetaDetalleItem As New RecetaDetalleItem
1090          Set oRecetaDetalleItem.Conexion = oConexion
              oDoRecetaDetalleItem.IdReceta = lnIdReceta
              oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
1100          rsFacturacionProductos.MoveFirst
1110          Do While Not rsFacturacionProductos.EOF
1120             oDoRecetaDetalleItem.CantidadDespachada = 1
1130             oDoRecetaDetalleItem.DocumentoDespacho = " "
1140             oDoRecetaDetalleItem.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
                 oDoRecetaDetalleItem.idItem = rsFacturacionProductos!idProducto
                 oDoRecetaDetalleItem.idEstado = sghRecetaEstados.sighRecetaConBoleta
1170             If Not oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) Then
1180                bProcesoOK = False: GoTo ErrorManager
1190             End If
1200             rsFacturacionProductos.MoveNext
1210          Loop
1220          Set oDoRecetaDetalleItem = Nothing
1230          Set oRecetaDetalleItem = Nothing
              'debb-18/05/2016 (fin)
1240      End If    '
1250      oConexion.CommitTrans
1260      oConexion.Close
1270      Set oConexion = Nothing
1280      ms_MensajeError = ""
1290      CajaComprobantePagoServicioAgregarHISTORICO = True
          
ErrorManager:

1300      If CajaComprobantePagoServicioAgregarHISTORICO = False Then

1310          ms_MensajeError = Err.Description
1320          oConexion.RollbackTrans
1330          oConexion.Close
1340          Set oConexion = Nothing
1350          ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                             "Function CajaComprobantePagoServicioAgregar", "dllFactUCGestionCaja.cls")  'debb-02/05/2016
                  
1360      End If
1370      Set oDOFacturacionServicios = Nothing
1380      Set oDoFacturacionBienes = Nothing
1390      Set oDOSeguros = Nothing
1400      Set oFacturacionServicios = Nothing
1410      Set oFacturacionBienes = Nothing
1420      Set oFacturacionSeguros = Nothing
1430      Set oFactOrdenesServicio = Nothing
1440      Set oCajaComprobantePago = Nothing
1450      Set oCajaNroDocumento = Nothing
1460      Set oFactOrdenServPagos = Nothing
1470      Set oDoFacturacionServicioPagos = Nothing
1480      Set oFacturacionServicioPagos = Nothing
1490      Set oFacturacionServicioDespacho = Nothing
1500      Set oDoFacturacionServicioDespacho = Nothing
1510      Set oDoFactOrdenServ = Nothing
1520      Set oFactOrdenServ = Nothing
1530      Set oCitas = Nothing
1540      Set mo_ReglasSeguridad = Nothing
End Function




