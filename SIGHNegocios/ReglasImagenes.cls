VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para IMAGENES
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
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
'Modificado por yamill palomino
Function ImagMovimientoSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, _
                                                      ldFechaEgresoFin As Date, lcPacienteFiltro As String, _
                                                      lnUsandoFmovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoSeleccionarPorFechasPuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSeleccionarPorFechasPuntoCarga"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 50, lcPacienteFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnUsandoFmovimiento", adInteger, adParamInput, 0, lnUsandoFmovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoSeleccionarPorFechasPuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function






Function ImagMovimientoImagenesAgregar(oDOImagMovimiento As DoImagMovimiento, _
         oDoImagMovimientoImagenes As DoImagMovimientoImagenes, oDoPaciente As DOPaciente, _
         oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, idCuentaAtencion As Long, _
         rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
         lnIdReceta As Long, lcBoleta As String, lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
         lcDx As String, lcDxCodigo As String, lcUPS As String, lnEpsPorcentaje As Double, lnSiIdCita As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoImagenes As New ImagMovimientoImagenes
    Dim oPacientes As New Pacientes
    Dim oFactOrdenesServicio As New FactOrdenServicio
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
    Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
    Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
    Dim oFacturacionServicioPagos As New FacturacionServicioPagos
    Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
    Dim oFacturacionServicioFinan As New FacturacionServicioFin
    Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
    Dim oFactServicioPagos As New FactOrdenServicioPagos
    Dim oDoImagMovimientoCPT As New DoImagMovimientoCPT
    Dim oImagMovimientoCPT As New ImagMovimientoCPT
    Dim oHistoria As New HistoriasClinicas
    Dim oDoRecetaDetalle As New DoRecetaDetalle, oRecetaDetalle As New RecetaDetalle
    Dim oDoRecetaCabecera As New DoRecetaCabecera, oRecetaCabecera As New RecetaCabecera
    Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
    Dim oDoSiCitas As New DoSiCitas
    Dim oSiCitas As New SiCitas
    Dim bProcesoOK As Boolean
    Dim rsPacientes As New Recordset
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oRsTmp1 As New Recordset
    Dim lnIdComprobantePago As Long
    Dim lcSql As String, lnImporteTotal As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoImagenes.Conexion = oConexion
    Set oPacientes.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oImagMovimientoCPT.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalleItem.Conexion = oConexion
    Set oSiCitas.Conexion = oConexion
    '
    If Not oImagMovimiento.Insertar(oDOImagMovimiento) Then
            bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarIIMAGA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "A", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")               'ListBarItems.idListItem
    '
    If idCuentaAtencion > 0 Then
        'Paciente registrado ingresando Nro Cuenta
        If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
                bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIIMAGA
        End If
        oDoImagMovimientoImagenes.IdOrden = oDoFactOrdenServ.IdOrden
    Else
        'Paciente registrado ingresando Nro Boleta
        If oDoPaciente.idPaciente = 0 Then
            Set rsPacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDoPaciente)
            If rsPacientes.RecordCount > 0 Then
                oDoPaciente.idPaciente = rsPacientes.Fields!idPaciente
            Else
                'oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
                'If Not oPacientes.Insertar(oDOPAciente) Then
                '        bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIIMAGA
                'End If
            End If
        End If
        oDoFactOrdenServ.IdOrden = oDoImagMovimientoImagenes.IdOrden
        'If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
        '        bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIIMAGA
        'End If
        oDoFactOrdenServ.idPaciente = oDoPaciente.idPaciente
        oDoFactOrdenServ.IdPuntoCarga = oDOImagMovimiento.IdPuntoCarga
        oDoFactOrdenServ.fechaDespacho = oDOImagMovimiento.fecha
        If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
                bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIIMAGA
        End If
        'actualiza datos del Paciente en BOLETA
        lnIdComprobantePago = 0
        If oDoImagMovimientoImagenes.IdComprobantePago > 0 Then
           lnIdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
           

            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "CajaComprobantesPagoModificaIdPacienteYRazonSocial"
              Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDoPaciente.idPaciente): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, Left(Trim(oDoPaciente.ApellidoPaterno) & " " & Trim(oDoPaciente.ApellidoMaterno) & " " & Trim(oDoPaciente.PrimerNombre), 50)): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
  
        End If
        'Actualiza Paquetes (si fuera el caso)
        If oDoImagMovimientoImagenes.IdComprobantePago > 0 Then
            lnIdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
            rsFacturacionProductosCPT.MoveFirst
            Do While Not rsFacturacionProductosCPT.EOF
            

               
                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "FacturacionPaquetesModificaAtencionId"
                  Set oParameter = .CreateParameter("@AtencionId", adInteger, adParamInput, 0, oDOImagMovimiento.IdMovimiento): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, rsFacturacionProductosCPT.Fields!idProducto): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
                  .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
                           
               rsFacturacionProductosCPT.MoveNext
            Loop
        End If
    End If
    '
    oDoImagMovimientoImagenes.IdMovimiento = oDOImagMovimiento.IdMovimiento
    If Not oImagMovimientoImagenes.Insertar(oDoImagMovimientoImagenes) Then
            bProcesoOK = False: ms_MensajeError = oImagMovimientoImagenes.MensajeError: GoTo TerminarIIMAGA
    End If
    '
    If idCuentaAtencion > 0 Then
        'Paciente registrado ingresando Nro Cuenta
        If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
          '***********Contado ***********
          With oDOFactOrdenServicioPagos
              .FechaCreacion = oDoFactOrdenServ.FechaCreacion
              .IdEstadoFacturacion = 1   'Registrado
              .IdOrden = oDoFactOrdenServ.IdOrden
              .idUsuario = oDoFactOrdenServ.IdUsuarioAuditoria
              .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
          End With
          If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
              bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo TerminarIIMAGA
          End If
        End If
   End If
   lnImporteTotal = 0
   'Agregar Productos (insumos)
   If rsFacturacionProductos.RecordCount > 0 Then
        oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
        oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
        rsFacturacionProductos.MoveFirst
        Do While Not rsFacturacionProductos.EOF
            oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
            oDoImagMovimientoDetalle.idProductoCPT = rsFacturacionProductos.Fields!idProductoCPT
            oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
            oDoImagMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
            If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
                   bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarIIMAGA
            End If
             rsFacturacionProductos.MoveNext
        Loop
   End If
   'Agregar Productos (cpt)
   oDoImagMovimientoCPT.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
   If rsFacturacionProductosCPT.RecordCount = 0 Then
      bProcesoOK = False: ms_MensajeError = "No hay CPT": GoTo TerminarIIMAGA
   End If
   rsFacturacionProductosCPT.MoveFirst
   Do While Not rsFacturacionProductosCPT.EOF
       oDoImagMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
       oDoImagMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
       oDoImagMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
       oDoImagMovimientoCPT.Importe = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
       oDoImagMovimientoCPT.resultadoAutomatico = IIf(rsFacturacionProductosCPT!resultadoAutomatico = True, 1, 0)
       If Not oImagMovimientoCPT.Insertar(oDoImagMovimientoCPT) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoCPT.MensajeError: GoTo TerminarIIMAGA
       End If
       
       If idCuentaAtencion > 0 Then
           'Paciente registrado ingresando Nro Cuenta
           oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
           oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
           oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
           oDoFacturacionServicioDesp.idProducto = rsFacturacionProductosCPT.Fields!idProducto
           oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
           oDoFacturacionServicioDesp.total = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
           If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIIMAGA
           End If
           If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
               '***********Contado ***********
               With oDoFacturacionServicioPagos
                   .cantidad = rsFacturacionProductosCPT.Fields!cantidad
                   .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                   .idProducto = rsFacturacionProductosCPT.Fields!idProducto
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
                   .total = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
               End With
               If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:  GoTo TerminarIIMAGA
               End If
           Else
               '***********Sis,Soat,exoneracion, convenio ***********
               With oDoFacturacionServicioFinanc
                   .CantidadFinanciada = rsFacturacionProductosCPT.Fields!cantidad
                   .FechaAutoriza = oDoFactOrdenServ.FechaCreacion
                   .IdOrden = oDoFactOrdenServ.IdOrden
                   .idProducto = rsFacturacionProductosCPT.Fields!idProducto
                   .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
                   .PrecioFinanciado = rsFacturacionProductosCPT.Fields!PrecioUnitario
                   .TotalFinanciado = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
                   .idFuenteFinanciamiento = oDoFactOrdenServ.idFuenteFinanciamiento
               End With
               If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIIMAGA
               End If
               lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
           End If
           '
        End If
        rsFacturacionProductosCPT.MoveNext
    Loop
    'Recetas
    'debb-18/05/2016 (inicio)
    'If lnIdReceta > 0 Then
        If idCuentaAtencion > 0 Then
            'Paciente con algun Seguro
            oDoRecetaCabecera.idReceta = lnIdReceta
            If Not oRecetaCabecera.SeleccionarPorId(oDoRecetaCabecera) Then
               bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo TerminarIIMAGA
            End If
        Else
            'Paciente con Boleta
            oDoRecetaCabecera.IdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
            If Not oRecetaCabecera.SeleccionarPorIdComprobantePagoRecetaDetalleItem(oDoRecetaCabecera) Then
               bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo TerminarIIMAGA
            End If
            lnIdReceta = oDoRecetaCabecera.idReceta
        End If
        If oDoRecetaCabecera.idCuentaAtencion > 0 Then
           oDoRecetaCabecera.IdEstado = 2
           oDoRecetaCabecera.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
           If Not oRecetaCabecera.modificar(oDoRecetaCabecera) Then
               bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo TerminarIIMAGA
           End If
           Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnIdReceta, oConexion)
           Set oRecetaDetalle.Conexion = oConexion
           If oRsTmp1.RecordCount > 0 Then
              oDoRecetaDetalle.idReceta = lnIdReceta
              oDoRecetaDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
              rsFacturacionProductosCPT.MoveFirst
              Do While Not rsFacturacionProductosCPT.EOF
                   oDoRecetaDetalle.idItem = rsFacturacionProductosCPT.Fields!idProducto
                   If Not oRecetaDetalle.SeleccionarPorIdRecetaIditem(oDoRecetaDetalle) Then
                      bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarIIMAGA
                   End If
                   oDoRecetaDetalle.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
                   If Not oRecetaDetalle.modificar(oDoRecetaDetalle, True) Then
                      bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarIIMAGA
                   End If
                   oDoRecetaDetalleItem.idReceta = lnIdReceta
                   oDoRecetaDetalleItem.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
                   oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDOImagMovimiento.IdMovimiento))
                   oDoRecetaDetalleItem.idItem = rsFacturacionProductosCPT.Fields!idProducto
                   oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
                   
                   If oDoImagMovimientoImagenes.IdComprobantePago = 0 Then
                        If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
                            bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIIMAGA
                        End If
                   Else
                        oDoRecetaDetalleItem.IdEstado = sghRecetaEstados.sighRecetaDespachada
                        oDoRecetaDetalleItem.IdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
                        If oRecetaDetalleItem.modificar(oDoRecetaDetalleItem) = False Then
                            bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIIMAGA
                        End If
                   End If
                   rsFacturacionProductosCPT.MoveNext
              Loop
           End If
        End If
        
    'End If
    'debb-18/05/2016 (fin)
    '
    If lnEpsPorcentaje > 0 And oDoFactOrdenServ.idCuentaAtencion > 0 Then
       mo_AdminAdmision.AsignaPorcentajeEPS oConexion, oDoFactOrdenServ.idCuentaAtencion, lnEpsPorcentaje, mo_lnIdTablaLISTBARITEMS, _
                       sghAgregar, oDoFactOrdenServ, oFactOrdenesServicio, oDoFacturacionServicioDesp, _
                       oFacturacionServicioDespacho, oDOFactOrdenServicioPagos, oFactServicioPagos, _
                       oDoFacturacionServicioPagos, oFacturacionServicioPagos, oDoFactOrdenServ.FechaCreacion, _
                       oDoFactOrdenServ.IdServicioPaciente, oDoFactOrdenServ.idPaciente, _
                       oDoFactOrdenServ.IdUsuarioAuditoria, oDoFactOrdenServ.IdOrden, lnImporteTotal, _
                       sghPuntosCargaBasicos.sghPtoCargaCaja, 0, 0
       mo_idOrdenPago = mo_AdminAdmision.idOrdenPago
    End If
    '
    If ImagenesActualizaInterconexionSistemaIMG(sghAgregar, oDoPaciente, rsFacturacionProductosCPT, _
                                                oDoImagMovimientoImagenes, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion, _
                                                oDOImagMovimiento.IdPuntoCarga) = False Then
       bProcesoOK = False
    End If
    '
    If lnSiIdCita > 0 Then
       oDoSiCitas.idCitaSI = lnSiIdCita
       oDoSiCitas.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
       If oSiCitas.SeleccionarPorId(oDoSiCitas) = True Then
          oDoSiCitas.IdMovimiento = oDOImagMovimiento.IdMovimiento
          oDoSiCitas.IdEstado = sghSiCitasEstados.sghSiCitaConTomaMuestra
          If oSiCitas.modificar(oDoSiCitas) = True Then
          End If
       End If
    End If
TerminarIIMAGA:
    If bProcesoOK Then
        ImagMovimientoImagenesAgregar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoImagenesAgregar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oDoSiCitas = Nothing
    Set oSiCitas = Nothing
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoImagenes = Nothing
    Set oPacientes = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oImagMovimientoDetalle = Nothing
    Set oDoImagMovimientoDetalle = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set mo_AdminAdmision = Nothing
    Set oDoImagMovimientoCPT = Nothing
    Set oImagMovimientoCPT = Nothing
    Set oHistoria = Nothing
    Set oRsTmp1 = Nothing
    Set mo_reglasComunes = Nothing
    Set oDoRecetaDetalle = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaCabecera = Nothing
    Set oRecetaDetalleItem = Nothing
    Set oDoRecetaDetalleItem = Nothing
End Function

Function DevuelveNombreArchivoImagenes(lnIdMovimiento As Long, lnIdPacienteParaImagenes As Long, oRsItemsCpt As Recordset) As String
    On Error GoTo errDevN
    If lcBuscaParametro.SeleccionaFilaParametro(567) = "W" Then
       DevuelveNombreArchivoImagenes = ""
       Exit Function
    End If
    
    Dim oRsTmp2 As New Recordset
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim lcEquipo As String, lcRuta As String
    DevuelveNombreArchivoImagenes = "Arc.Img: "
    oRsItemsCpt.Filter = ""
    oRsItemsCpt.MoveFirst
    Do While Not oRsItemsCpt.EOF
        lcEquipo = "": lcRuta = ""
        Set oRsTmp2 = mo_ReglasCaja.FactCatalogoServiciosSeleccionarPorId(oRsItemsCpt!idProducto)
        If oRsTmp2.RecordCount > 0 Then
           If Not IsNull(oRsTmp2!equipoCodigo) Then
              lcEquipo = oRsTmp2!equipoCodigo
              oRsTmp2.Close
              Set oRsTmp2 = mo_reglasComunes.InteroperaEquiposSeleccionarPorId(lcEquipo)
              If oRsTmp2.RecordCount > 0 Then
                 lcRuta = oRsTmp2!Ruta & "\"
              End If
           End If
        End If
        oRsTmp2.Close
        DevuelveNombreArchivoImagenes = DevuelveNombreArchivoImagenes & lcRuta & "ord" & Trim(Str(lnIdMovimiento)) & _
                                       "cpt" & oRsItemsCpt!Codigo & "pac" & Trim(Str(lnIdPacienteParaImagenes)) & " >> "
        oRsItemsCpt.MoveNext
    Loop
    oRsItemsCpt.MoveFirst
errDevN:
    Set mo_ReglasCaja = Nothing
    Set oRsTmp2 = Nothing
    Set mo_reglasComunes = Nothing
End Function

Function DevuelveIdPacienteParaLeerImagenes(oDoPacientes As DOPaciente, oDoImagMovimientoImagenes As DoImagMovimientoImagenes) As Long
            If oDoPacientes.idPaciente = 0 Then
               DevuelveIdPacienteParaLeerImagenes = 9000000 + oDoImagMovimientoImagenes.IdOrden
            Else
               DevuelveIdPacienteParaLeerImagenes = oDoPacientes.idPaciente
            End If
End Function

Function ImagenesActualizaInterconexionSistemaIMG(mi_opcion As sghOpciones, oDoPacientes As DOPaciente, _
                                                        rsFacturacionProductos As Recordset, _
                                                        oDoImagMovimientoImagenes As DoImagMovimientoImagenes, _
                                                        oDoFactOrdenServ As DoFactOrdenServ, lcBoleta As String, _
                                                        lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
                                                        lcDx As String, lcDxCodigo As String, lcUPS As String, _
                                                        oConexion As Connection, lnIdPuntoCarga As sghPuntosCargaBasicos) As Boolean
    On Error GoTo LabActInter
    Dim TodoOk As Boolean
    Dim DoInteroperaPacs As New DoInteroperaPacs
    Dim InteroperaPacs As New InteroperaPacs
    Dim oDOPacientes1 As New DOPaciente
    Dim oPacientes1 As New Pacientes
    Dim oConexionExterna As New Connection
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oRsTmp2 As New Recordset
    Dim lcEquipo As String, lcRuta As String
    TodoOk = False
    ImagenesActualizaInterconexionSistemaIMG = True
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.BeginTrans
    Set InteroperaPacs.Conexion = oConexionExterna
    Set oPacientes1.Conexion = oConexion
    If oDoPacientes.idPaciente > 0 Then
        oDOPacientes1.idPaciente = oDoPacientes.idPaciente
        If oPacientes1.SeleccionarPorId(oDOPacientes1) Then
        End If
    End If
    
    DoInteroperaPacs.IdOrden = oDoImagMovimientoImagenes.IdMovimiento
    DoInteroperaPacs.IdUsuarioAuditoria = sighentidades.Usuario
    If mi_opcion <> sghAgregar Then
        If InteroperaPacs.SeleccionarPorId(DoInteroperaPacs) = False Then
           mi_opcion = sghAgregar
        Else
           If DoInteroperaPacs.IdEstado <> sghInterLabEstados.sghIntRegistradoEnGalenhos Then
              ImagenesActualizaInterconexionSistemaIMG = False
              MsgBox "No se podrá Modificar/Eliminar porque ya se leyó  en el SISTEMA DE IMAGENES", vbInformation, ""
              GoTo LabActInter
           End If
        End If
    End If
    If mi_opcion <> sghAgregar Then
        If InteroperaPacs.Eliminar(DoInteroperaPacs) = False Then
            GoTo LabActInter
        End If
    End If
    If mi_opcion <> sghEliminar Then
        rsFacturacionProductos.Filter = "ResultadoAutomatico=true"
        If rsFacturacionProductos.RecordCount > 0 Then
            'DoInteroperaPacs.cama = lcCAma
            'DoInteroperaPacs.dx = lcDx
            'DoInteroperaPacs.dxCodigo = lcDxCodigo
            
            DoInteroperaPacs.fechaHoraMuestra = oDoFactOrdenServ.FechaHoraRealizaCpt
            If oDoPacientes.idPaciente = 0 Then
               DoInteroperaPacs.PacienteFNacimiento = IIf(oDoImagMovimientoImagenes.FechaNacimiento = 0, CDate("01/01/1990"), oDoImagMovimientoImagenes.FechaNacimiento)
            Else
               DoInteroperaPacs.PacienteFNacimiento = oDOPacientes1.FechaNacimiento
            End If
            DoInteroperaPacs.PacienteNroHistoria = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oDOPacientes1.NroHistoriaClinica)), False)
            DoInteroperaPacs.IdEstado = sghInterLabEstados.sghIntRegistradoEnGalenhos
            
            DoInteroperaPacs.MedPaterno = "MEDICO"
            DoInteroperaPacs.MedMaterno = "EXTERNO"
            DoInteroperaPacs.MedNombres = "."
            If oDoPacientes.idPaciente = 0 Then
                DoInteroperaPacs.Medico = "MEDICO EXTERNO ."
                DoInteroperaPacs.medicoDNI = "00000000"
            Else
                DoInteroperaPacs.Medico = lcMedico
                DoInteroperaPacs.medicoDNI = lcMedicoDNI
                Set oRsTmp2 = mo_reglasComunes.EmpleadosSeleccionarPorDNI(lcMedicoDNI)
                If oRsTmp2.RecordCount > 0 Then
                    DoInteroperaPacs.MedPaterno = oRsTmp2!ApellidoPaterno
                    DoInteroperaPacs.MedMaterno = oRsTmp2!ApellidoMaterno
                    DoInteroperaPacs.MedNombres = oRsTmp2!Nombres
                End If
                oRsTmp2.Close
            End If
            
            DoInteroperaPacs.FteFinanciamiento = "PAGANTE"
            DoInteroperaPacs.TipoServicio = "EXTERNO"
            If oDoPacientes.idPaciente = 0 Or oDoImagMovimientoImagenes.idCuentaAtencion = 0 Then
                DoInteroperaPacs.PacPaterno = sighentidades.RetornaPrimerNombre(oDoImagMovimientoImagenes.Paciente)
                DoInteroperaPacs.PacMaterno = sighentidades.RetornaSegundoNombre(oDoImagMovimientoImagenes.Paciente)
                DoInteroperaPacs.PacPrimerNombre = sighentidades.RetornaTercerNombre(oDoImagMovimientoImagenes.Paciente)
                DoInteroperaPacs.PacSegundoNombre = ""
            Else
                DoInteroperaPacs.PacPaterno = oDoPacientes.ApellidoPaterno
                DoInteroperaPacs.PacMaterno = oDoPacientes.ApellidoMaterno
                DoInteroperaPacs.PacPrimerNombre = oDoPacientes.PrimerNombre
                DoInteroperaPacs.PacSegundoNombre = oDoPacientes.SegundoNombre
                Set oRsTmp2 = mo_AdminAdmision.AtencionesSeleccionarTSyFFporCuenta(oDoImagMovimientoImagenes.idCuentaAtencion)
                If oRsTmp2.RecordCount > 0 Then
                    If Not IsNull(oRsTmp2!FuenteFinanciamiento) Then
                       DoInteroperaPacs.FteFinanciamiento = Left(oRsTmp2!FuenteFinanciamiento, 20)
                    End If
                    If Not IsNull(oRsTmp2!TiposServicio) Then
                       DoInteroperaPacs.TipoServicio = Left(oRsTmp2!TiposServicio, 20)
                    End If
                End If
                oRsTmp2.Close
            End If
            'DoInteroperaPacs.IdOrden = IIf(lcBoleta = "", Trim(Str(oDoImagMovimientoImagenes.idCuentaAtencion)), lcBoleta)
            DoInteroperaPacs.Paciente = DoInteroperaPacs.PacPaterno & " " & DoInteroperaPacs.PacMaterno & " " & _
                                         DoInteroperaPacs.PacPrimerNombre & IIf(DoInteroperaPacs.PacSegundoNombre = "", "", _
                                         " " & DoInteroperaPacs.PacSegundoNombre)
            DoInteroperaPacs.PacienteDocumento = IIf(oDoPacientes.idPaciente = 0, "00000000", oDOPacientes1.NroDocumento)
            'DoInteroperaPacs.PacienteEmail = IIf(oDoPacientes.idPaciente = 0, "@", oDOPacientes1.email)
            'DoInteroperaPacs.PacienteFiliacion = IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 1, "P", _
                                                  IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 2, "S", _
                                                  IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 4, "C", "O")))
                                                  
'            If oDoPacientes.idPaciente = 0 Then
'               DoInteroperaPacs.PacienteId = 9000000 + oDoImagMovimientoImagenes.IdOrden
'            Else
'               DoInteroperaPacs.PacienteId = oDoPacientes.idPaciente
'            End If
            DoInteroperaPacs.PacienteId = DevuelveIdPacienteParaLeerImagenes(oDoPacientes, oDoImagMovimientoImagenes)
            
            DoInteroperaPacs.PacienteIdTipoDocumento = IIf(oDoPacientes.idPaciente = 0, 1, oDOPacientes1.idDocIdentidad)
            DoInteroperaPacs.PacienteSexo = IIf(oDoImagMovimientoImagenes.idTipoSexo = 1, "M", "F")
            'DoInteroperaPacs.servicioCodigo = lcUPS
            'DoInteroperaPacs.medicoColegiatura = oDoImagMovimientoImagenes.colegiatura
            Select Case lnIdPuntoCarga
            Case sghPtoCargaTomografia
                 DoInteroperaPacs.PuntoCarga = "Tomografia"
            Case sghPtoCargaRayosX
                 DoInteroperaPacs.PuntoCarga = "RayosX"
            Case sghPtoCargaEcogObstetrica
                 DoInteroperaPacs.PuntoCarga = "EcografíaObstet"
            Case Else
                 DoInteroperaPacs.PuntoCarga = "EcografiaGeneral"
            End Select
            rsFacturacionProductos.MoveFirst
            Do While Not rsFacturacionProductos.EOF
                lcEquipo = "": lcRuta = ""
                Set oRsTmp2 = mo_ReglasCaja.FactCatalogoServiciosSeleccionarPorId(rsFacturacionProductos!idProducto)
                If oRsTmp2.RecordCount > 0 Then
                   If Not IsNull(oRsTmp2!equipoCodigo) Then
                      lcEquipo = oRsTmp2!equipoCodigo
                      oRsTmp2.Close
                      Set oRsTmp2 = mo_reglasComunes.InteroperaEquiposSeleccionarPorId(lcEquipo)
                      If oRsTmp2.RecordCount > 0 Then
                         lcRuta = oRsTmp2!Ruta
                      End If
                   End If
                End If
                oRsTmp2.Close
                DoInteroperaPacs.EquipoRuta = lcRuta
                DoInteroperaPacs.equipoCodigo = lcEquipo
                DoInteroperaPacs.cpt = rsFacturacionProductos!nombreProducto
                DoInteroperaPacs.cptCodigo = rsFacturacionProductos!Codigo
                DoInteroperaPacs.LLAVE = Trim(Str(DoInteroperaPacs.IdOrden)) & DoInteroperaPacs.cptCodigo
                If InteroperaPacs.Insertar(DoInteroperaPacs) = False Then
                  GoTo LabActInter
                End If
                rsFacturacionProductos.MoveNext
            Loop
        End If
    End If
    TodoOk = True
LabActInter:
    If TodoOk = True Then
       oConexionExterna.CommitTrans
    Else
       oConexionExterna.RollbackTrans
    End If
    oConexionExterna.Close
    Set DoInteroperaPacs = Nothing
    Set InteroperaPacs = Nothing
    Set oConexionExterna = Nothing
    Set oDOPacientes1 = Nothing
    Set oPacientes1 = Nothing
    Set mo_ReglasCaja = Nothing
    Set oRsTmp2 = Nothing
    Set mo_reglasComunes = Nothing
    Exit Function
    Resume
End Function


Function ImagMovimientoImagenesModificar(oDOImagMovimiento As DoImagMovimiento, _
         oDoImagMovimientoImagenes As DoImagMovimientoImagenes, oDoPaciente As DOPaciente, _
         oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, idCuentaAtencion As Long, _
         rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
         lnIdReceta As Long, lcBoleta As String, lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
         lcDx As String, lcDxCodigo As String, lcUPS As String, lnEpsPorcentaje As Double, lnIdOrdenPago As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoImagenes As New ImagMovimientoImagenes
    Dim oPacientes As New Pacientes
    Dim oFactOrdenesServicio As New FactOrdenServicio
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
    Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
    Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
    Dim oFacturacionServicioPagos As New FacturacionServicioPagos
    Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
    Dim oFacturacionServicioFinan As New FacturacionServicioFin
    Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
    Dim oFactServicioPagos As New FactOrdenServicioPagos
    Dim oDoImagMovimientoCPT As New DoImagMovimientoCPT
    Dim oImagMovimientoCPT As New ImagMovimientoCPT
    Dim oHistoria As New HistoriasClinicas
    Dim oDoRecetaDetalle As New DoRecetaDetalle, oRecetaDetalle As New RecetaDetalle
    Dim oDoRecetaCabecera As New DoRecetaCabecera, oRecetaCabecera As New RecetaCabecera
    Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
    Dim bProcesoOK As Boolean
    Dim oRsTmp1 As New Recordset
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim lcSql As String, lnImporteTotal As Double
    '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoImagenes.Conexion = oConexion
    Set oPacientes.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oImagMovimientoCPT.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    Set oRecetaDetalleItem.Conexion = oConexion
   '
'   If IdCuentaAtencion = 0 Then
'      If oDoPaciente.idPaciente = 0 Then
'        'oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
'        'If Not oPacientes.Insertar(oDOPAciente) Then
'        '   bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
'        'End If
'      Else
'        If Not oPacientes.modificar(oDoPaciente, False) Then
'           bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
'        End If
'      End If
'   End If
   '
   If Not oImagMovimientoImagenes.modificar(oDoImagMovimientoImagenes) Then
      bProcesoOK = False: ms_MensajeError = oImagMovimientoImagenes.MensajeError: GoTo TerminarIMMAGA
   End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoImagMovimientoImagenes.IdUsuarioAuditoria, "M", oDoImagMovimientoImagenes.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    oDoFactOrdenServ.IdOrden = oDoImagMovimientoImagenes.IdOrden
'    If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
'       bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
'    End If
    oDoFactOrdenServ.IdUsuarioAuditoria = oDoImagMovimientoImagenes.IdUsuarioAuditoria
    If oDoPaciente.idPaciente > 0 And oDoFactOrdenServ.idPaciente <> oDoPaciente.idPaciente Then
        oDoFactOrdenServ.idPaciente = oDoPaciente.idPaciente
        If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
           bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
        End If
    Else
        If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
           bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
        End If
    End If
   'Elimina Productos
   If idCuentaAtencion > 0 Then
        Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden, oConexion)
        If oRsTmp1.RecordCount > 0 Then
          If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
               oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
               oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
               oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
               If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIMMAGA
               End If
          End If
        End If
        oDoFacturacionServicioFinanc.IdOrden = oDoFactOrdenServ.IdOrden
        oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
        If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
           bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIMMAGA
        End If
        If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
           bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIMMAGA
        End If
        
   End If
   oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
   If Not oImagMovimientoDetalle.Eliminar(oDoImagMovimientoDetalle) Then
          bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarIMMAGA
   End If
   oDoImagMovimientoCPT.IdMovimiento = oDOImagMovimiento.IdMovimiento
   If Not oImagMovimientoCPT.Eliminar(oDoImagMovimientoCPT) Then
           bProcesoOK = False: ms_MensajeError = oImagMovimientoCPT.MensajeError: GoTo TerminarIMMAGA
   End If
   'Agregar Productos (insumos)
   If rsFacturacionProductos.RecordCount > 0 Then
        oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
        oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
        rsFacturacionProductos.MoveFirst
        Do While Not rsFacturacionProductos.EOF
            oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
            oDoImagMovimientoDetalle.idProductoCPT = rsFacturacionProductos.Fields!idProductoCPT
            oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
            oDoImagMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
            If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
                   bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarIMMAGA
            End If
             rsFacturacionProductos.MoveNext
        Loop
   End If
   'Agregar Productos (cpt)
   oDoImagMovimientoCPT.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
   If rsFacturacionProductosCPT.RecordCount = 0 Then
      bProcesoOK = False: ms_MensajeError = "No existen ITEMS": GoTo TerminarIMMAGA
   End If
   lnImporteTotal = 0
   rsFacturacionProductosCPT.MoveFirst
   Do While Not rsFacturacionProductosCPT.EOF
       oDoImagMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
       oDoImagMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
       oDoImagMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
       oDoImagMovimientoCPT.Importe = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2)  'rsFacturacionProductosCPT.Fields!totalPorPagar
       oDoImagMovimientoCPT.resultadoAutomatico = IIf(rsFacturacionProductosCPT!resultadoAutomatico = True, 1, 0)
       If Not oImagMovimientoCPT.Insertar(oDoImagMovimientoCPT) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoCPT.MensajeError: GoTo TerminarIMMAGA
       End If
       
       If idCuentaAtencion > 0 Then
           'Paciente registrado ingresando Nro Cuenta
           oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
           oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
           oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
           oDoFacturacionServicioDesp.idProducto = rsFacturacionProductosCPT.Fields!idProducto
           oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
           oDoFacturacionServicioDesp.total = Round(rsFacturacionProductosCPT.Fields!PrecioUnitario * rsFacturacionProductosCPT.Fields!cantidad, 2)   'rsFacturacionProductosCPT.Fields!totalPorPagar
           If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIMMAGA
           End If
           If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
               '***********Contado ***********
               With oDoFacturacionServicioPagos
                   .cantidad = rsFacturacionProductosCPT.Fields!cantidad
                   .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                   .idProducto = rsFacturacionProductosCPT.Fields!idProducto
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
                   .total = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
               End With
               If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:  GoTo TerminarIMMAGA
               End If
           Else
               '***********Sis,Soat,exoneracion, convenio ***********
               With oDoFacturacionServicioFinanc
                   .CantidadFinanciada = rsFacturacionProductosCPT.Fields!cantidad
                   .FechaAutoriza = oDoFactOrdenServ.FechaCreacion
                   .IdOrden = oDoFactOrdenServ.IdOrden
                   .idProducto = rsFacturacionProductosCPT.Fields!idProducto
                   .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
                   .PrecioFinanciado = rsFacturacionProductosCPT.Fields!PrecioUnitario
                   .TotalFinanciado = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2)  'rsFacturacionProductosCPT.Fields!totalPorPagar
                   .idFuenteFinanciamiento = oDoFactOrdenServ.idFuenteFinanciamiento
               End With
               If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIMMAGA
               End If
               lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
           End If
           '
        End If
        rsFacturacionProductosCPT.MoveNext
    Loop
    
    'Receta automatica debb-03/04/2012
    'debb-18/05/2016 (inicio)
    If oDoImagMovimientoImagenes.IdComprobantePago > 0 Then
        'Paciente con Boleta
        oDoRecetaCabecera.IdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
        If oRecetaCabecera.SeleccionarPorIdComprobantePagoRecetaDetalleItem(oDoRecetaCabecera) = True Then
           lnIdReceta = oDoRecetaCabecera.idReceta
        End If
    End If
    If lnIdReceta > 0 Then
       If oDoImagMovimientoImagenes.IdComprobantePago = 0 Then
          RecetaAutomaticaCantidadDespachadaCERO lnIdReceta, oConexion
       End If
       Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnIdReceta, oConexion)
       Set oRecetaDetalle.Conexion = oConexion
       If oRsTmp1.RecordCount > 0 Then
            oDoRecetaDetalle.idReceta = lnIdReceta
            oDoRecetaDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
            rsFacturacionProductosCPT.MoveFirst
            Do While Not rsFacturacionProductosCPT.EOF
               oDoRecetaDetalle.idItem = rsFacturacionProductosCPT.Fields!idProducto
               If Not oRecetaDetalle.SeleccionarPorIdRecetaIditem(oDoRecetaDetalle) Then
                  bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarIMMAGA
               End If
               oDoRecetaDetalle.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
               If Not oRecetaDetalle.modificar(oDoRecetaDetalle, True) Then
                  bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarIMMAGA
               End If
               oDoRecetaDetalleItem.idReceta = lnIdReceta
               oDoRecetaDetalleItem.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
               oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDOImagMovimiento.IdMovimiento))
               oDoRecetaDetalleItem.idItem = rsFacturacionProductosCPT.Fields!idProducto
               oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
               If oDoImagMovimientoImagenes.IdComprobantePago = 0 Then
                    If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
                         bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIMMAGA
                    End If
               Else
                    oDoRecetaDetalleItem.IdEstado = sghRecetaEstados.sighRecetaDespachada
                    oDoRecetaDetalleItem.IdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
                    If oRecetaDetalleItem.modificar(oDoRecetaDetalleItem) = False Then
                         bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIMMAGA
                    End If
               End If
               rsFacturacionProductosCPT.MoveNext
            Loop
       End If
       oRsTmp1.Close
    End If
    'debb-18/05/2016 (fin)
    If lnEpsPorcentaje > 0 And oDoFactOrdenServ.idCuentaAtencion > 0 Then
       mo_AdminAdmision.AsignaPorcentajeEPS oConexion, oDoFactOrdenServ.idCuentaAtencion, lnEpsPorcentaje, mo_lnIdTablaLISTBARITEMS, _
                       sghModificar, oDoFactOrdenServ, oFactOrdenesServicio, oDoFacturacionServicioDesp, _
                       oFacturacionServicioDespacho, oDOFactOrdenServicioPagos, oFactServicioPagos, _
                       oDoFacturacionServicioPagos, oFacturacionServicioPagos, oDoFactOrdenServ.FechaCreacion, _
                       oDoFactOrdenServ.IdServicioPaciente, oDoFactOrdenServ.idPaciente, _
                       oDoFactOrdenServ.IdUsuarioAuditoria, oDoFactOrdenServ.IdOrden, lnImporteTotal, _
                       sghPuntosCargaBasicos.sghPtoCargaCaja, lnIdOrdenPago, 0
       mo_idOrdenPago = mo_AdminAdmision.idOrdenPago
    End If
    '
    If ImagenesActualizaInterconexionSistemaIMG(sghModificar, oDoPaciente, rsFacturacionProductosCPT, _
                                                oDoImagMovimientoImagenes, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion, _
                                                oDOImagMovimiento.IdPuntoCarga) = False Then
       bProcesoOK = False
    End If
    
TerminarIMMAGA:
    If bProcesoOK Then
        ImagMovimientoImagenesModificar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoImagenesModificar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoImagenes = Nothing
    Set oPacientes = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oImagMovimientoDetalle = Nothing
    Set oDoImagMovimientoDetalle = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set mo_AdminAdmision = Nothing
    Set oDoImagMovimientoCPT = Nothing
    Set oImagMovimientoCPT = Nothing
    Set oHistoria = Nothing
    Set oRsTmp1 = Nothing
    Set mo_reglasComunes = Nothing
    Set oDoRecetaDetalle = Nothing
    Set oRecetaDetalle = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaCabecera = Nothing
    Set oRecetaDetalleItem = Nothing
    Set oDoRecetaDetalleItem = Nothing
End Function

Function ImagMovimientoImagenesAnular(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoImagenes As DoImagMovimientoImagenes, _
                                      oDoPaciente As DOPaciente, oDoFactOrdenServ As DoFactOrdenServ, _
                                      rsFacturacionProductos As Recordset, idCuentaAtencion As Long, _
                                      rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                      mo_lcNombrePc As String, lnIdReceta As Long, lcBoleta As String, lcMedicoDNI As String, _
                                      lcMedico As String, lcCAma As String, lcDx As String, lcDxCodigo As String, _
                                      lcUPS As String, lnIdOrdenPago As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoImagenes As New ImagMovimientoImagenes
    Dim oPacientes As New Pacientes
    Dim oFactOrdenesServicio As New FactOrdenServicio
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
    Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
    Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
    Dim oFacturacionServicioPagos As New FacturacionServicioPagos
    Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
    Dim oFacturacionServicioFinan As New FacturacionServicioFin
    Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
    Dim oFactServicioPagos As New FactOrdenServicioPagos
    Dim oHistoria As New HistoriasClinicas
    Dim oFacturacionServicioFin As New FacturacionServicioFin
    Dim oDoRecetaCabecera As New DoRecetaCabecera, oRecetaCabecera As New RecetaCabecera
    Dim oDoSiCitas As New DoSiCitas
    Dim oSiCitas As New SiCitas
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim bProcesoOK As Boolean
    Dim lnIdComprobantePago As Long
    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.BeginTrans
    
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoImagenes.Conexion = oConexion
    Set oPacientes.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    Set oHistoria.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oRecetaCabecera.Conexion = oConexion
    Set oSiCitas.Conexion = oConexion
   '
    oDOImagMovimiento.IdImagEstado = sghEstadoTabla.sghAnulado    'Anulado
    If Not oImagMovimiento.modificar(oDOImagMovimiento) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarIAMAGA
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "E", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
   'Anula Cabecera Productos
   If idCuentaAtencion > 0 Then
        oDoFactOrdenServ.IdOrden = oDoImagMovimientoImagenes.IdOrden
        If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
           bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
        End If
        oDoFactOrdenServ.IdUsuarioAuditoria = oDoImagMovimientoImagenes.IdUsuarioAuditoria
        oDoFactOrdenServ.IdEstadoFacturacion = 9  'anulado
        If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
           bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
        End If
        Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden, oConexion)
        If oRsTmp1.RecordCount > 0 Then
          If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
               oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
               If Not oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
               End If
               oDOFactOrdenServicioPagos.IdEstadoFacturacion = 9 'anulado
               oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoImagMovimientoImagenes.IdUsuarioAuditoria
               If Not oFactServicioPagos.modificar(oDOFactOrdenServicioPagos) Then
                   bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
               End If
          End If
        End If
        If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacion(oDoFactOrdenServ.IdOrden, 9) Then
           bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo TerminarIAMAGA
        End If
   Else
        'Actualiza Paquetes (si fuera el caso)
        If oDoImagMovimientoImagenes.IdComprobantePago > 0 Then
             lnIdComprobantePago = oDoImagMovimientoImagenes.IdComprobantePago
             rsFacturacionProductosCPT.MoveFirst
             Do While Not rsFacturacionProductosCPT.EOF
             
                
                With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "FacturacionPaquetesModificarAtencionIdNull"
                  Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, rsFacturacionProductosCPT.Fields!idProducto): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
                   .Execute
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
            
                rsFacturacionProductosCPT.MoveNext
             Loop
        End If
   End If
   
   
    'Receta automatica debb-03/04/2012
    If lnIdReceta > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalle lnIdReceta, oDoImagMovimientoImagenes.idCuentaAtencion, oDoImagMovimientoImagenes.IdComprobantePago, oConexion
    ElseIf oDoImagMovimientoImagenes.IdComprobantePago > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalleSoloPagantes oDoImagMovimientoImagenes.IdComprobantePago, oConexion
    End If
    '
    If mo_AdminAdmision.EPSeliminaOrdenPago(oConexion, lnIdOrdenPago, oDoImagMovimientoImagenes.IdUsuarioAuditoria) = False Then
    End If
    '
    If ImagenesActualizaInterconexionSistemaIMG(sghEliminar, oDoPaciente, rsFacturacionProductosCPT, _
                                                oDoImagMovimientoImagenes, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion, _
                                                oDOImagMovimiento.IdPuntoCarga) = False Then
       bProcesoOK = False
    End If
   
    oDoSiCitas.IdMovimiento = oDOImagMovimiento.IdMovimiento
    oDoSiCitas.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
    If oSiCitas.SiCitasSeleccionarPorIdMovimiento(oDoSiCitas) = True Then
       oDoSiCitas.IdMovimiento = 0
       oDoSiCitas.IdEstado = sghSiCitasEstados.sghSiCitaActiva
       If oSiCitas.modificar(oDoSiCitas) = True Then
       End If
    End If

TerminarIAMAGA:
    If bProcesoOK Then
        ImagMovimientoImagenesAnular = True
        oConexion.CommitTrans
    Else
        ImagMovimientoImagenesAnular = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oDoSiCitas = Nothing
    Set oSiCitas = Nothing
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoImagenes = Nothing
    Set oPacientes = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oImagMovimientoDetalle = Nothing
    Set oDoImagMovimientoDetalle = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set oFactServicioPagos = Nothing
    Set mo_AdminAdmision = Nothing
    Set oHistoria = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oRsTmp1 = Nothing
    Set oCommand = Nothing
    Set mo_ReglasImagenes = Nothing
    Set oDoRecetaCabecera = Nothing
    Set oRecetaCabecera = Nothing
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoSeleccionarPorId(IdMovimiento As Long, Optional oConexion1 As Connection) As DoImagMovimiento
Dim oConexion As New ADODB.Connection
Dim oImagMovimiento As New ImagMovimiento
Dim oDOImagMovimiento As New DoImagMovimiento
    Set ImagMovimientoSeleccionarPorId = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    If oConexion1 Is Nothing Then
       Set oImagMovimiento.Conexion = oConexion
    Else
       Set oImagMovimiento.Conexion = oConexion1
    End If
    oDOImagMovimiento.IdMovimiento = IdMovimiento
    If oImagMovimiento.SeleccionarPorId(oDOImagMovimiento) Then
        Set ImagMovimientoSeleccionarPorId = oDOImagMovimiento
    Else
        ms_MensajeError = oImagMovimiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oDOImagMovimiento = Nothing
End Function

    'Modificado por yamill palomino

Function ImagMovimientoImagenesSeleccionarPorIdMovimiento(lnIdMovimiento As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoImagenesSeleccionarPorIdMovimiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoImagenesSeleccionarPorIdMovimiento"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute

   End With
   Set ImagMovimientoImagenesSeleccionarPorIdMovimiento = oRecordset

   
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

    'Modificado por yamill palomino

Function ImagMovimientoDetalleSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoDetalleSeleccionarPorIdMovimiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoDetalleSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoDetalleSeleccionarPorIdMovimiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

    'Modificado por yamill palomino

Function ImagMovimientoCPTSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoCPTSeleccionarPorIdMovimiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoCPTSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoCPTSeleccionarPorIdMovimiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
   
    'Modificado por yamill palomino
Function ImagMovimientoImagenesSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoImagenesSeleccionarPorIdComprobantePago = Nothing
    ms_MensajeError = ""
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
        .CommandText = "ImagMovimientoImagenesSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set ImagMovimientoImagenesSeleccionarPorIdComprobantePago = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
        Set oConexion = Nothing
   End If
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoIngresosSeleccionarPorId(IdMovimiento As Long, Optional oConexion1 As Connection) As DoImagMovimientoIngresos
Dim oConexion As New ADODB.Connection
Dim oImagMovimientoIngresos As New ImagMovimientoIngresos
Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Set ImagMovimientoIngresosSeleccionarPorId = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    If oConexion1 Is Nothing Then
       Set oImagMovimientoIngresos.Conexion = oConexion
    Else
       Set oImagMovimientoIngresos.Conexion = oConexion1
    End If
    oDoImagMovimientoIngresos.IdMovimiento = IdMovimiento
    If oImagMovimientoIngresos.SeleccionarPorId(oDoImagMovimientoIngresos) Then
        Set ImagMovimientoIngresosSeleccionarPorId = oDoImagMovimientoIngresos
    Else
        ms_MensajeError = oImagMovimientoIngresos.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimientoIngresos = Nothing
    Set oDoImagMovimientoIngresos = Nothing
End Function

    'Modificado por yamill palomino
Function ImagMovimientoSalidasSeleccionarPorId(IdMovimiento As Long) As DoImagMovimientoSalidas
Dim oConexion As New ADODB.Connection
Dim oImagMovimientoSalidas As New ImagMovimientoSalidas
Dim oDoImagMovimientoSalidas As New DoImagMovimientoSalidas
    Set ImagMovimientoSalidasSeleccionarPorId = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oImagMovimientoSalidas.Conexion = oConexion
    oDoImagMovimientoSalidas.IdMovimiento = IdMovimiento
    If oImagMovimientoSalidas.SeleccionarPorId(oDoImagMovimientoSalidas) Then
        Set ImagMovimientoSalidasSeleccionarPorId = oDoImagMovimientoSalidas
    Else
        ms_MensajeError = oImagMovimientoSalidas.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimientoSalidas = Nothing
    Set oDoImagMovimientoSalidas = Nothing
End Function
   
    'Modificado por yamill palomino

Function ImagMovimientoSeleccionarPorFechasPuntoCargaIngresos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoSeleccionarPorFechasPuntoCargaIngresos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSeleccionarPorFechasPuntoCargaIngresos"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoSeleccionarPorFechasPuntoCargaIngresos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

   
    'Modificado por yamill palomino
Function ImagMovimientoSeleccionarPorFechasPuntoCargaSalidas(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoSeleccionarPorFechasPuntoCargaSalidas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSeleccionarPorFechasPuntoCargaSalidas"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoSeleccionarPorFechasPuntoCargaSalidas = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function ImagMovimientoIngresosAgregar(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoIngresos As DoImagMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoIngresos As New ImagMovimientoIngresos
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoIngresos.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    If Not oImagMovimiento.Insertar(oDOImagMovimiento) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarImagIng
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "A", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '
    oDoImagMovimientoIngresos.IdMovimiento = oDOImagMovimiento.IdMovimiento
    If Not oImagMovimientoIngresos.Insertar(oDoImagMovimientoIngresos) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimientoIngresos.MensajeError: GoTo TerminarImagIng
    End If
    '
   oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
   rsFacturacionProductos.MoveFirst
   Do While Not rsFacturacionProductos.EOF
       oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
       oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
       If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagIng
       End If
       rsFacturacionProductos.MoveNext
   Loop
TerminarImagIng:
    If bProcesoOK Then
        ImagMovimientoIngresosAgregar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoIngresosAgregar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoIngresos = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function

Function ImagMovimientoIngresosModificar(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoIngresos As DoImagMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoIngresos As New ImagMovimientoIngresos
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoIngresos.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    If Not oImagMovimientoIngresos.modificar(oDoImagMovimientoIngresos) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimientoIngresos.MensajeError: GoTo TerminarImagMod
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "M", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '
   oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
   If Not oImagMovimientoDetalle.Eliminar(oDoImagMovimientoDetalle) Then
          bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagMod
   End If
   rsFacturacionProductos.MoveFirst
   Do While Not rsFacturacionProductos.EOF
       oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
       oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
       If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagMod
       End If
       rsFacturacionProductos.MoveNext
   Loop
TerminarImagMod:
    If bProcesoOK Then
        ImagMovimientoIngresosModificar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoIngresosModificar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoIngresos = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function

Function ImagMovimientoIngresosAnular(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoIngresos As DoImagMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoIngresos As New ImagMovimientoIngresos
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoIngresos.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    oDOImagMovimiento.IdImagEstado = 0 'Anulado
    If Not oImagMovimiento.modificar(oDOImagMovimiento) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarImagAnu
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "E", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
TerminarImagAnu:
    If bProcesoOK Then
        ImagMovimientoIngresosAnular = True
        oConexion.CommitTrans
    Else
        ImagMovimientoIngresosAnular = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoIngresos = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function


Function ImagMovimientoSalidasAgregar(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoSalidas As DoImagMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoSalidas As New ImagMovimientoSalidas
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoSalidas.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    If Not oImagMovimiento.Insertar(oDOImagMovimiento) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarImagSalIng
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "A", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '
    oDoImagMovimientoSalidas.IdMovimiento = oDOImagMovimiento.IdMovimiento
    If Not oImagMovimientoSalidas.Insertar(oDoImagMovimientoSalidas) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimientoSalidas.MensajeError: GoTo TerminarImagSalIng
    End If
    '
   oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
   rsFacturacionProductos.MoveFirst
   Do While Not rsFacturacionProductos.EOF
       oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
       oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
       If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagSalIng
       End If
       rsFacturacionProductos.MoveNext
   Loop
TerminarImagSalIng:
    If bProcesoOK Then
        ImagMovimientoSalidasAgregar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoSalidasAgregar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoSalidas = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function


Function ImagMovimientoSalidasModificar(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoSalidas As DoImagMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoSalidas As New ImagMovimientoSalidas
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoSalidas.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    If Not oImagMovimientoSalidas.modificar(oDoImagMovimientoSalidas) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimientoSalidas.MensajeError: GoTo TerminarImagSalMod
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "M", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
    '
   oDoImagMovimientoDetalle.IdMovimiento = oDOImagMovimiento.IdMovimiento
   oDoImagMovimientoDetalle.IdUsuarioAuditoria = oDOImagMovimiento.IdUsuarioAuditoria
   If Not oImagMovimientoDetalle.Eliminar(oDoImagMovimientoDetalle) Then
          bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagSalMod
   End If
   rsFacturacionProductos.MoveFirst
   Do While Not rsFacturacionProductos.EOF
       oDoImagMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
       oDoImagMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
       If Not oImagMovimientoDetalle.Insertar(oDoImagMovimientoDetalle) Then
              bProcesoOK = False: ms_MensajeError = oImagMovimientoDetalle.MensajeError: GoTo TerminarImagSalMod
       End If
       rsFacturacionProductos.MoveNext
   Loop
TerminarImagSalMod:
    If bProcesoOK Then
        ImagMovimientoSalidasModificar = True
        oConexion.CommitTrans
    Else
        ImagMovimientoSalidasModificar = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoSalidas = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function

Function ImagMovimientoSalidasAnular(oDOImagMovimiento As DoImagMovimiento, oDoImagMovimientoSalidas As DoImagMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagMovimiento As New ImagMovimiento
    Dim oImagMovimientoSalidas As New ImagMovimientoSalidas
    Dim oImagMovimientoDetalle As New ImagMovimientoDetalle
    Dim oDoImagMovimientoDetalle As New DoImagMovimientoDetalle
    Dim bProcesoOK As Boolean
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oImagMovimiento.Conexion = oConexion
    Set oImagMovimientoSalidas.Conexion = oConexion
    Set oImagMovimientoDetalle.Conexion = oConexion
    '
    oDOImagMovimiento.IdImagEstado = 0   'Anulado
    If Not oImagMovimiento.modificar(oDOImagMovimiento) Then
       bProcesoOK = False: ms_MensajeError = oImagMovimiento.MensajeError: GoTo TerminarImagSalAnul
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagMovimiento.IdUsuarioAuditoria, "E", oDOImagMovimiento.IdMovimiento, "ImagMovimiento/" & oDOImagMovimiento.MovTipo & "/" & Trim(Str(oDOImagMovimiento.IdMovimiento)), oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")              'ListBarItems.idListItem
TerminarImagSalAnul:
    If bProcesoOK Then
        ImagMovimientoSalidasAnular = True
        oConexion.CommitTrans
    Else
        ImagMovimientoSalidasAnular = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagMovimiento = Nothing
    Set oImagMovimientoSalidas = Nothing
    Set oImagMovimientoDetalle = Nothing
End Function

' MODIFICADO POR YAMILL PALOMINO

Function ImagMotivoSalidasSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMotivoSalidasSeleccionarTodos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMotivoSalidasSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMotivoSalidasSeleccionarTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

' MODIFICADO POR YAMILL PALOMINO

Function ImagRecojeExamenSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagRecojeExamenSeleccionarTodos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagRecojeExamenSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagRecojeExamenSeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function



End Function


'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnIdProducto As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter    '3->Por Codigo     4->Por fecha
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoIngresosSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoIngresosSeleccionarPorIdMovimiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoIngresosSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoIngresosSeleccionarPorIdMovimiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoSalidasSeleccionarPorIdMovimiento(lnIdMovimiento As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoSalidasSeleccionarPorIdMovimiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSalidasSeleccionarPorIdMovimiento"
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set ImagMovimientoSalidasSeleccionarPorIdMovimiento = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


' MODIFICADO POR YAMILL PALOMINO
Function ImagRecojeExamenSeleccionaPorID(lnIdRecojeExamen As Long) As String
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    ImagRecojeExamenSeleccionaPorID = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagRecojeExamenSeleccionaPorID"
        Set oParameter = .CreateParameter("@idRecojeExamen", adInteger, adParamInput, 0, lnIdRecojeExamen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   
    If oRecordset.RecordCount > 0 Then
       ImagRecojeExamenSeleccionaPorID = oRecordset.Fields!RecojeExamen
    Else
       ImagRecojeExamenSeleccionaPorID = ""
    End If

   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
   
End Function


Function ImagMovimientoImagenesSeleccionarPorFechasPToInsumos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoImagenesSeleccionarPorFechasPToInsumos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoImagenesSeleccionarPorFechasPtoInsumos"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter    '3->Por Codigo     4->Por fecha
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoImagenesSeleccionarPorFechasPToInsumos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoImagenesSeleccionarPorIdOrden(lnIdOrden As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set ImagMovimientoImagenesSeleccionarPorIdOrden = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoImagenesSeleccionarPorIdOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagMovimientoImagenesSeleccionarPorIdOrden = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

' MODIFICADO POR YAMILL PALOMINO

Function CptHistoricosPorPaciente(lnIdPaciente As Long, lnIdPuntoCarga As Long, lnIdMovimiento As Long) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set CptHistoricosPorPaciente = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CptHistoricosPorPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CptHistoricosPorPaciente = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function

' MODIFICADO POR YAMILL PALOMINO

Function ImagenesDevuelveAlertaDeCantidades() As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set ImagenesDevuelveAlertaDeCantidades = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagenesDevuelveAlertaDeCantidades"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ImagenesDevuelveAlertaDeCantidades = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
 
  
End Function


' MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoImagenesXidOrden(lnIdOrden As Long, oConexion As Connection) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set ImagMovimientoImagenesXidOrden = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoImagenesXidOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set ImagMovimientoImagenesXidOrden = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
  
End Function



' MODIFICADO POR YAMILL PALOMINO
Function ImagMovimientoParaSip2000(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    Set ImagMovimientoParaSip2000 = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ImagMovimientoParaSip2000"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set ImagMovimientoParaSip2000 = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
End Function



Sub RecetaAutomaticaCantidadDespachadaCERO(lnIdReceta As Long, oConexion As Connection)
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaAutomatica"
        Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
End Sub

Sub ActualizaRecetaCabeceraYDetalle(lnIdReceta As Long, idCuentaAtencion As Long, _
                                    lnIdComprobantePago As Long, oConexion As Connection)
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
        With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaRecetaCabeceraYDetalle"
              Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idCuentaAtencion): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing

End Sub

Sub ActualizaRecetaCabeceraYDetalleSoloPagantes(lnIdComprobantePago As Long, oConexion As Connection)
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "ActualizaRecetaCabeceraYDetalleSoloPagantes"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          .Execute
        End With
        Set oCommand = Nothing

End Sub



Sub ActualizaRecetaCabeceraYDetalleSoloBoletaAnulada(lnIdComprobantePago As Long, oConexion As Connection)
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
        With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaRecetaCabeceraYDetalleSoloBoletaAnulada"
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing

End Sub

'==============================================================================
'TIPO DE MODALIDAD DE SALA
'==============================================================================
Function ImagTipoModalidadSalaAgregar(oDOTipoModalidadSala As DOImagTipoModalidadSala, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagTipoModalidadSala As String) As Boolean
    
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    
    ImagTipoModalidadSalaAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagTipoModalidadSala.Conexion = oConexion
    If oImagTipoModalidadSala.Insertar(oDOTipoModalidadSala) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTipoModalidadSala.IdUsuarioAuditoria, "A", oDOTipoModalidadSala.IdTipoModalidadSala, _
                        "ImagTipoModalidadSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagTipoModalidadSala)
        ImagTipoModalidadSalaAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagTipoModalidadSala.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagTipoModalidadSala = Nothing
End Function

Function ImagTipoModalidadSalaModificar(oDOTipoModalidadSala As DOImagTipoModalidadSala, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagTipoModalidadSala As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
        
    ImagTipoModalidadSalaModificar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagTipoModalidadSala.Conexion = oConexion
    If oImagTipoModalidadSala.modificar(oDOTipoModalidadSala) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTipoModalidadSala.IdUsuarioAuditoria, "M", oDOTipoModalidadSala.IdTipoModalidadSala, _
                            "ImagTipoModalidadSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagTipoModalidadSala)
                            
        ImagTipoModalidadSalaModificar = True
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oImagTipoModalidadSala.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagTipoModalidadSala = Nothing
End Function

Function ImagTipoModalidadSalaEliminar(oDOTipoModalidadSala As DOImagTipoModalidadSala, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagTipoModalidadSala As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    
    ImagTipoModalidadSalaEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagTipoModalidadSala.Conexion = oConexion
    If oImagTipoModalidadSala.Eliminar(oDOTipoModalidadSala) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOTipoModalidadSala.IdUsuarioAuditoria, "E", oDOTipoModalidadSala.IdTipoModalidadSala, _
                    "ImagTipoModalidadSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagTipoModalidadSala)
                    
        ImagTipoModalidadSalaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagTipoModalidadSala.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagTipoModalidadSala = Nothing
End Function

Function ImagTipoModalidadSalaSeleccionarPorId(IdTipoModalidadSala As Integer) As DOImagTipoModalidadSala
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    Dim oDOImagTipoModalidadSala As New DOImagTipoModalidadSala
    
    Set ImagTipoModalidadSalaSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagTipoModalidadSala.Conexion = oConexion
    oDOImagTipoModalidadSala.IdTipoModalidadSala = IdTipoModalidadSala
    If oImagTipoModalidadSala.SeleccionarPorId(oDOImagTipoModalidadSala) Then
        Set ImagTipoModalidadSalaSeleccionarPorId = oDOImagTipoModalidadSala
    Else
        Set ImagTipoModalidadSalaSeleccionarPorId = Nothing
    End If
    
    ms_MensajeError = oImagTipoModalidadSala.MensajeError
        
    oConexion.Close
    Set oConexion = Nothing
    Set oImagTipoModalidadSala = Nothing

End Function

Function ImagTipoModalidadSalaVerificarPorNombre(oDOImagTipoModalidadSala As DOImagTipoModalidadSala) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    
    Set ImagTipoModalidadSalaVerificarPorNombre = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagTipoModalidadSala.Conexion = oConexion
    Set ImagTipoModalidadSalaVerificarPorNombre = oImagTipoModalidadSala.VerificarPorNombre(oDOImagTipoModalidadSala)
    
    ms_MensajeError = oImagTipoModalidadSala.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oImagTipoModalidadSala = Nothing
End Function


Function ImagTipoModalidadSalaSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    
    Dim ms_MensajeError  As String
    
    Set ImagTipoModalidadSalaSeleccionarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagTipoModalidadSala.Conexion = oConexion
    Set ImagTipoModalidadSalaSeleccionarTodos = oImagTipoModalidadSala.SeleccionarTodos
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagTipoModalidadSala = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

Function ImagTipoModalidadSalaFiltrarTodos(oDOImagTipoModalidadSala As DOImagTipoModalidadSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagTipoModalidadSala As New ImagTipoModalidadSala
    
    Dim ms_MensajeError  As String
    
    Set ImagTipoModalidadSalaFiltrarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagTipoModalidadSala.Conexion = oConexion
    Set ImagTipoModalidadSalaFiltrarTodos = oImagTipoModalidadSala.FiltrarTodos(oDOImagTipoModalidadSala)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagTipoModalidadSala = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

'======================================================
'SALA DE IMAGENOLOGIA
'======================================================
Function ImagSalaAgregar(oDOSala As DOImagSala, oPuntosCarga As Collection, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSala As String) As Boolean
    
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    ImagSalaAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagSala.Conexion = oConexion
    If oImagSala.Insertar(oDOSala) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOSala.IdUsuarioAuditoria, "A", oDOSala.IdSala, _
                        "ImagSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSala)
        If ImagSalaPuntoCargaAgregarVarios(oDOSala.IdSala, oDOSala.IdUsuarioAuditoria, oPuntosCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "", oConexion) = False Then
            oConexion.RollbackTrans
        Else
            ImagSalaAgregar = True
            oConexion.CommitTrans
        End If
    Else
        ms_MensajeError = oImagSala.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSala = Nothing
End Function

Function ImagSalaModificar(oDOSala As DOImagSala, oPuntosCarga As Collection, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSala As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
        
    ImagSalaModificar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagSala.Conexion = oConexion
    If oImagSala.modificar(oDOSala) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOSala.IdUsuarioAuditoria, "M", oDOSala.IdSala, _
                            "ImagSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSala)
                            
        If ImagSalaPuntoCargaAgregarVarios(oDOSala.IdSala, oDOSala.IdUsuarioAuditoria, oPuntosCarga, _
                            mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "", oConexion) = False Then
            oConexion.RollbackTrans
        Else
            ImagSalaModificar = True
            oConexion.CommitTrans
        End If
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oImagSala.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSala = Nothing
End Function

Function ImagSalaEliminar(oDOSala As DOImagSala, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSala As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    ImagSalaEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagSala.Conexion = oConexion
    If oImagSala.Eliminar(oDOSala) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOSala.IdUsuarioAuditoria, "E", oDOSala.IdSala, _
                    "ImagSala", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSala)
                    
        ImagSalaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagSala.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSala = Nothing
End Function

Function ImagSalaSeleccionarPorId(IdSala As Integer) As DOImagSala
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    Dim oDOImagSala As New DOImagSala
    
    Set ImagSalaSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagSala.Conexion = oConexion
    oDOImagSala.IdSala = IdSala
    If oImagSala.SeleccionarPorId(oDOImagSala) Then
        Set ImagSalaSeleccionarPorId = oDOImagSala
    Else
        Set ImagSalaSeleccionarPorId = Nothing
    End If
    
    ms_MensajeError = oImagSala.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSala = Nothing

End Function

Function ImagSalaVerificarPorNombre(oDOImagSala As DOImagSala) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    Set ImagSalaVerificarPorNombre = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagSala.Conexion = oConexion
    Set ImagSalaVerificarPorNombre = oImagSala.VerificarPorNombre(oDOImagSala)
    
    ms_MensajeError = oImagSala.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSala = Nothing
End Function


Function ImagSalaSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    Dim ms_MensajeError  As String
    
    Set ImagSalaSeleccionarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagSala.Conexion = oConexion
    Set ImagSalaSeleccionarTodos = oImagSala.SeleccionarTodos
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagSala = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

Function ImagSalaFiltrarPorTipoMadalidad(oDOImagSala As DOImagSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    Dim ms_MensajeError  As String
    
    Set ImagSalaFiltrarPorTipoMadalidad = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagSala.Conexion = oConexion
    Set ImagSalaFiltrarPorTipoMadalidad = oImagSala.FiltrarPorTipoModalidad(oDOImagSala)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagSala = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

Function ImagSalaFiltrarTodos(oDOImagSala As DOImagSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagSala As New ImagSala
    
    Dim ms_MensajeError  As String
    
    Set ImagSalaFiltrarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagSala.Conexion = oConexion
    Set ImagSalaFiltrarTodos = oImagSala.FiltrarTodos(oDOImagSala)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagSala = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

'=======================================================
'PUNTOS DE CARGA POR SALA
'=======================================================
Function ImagSalaPuntoCargaAgregar(oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga, _
        mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSalaPuntoCarga As String, _
        oConexion As ADODB.Connection) As Boolean
    
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
    
    ImagSalaPuntoCargaAgregar = False

    Set oImagSalaPuntoCarga.Conexion = oConexion
    If oImagSalaPuntoCarga.Insertar(oDOImagSalaPuntoCarga) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagSalaPuntoCarga.IdUsuarioAuditoria, "A", oDOImagSalaPuntoCarga.IdPuntoCarga, _
                        "ImagSalaPuntoCarga", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSalaPuntoCarga)
        ImagSalaPuntoCargaAgregar = True
    Else
        ms_MensajeError = oImagSalaPuntoCarga.MensajeError
    End If

    Set oImagSalaPuntoCarga = Nothing
End Function

Function ImagSalaPuntoCargaModificar(oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga, _
        mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSalaPuntoCarga As String, _
        oConexion As ADODB.Connection) As Boolean
    
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
        
    ImagSalaPuntoCargaModificar = False

    Set oImagSalaPuntoCarga.Conexion = oConexion
    If oImagSalaPuntoCarga.modificar(oDOImagSalaPuntoCarga) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagSalaPuntoCarga.IdUsuarioAuditoria, "M", oDOImagSalaPuntoCarga.IdPuntoCarga, _
                            "ImagSalaPuntoCarga", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSalaPuntoCarga)
                            
        ImagSalaPuntoCargaModificar = True
    Else
        ms_MensajeError = oImagSalaPuntoCarga.MensajeError
    End If
    Set oImagSalaPuntoCarga = Nothing
End Function

Function ImagSalaPuntoCargaEliminar(oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSalaPuntoCarga As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
    
    ImagSalaPuntoCargaEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagSalaPuntoCarga.Conexion = oConexion
    If oImagSalaPuntoCarga.Eliminar(oDOImagSalaPuntoCarga) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagSalaPuntoCarga.IdUsuarioAuditoria, "E", oDOImagSalaPuntoCarga.IdSala, _
                    "ImagSalaPuntoCarga", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSalaPuntoCarga)
                    
        ImagSalaPuntoCargaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagSalaPuntoCarga.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSalaPuntoCarga = Nothing
End Function

Function ImagSalaPuntoCargaSeleccionarPorId(IdSala As Integer) As DOImagSalaPuntoCarga
    Dim oConexion As New ADODB.Connection
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
    Dim oDOImagSalaPuntoCarga As New DOImagSalaPuntoCarga
    
    Set ImagSalaPuntoCargaSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagSalaPuntoCarga.Conexion = oConexion
    oDOImagSalaPuntoCarga.IdSala = IdSala
    If oImagSalaPuntoCarga.SeleccionarPorId(oDOImagSalaPuntoCarga) Then
        Set ImagSalaPuntoCargaSeleccionarPorId = oDOImagSalaPuntoCarga
    Else
        Set ImagSalaPuntoCargaSeleccionarPorId = Nothing
    End If
    
    ms_MensajeError = oImagSalaPuntoCarga.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oImagSalaPuntoCarga = Nothing

End Function

Function ImagSalaPuntoCargaEliminarPorIdSala(oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSalaPuntoCarga As String, _
        oConexion As ADODB.Connection) As Boolean
    
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
    
    ImagSalaPuntoCargaEliminarPorIdSala = False

    Set oImagSalaPuntoCarga.Conexion = oConexion
    If oImagSalaPuntoCarga.EliminarPorIdSala(oDOImagSalaPuntoCarga) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagSalaPuntoCarga.IdUsuarioAuditoria, "E", oDOImagSalaPuntoCarga.IdPuntoCarga, _
                    "ImagSalaPuntoCarga", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagSalaPuntoCarga)
                    
        ImagSalaPuntoCargaEliminarPorIdSala = True

    Else
        ms_MensajeError = oImagSalaPuntoCarga.MensajeError

    End If
    Set oImagSalaPuntoCarga = Nothing
End Function

Function ImagSalaPuntoCargaAgregarVarios(lIdSala As Long, lIdUsuario As Long, oSalaPuntosCarga As Collection, _
        mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagSalaPuntoCarga As String, _
        oConexion As ADODB.Connection) As Boolean
        
    Dim oDOImagSalaPuntoCarga As New DOImagSalaPuntoCarga
    
    ImagSalaPuntoCargaAgregarVarios = False
    
    oDOImagSalaPuntoCarga.IdSala = lIdSala
    oDOImagSalaPuntoCarga.IdUsuarioAuditoria = lIdUsuario
    
    If ImagSalaPuntoCargaEliminarPorIdSala(oDOImagSalaPuntoCarga, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                        lcDImagSalaPuntoCarga, oConexion) = True Then
                        
        If Not oSalaPuntosCarga Is Nothing Then
            For Each oDOImagSalaPuntoCarga In oSalaPuntosCarga
                oDOImagSalaPuntoCarga.IdSala = lIdSala
                oDOImagSalaPuntoCarga.IdUsuarioAuditoria = lIdUsuario
                If Not ImagSalaPuntoCargaAgregar(oDOImagSalaPuntoCarga, mo_lnIdTablaLISTBARITEMS, _
                            mo_lcNombrePc, lcDImagSalaPuntoCarga, oConexion) Then
'                    ms_MensajeError = oImagSalaPuntoCarga.MensajeError
                    Exit Function
                End If
             Next
        End If
    End If
    ImagSalaPuntoCargaAgregarVarios = True
End Function

Function ImagSalaPuntoCargaFiltrarPorIdSala(oDOImagSalaPuntoCarga As DOImagSalaPuntoCarga) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagSalaPuntoCarga As New ImagSalaPuntoCarga
    
    Dim ms_MensajeError  As String
    
    Set ImagSalaPuntoCargaFiltrarPorIdSala = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagSalaPuntoCarga.Conexion = oConexion
    Set ImagSalaPuntoCargaFiltrarPorIdSala = oImagSalaPuntoCarga.FiltrarPorIdSala(oDOImagSalaPuntoCarga)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagSalaPuntoCarga = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function
'===============================================================
'TIEMPO DE DURACION DE PROCEDMIENTO EN IMAGENES
'===============================================================

Function ImagCatalgoServicioDuracionAgregar(oDOImagCatalgoServicioDuracion As DOImagCatalgoServicioDuracion, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagCatalgoServicioDuracion As String) As Boolean
    
    Dim oConexion As New ADODB.Connection
    Dim oImagCatalgoServicioDuracion As New ImagCatalgoServicioDuracion
    
    ImagCatalgoServicioDuracionAgregar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagCatalgoServicioDuracion.Conexion = oConexion
    If oImagCatalgoServicioDuracion.Insertar(oDOImagCatalgoServicioDuracion) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagCatalgoServicioDuracion.IdUsuarioAuditoria, "A", oDOImagCatalgoServicioDuracion.idProducto, _
                        "ImagCatalgoServicioDuracion", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagCatalgoServicioDuracion)
        ImagCatalgoServicioDuracionAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagCatalgoServicioDuracion.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagCatalgoServicioDuracion = Nothing
End Function

Function ImagCatalgoServicioDuracionModificar(oDOImagCatalgoServicioDuracion As DOImagCatalgoServicioDuracion, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagCatalgoServicioDuracion As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagCatalgoServicioDuracion As New ImagCatalgoServicioDuracion
        
    ImagCatalgoServicioDuracionModificar = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagCatalgoServicioDuracion.Conexion = oConexion
    If oImagCatalgoServicioDuracion.modificar(oDOImagCatalgoServicioDuracion) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagCatalgoServicioDuracion.IdUsuarioAuditoria, "M", oDOImagCatalgoServicioDuracion.idProducto, _
                            "ImagCatalgoServicioDuracion", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagCatalgoServicioDuracion)
                            
        ImagCatalgoServicioDuracionModificar = True
        oConexion.CommitTrans
    Else
        oConexion.RollbackTrans
        ms_MensajeError = oImagCatalgoServicioDuracion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagCatalgoServicioDuracion = Nothing
End Function

Function ImagCatalgoServicioDuracionEliminar(oDOImagCatalgoServicioDuracion As DOImagCatalgoServicioDuracion, _
    mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDImagCatalgoServicioDuracion As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oImagCatalgoServicioDuracion As New ImagCatalgoServicioDuracion
    
    ImagCatalgoServicioDuracionEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oImagCatalgoServicioDuracion.Conexion = oConexion
    If oImagCatalgoServicioDuracion.Eliminar(oDOImagCatalgoServicioDuracion) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOImagCatalgoServicioDuracion.IdUsuarioAuditoria, "E", oDOImagCatalgoServicioDuracion.idProducto, _
                    "ImagCatalgoServicioDuracion", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDImagCatalgoServicioDuracion)
                    
        ImagCatalgoServicioDuracionEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oImagCatalgoServicioDuracion.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oImagCatalgoServicioDuracion = Nothing
End Function

Function ImagCatalgoServicioDuracionSeleccionarPorId(lIdProducto As Integer) As DOImagCatalgoServicioDuracion
    Dim oConexion As New ADODB.Connection
    Dim oImagCatalgoServicioDuracion As New ImagCatalgoServicioDuracion
    Dim oDOImagCatalgoServicioDuracion As New DOImagCatalgoServicioDuracion
    
    Set ImagCatalgoServicioDuracionSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oImagCatalgoServicioDuracion.Conexion = oConexion
    oDOImagCatalgoServicioDuracion.idProducto = lIdProducto
    If oImagCatalgoServicioDuracion.SeleccionarPorId(oDOImagCatalgoServicioDuracion) Then
        Set ImagCatalgoServicioDuracionSeleccionarPorId = oDOImagCatalgoServicioDuracion
    Else
        Set ImagCatalgoServicioDuracionSeleccionarPorId = Nothing
    End If
    ms_MensajeError = oImagCatalgoServicioDuracion.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oImagCatalgoServicioDuracion = Nothing
    

End Function


Function ImagCatalgoServicioDuracionFiltrarTodos(oDOCatalogoServicio As DOCatalogoServicio) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim oImagCatalgoServicioDuracion As New ImagCatalgoServicioDuracion
    
    Dim ms_MensajeError  As String
    
    Set ImagCatalgoServicioDuracionFiltrarTodos = Nothing
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oImagCatalgoServicioDuracion.Conexion = oConexion
    Set ImagCatalgoServicioDuracionFiltrarTodos = oImagCatalgoServicioDuracion.FiltrarTodos(oDOCatalogoServicio)
    
   oConexion.Close
   Set oConexion = Nothing
   Set oImagCatalgoServicioDuracion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

'debb-18/05/2016
Sub ActualizaRecetaCabeceraYDetalleItem(lnIdComprobantePago As Long, oConexion As Connection)
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
        With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaRecetaCabeceraYDetalleItem"
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing

End Sub


'debb-09/09/2016
Function ImagMovimientoImagenesSeleccionarPorIdPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ImagMovimientoImagenesSeleccionarPorIdPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set ImagMovimientoImagenesSeleccionarPorIdPaciente = oRecordset
  Set oCommand = Nothing
  
End Function


Function ImagMovimientoCPTXcuenta(lnIdCuenta As Long, oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ImagMovimientoCPTXcuenta"
    Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set ImagMovimientoCPTXcuenta = oRecordset
  Set oCommand = Nothing
  
End Function

Sub ImagMovimientoImagenesActualizaHcRX(lcPaciente As String, lcHcRx As String)
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ImagMovimientoImagenesActualizaHcRX"
    Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, lcPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@HcRx", adVarChar, adParamInput, 10, lcHcRx): .Parameters.Append oParameter
    .Execute
  End With
  oConexion.Close
  Set oCommand = Nothing
  Set oConexion = Nothing
End Sub


Function ImagMovimientoResultadosSeleccionarPorId(lnIdMovimiento As Long) As Recordset
Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ImagMovimientoResultadosSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ImagMovimientoResultadosSeleccionarPorId = oRecordset
  oConexion.Close
  Set oCommand = Nothing
  Set oConexion = Nothing
End Function

Function TieneResultadoImagenesPorCpt(lnIdMovimiento As Long, lnIdProductoCpt As Long) As Boolean
    TieneResultadoImagenesPorCpt = False
    Dim rsResultados  As New Recordset
    Set rsResultados = ImagMovimientoResultadosSeleccionarPorId(lnIdMovimiento)
    rsResultados.Filter = "idProductoCpt=" & lnIdProductoCpt
    If rsResultados.RecordCount > 0 Then
       TieneResultadoImagenesPorCpt = True
    End If
    rsResultados.Close
    Set rsResultados = Nothing
End Function


Sub ResultadosAutomaticosActualizaImgHaciaGalenhos(lnIdOrden As Long)
    Dim lcMensajeLicencia As String
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    On Error GoTo ErrResAutHacGal
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New Connection
    Dim oConexionExterna As New Connection
    Dim oDoImagMovimientoResultados As New DoImagMovimientoResultados
    Dim oImagMovimientoResultados As New ImagMovimientoResultados
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim lcSql As String, lnItems As Long, lnGrupoItems As Long, lnId1 As Long, lcCpt1 As String, lnIdProductoCpt As Long
    Dim lnIdEmpleadoResultado As Long, lcFechaResultado As String, lnIdItem As Long, lnIdItemGrupo As Long, lnNumero As Long
    Dim lbTodoOK As Boolean, lbEncontroDato As Boolean, lbEsAgregar As Boolean, ml_idProductoCpt As Long
    
    oConexion.CommandTimeout = 300: oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexionExterna.CommandTimeout = 300: oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    lcSql = "select * from InteroperaPacs where idEstado=" & sghInterLabEstados.sghIntTieneResultadoLab & _
            IIf(lnIdOrden > 0, " and idorden=" & lnIdOrden, "")
    oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       lbEncontroDato = True
       oConexion.BeginTrans
       oConexionExterna.BeginTrans
       '
'       If  False Then   'licencia
'           GoTo ErrResAutHacGal
'       End If
       '
       Set oImagMovimientoResultados.Conexion = oConexion
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          ml_idProductoCpt = 0
          Set oRsTmp2 = mo_ReglasCaja.FactCatalogoServiciosSeleccionarPorCodigoOnombre(Trim(oRsTmp1!cptCodigo), "", oConexion)
          If oRsTmp2.RecordCount > 0 Then
             ml_idProductoCpt = oRsTmp2!idProducto
          End If
          oRsTmp2.Close
          lbEsAgregar = False
          Set oRsTmp2 = ImagMovimientoResultadosSeleccionarPorIdPorCpt(oRsTmp1!IdOrden, ml_idProductoCpt, oConexion)
          If oRsTmp2.RecordCount = 0 Then
             lbEsAgregar = True
          End If
          oRsTmp2.Close
          If ml_idProductoCpt > 0 Then
                If ResultadosImagenesActualizar(oDoImagMovimientoResultados, oImagMovimientoResultados, _
                                   lbEsAgregar, oRsTmp1!EquipoRuta, oRsTmp1!IdOrden, ml_idProductoCpt, oRsTmp1!Resultado, _
                                   oRsTmp1!resultadodni, oRsTmp1!ResultadoFechaHora) = False Then
                   GoTo ErrResAutHacGal
                End If
                InteroperaPacsActualizaEstado oConexionExterna, oRsTmp1!IdOrden, Trim(oRsTmp1!cptCodigo), _
                                              sghInterLabEstados.sghIntTieneResultadoLeidoEnGalenhos
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    lbTodoOK = True
ErrResAutHacGal:
    If lbEncontroDato = True Then
        If lbTodoOK = True Then
           oConexion.CommitTrans
           oConexionExterna.CommitTrans
        Else
           oConexion.RollbackTrans
           oConexionExterna.RollbackTrans
        End If
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Set mo_reglasComunes = Nothing
    Exit Sub
    Resume
End Sub


Function ResultadosImagenesActualizar(oDoImagMovimientoResultados As DoImagMovimientoResultados, _
                                 oImagMovimientoResultados As ImagMovimientoResultados, lbEsAgregar As Boolean, _
                                 txtRutaImg As String, ml_idMovimiento As Long, ml_idProductoCpt As Long, _
                                 txtResultadoFinal, DNIresponsable As String, txtFresultado As Date) As Boolean
    ResultadosImagenesActualizar = True
    With oDoImagMovimientoResultados
         .EquipoRuta = txtRutaImg
         .IdMovimiento = ml_idMovimiento
         .idProductoCPT = ml_idProductoCpt
         .IdUsuarioAuditoria = sighentidades.Usuario
         .Resultado = txtResultadoFinal
         .resultadodni = DNIresponsable 'mo_reglasComunes.EmpleadosDevuelveDNI(Val(mo_cmbResponsable.BoundText))
         .resultadoFecha = txtFresultado
    End With
    If lbEsAgregar = True Then
       If oImagMovimientoResultados.Insertar(oDoImagMovimientoResultados) = False Then
          ms_MensajeError = oImagMovimientoResultados.MensajeError: ResultadosImagenesActualizar = False
       End If
    Else
       If oImagMovimientoResultados.modificar(oDoImagMovimientoResultados) = False Then
          ms_MensajeError = oImagMovimientoResultados.MensajeError: ResultadosImagenesActualizar = False
       End If
    End If

End Function


Function ImagMovimientoResultadosSeleccionarPorIdPorCpt(lnIdMovimiento As Long, lnIdProductoCpt As Long, oConexion As Connection) As Recordset
Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
    With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ImagMovimientoResultadosSeleccionarPorIdPorCpt"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProductoCpt): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set ImagMovimientoResultadosSeleccionarPorIdPorCpt = oRecordset
  Set oCommand = Nothing
End Function

Function InteroperaPacsActualizaEstado(oConexionExterna As Connection, lnIdOrden As Long, lcCodigoCpt As String, lnIdEstado As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  ms_MensajeError = ""
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexionExterna
    .CommandTimeout = 150
    .CommandText = "InteroperaPacsActualizaEstado"
    Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, lcCodigoCpt): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, lnIdEstado): .Parameters.Append oParameter
     .Execute
  End With
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description
  Exit Function
End Function



Sub YaExisteUnResultado(btnAceptar As Object, lnIdMovimiento As Long)
        Dim rsResultados As New Recordset
        Set rsResultados = ImagMovimientoResultadosSeleccionarPorId(lnIdMovimiento)
        If rsResultados.RecordCount > 0 Then
           MsgBox "No podrá MODIFICAR/ANULAR porque ya existen RESULTADOS", vbInformation, ""
           btnAceptar.Visible = False
        End If
        rsResultados.Close
        Set rsResultados = Nothing
End Sub






Function SICitasMantenimiento(oDoSiCitas As DoSiCitas, _
         rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
         mo_lcNombrePc As String, mi_opcion As sghOpciones, _
         lbSeAgregaCitaPorCadaItem As Boolean, lnNroMinutos As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oSiCitas As New SiCitas
    Dim oDoSiCitasDetalle As New DoSiCitasDetalle
    Dim oSiCitasDetalle As New SiCitasDetalle
    Dim oPacientes As New Pacientes
    Dim oDoPaciente As New DOPaciente
    Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim bProcesoOK As Boolean
    Dim lnIdComprobantePago As Long, lnCupos As Long
    Dim lcSql As String, lnImporteTotal As Double, lnNumeroCitas As Integer, lnFor As Integer
    '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    '
    bProcesoOK = True
    ms_MensajeError = ""
    Set oSiCitas.Conexion = oConexion
    Set oSiCitasDetalle.Conexion = oConexion
    Set oPacientes.Conexion = oConexion
    If lbSeAgregaCitaPorCadaItem = True Then
    
       'chequea que antes de grabar en mi PC, en otra PC hayan dado cita ese mismo dia           debb-24/12/2018
       Dim oRsTmp1 As New Recordset
       Set oRsTmp1 = siCitasXidprogramacion(oDoSiCitas.idProgramacion, oConexion)
       lnCupos = oRsTmp1.RecordCount
       If lnCupos > 0 Then
           lnCupos = oRsTmp1!cupo
           'la ultima CITA grabada, ya tiene la hora inicio
           If CDate("01/01/2000 " & oDoSiCitas.HoraInicio) = CDate("01/01/2000 " & oRsTmp1!HoraInicio) Then
              oDoSiCitas.HoraInicio = oRsTmp1!horaFinal
              oDoSiCitas.horaFinal = mo_AdminProgramacionMedica.CalculaHoraFinal(oRsTmp1!horaFinal, lnNroMinutos)
              lnCupos = oRsTmp1!cupo
           Else
              'se quiere una HORA eliminada (al medio), pero ya existe
              oRsTmp1.Filter = "horaInicio='" & oDoSiCitas.HoraInicio & "'"
              If oRsTmp1.RecordCount > 0 Then
                 oRsTmp1.Filter = ""
                 oRsTmp1.MoveFirst
                 oDoSiCitas.HoraInicio = oRsTmp1!horaFinal
                 oDoSiCitas.horaFinal = mo_AdminProgramacionMedica.CalculaHoraFinal(oRsTmp1!horaFinal, lnNroMinutos)
                 lnCupos = oRsTmp1!cupo
              End If
           End If
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       '
       oDoSiCitas.llaveTicket = Format(Now, "ddmmyyhhmmss") & oDoSiCitas.HoraInicio & Trim(Str(oDoSiCitas.IdSala))
       rsFacturacionProductosCPT.MoveFirst
       Do While Not rsFacturacionProductosCPT.EOF
           lnNumeroCitas = rsFacturacionProductosCPT!IdFacturacionProducto
           For lnFor = 1 To lnNumeroCitas
                lnCupos = lnCupos + 1
                oDoSiCitas.cupo = lnCupos
                If Not oSiCitas.Insertar(oDoSiCitas) Then
                        bProcesoOK = False: ms_MensajeError = oSiCitas.MensajeError: GoTo TerminarIIMAGA
                End If
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoSiCitas.IdUsuarioAuditoria, _
                                        "A", _
                                        oDoSiCitas.idCitaSI, "SiCitas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
                oDoSiCitasDetalle.llaveTicket = oDoSiCitas.llaveTicket
                oDoSiCitasDetalle.idCitaSI = oDoSiCitas.idCitaSI
                oDoSiCitasDetalle.IdUsuarioAuditoria = oDoSiCitas.IdUsuarioAuditoria
                oDoSiCitasDetalle.cantidad = rsFacturacionProductosCPT.Fields!cantidad
                oDoSiCitasDetalle.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
                oDoSiCitasDetalle.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
                oDoSiCitasDetalle.observaciones = IIf(IsNull(rsFacturacionProductosCPT!poliza), "", rsFacturacionProductosCPT!poliza)
                If Not oSiCitasDetalle.Insertar(oDoSiCitasDetalle) Then
                       bProcesoOK = False: ms_MensajeError = oSiCitasDetalle.MensajeError: GoTo TerminarIIMAGA
                End If
                oDoSiCitas.HoraInicio = oDoSiCitas.horaFinal
                oDoSiCitas.horaFinal = mo_AdminProgramacionMedica.CalculaHoraFinal(oDoSiCitas.HoraInicio, lnNroMinutos)
           Next
           rsFacturacionProductosCPT.MoveNext
        Loop
        
    Else
        If mi_opcion = sghEliminar Then
            oDoSiCitasDetalle.idCitaSI = oDoSiCitas.idCitaSI
            oDoSiCitasDetalle.IdUsuarioAuditoria = oDoSiCitas.IdUsuarioAuditoria
            If Not oSiCitasDetalle.Eliminar(oDoSiCitasDetalle) Then
                   bProcesoOK = False: ms_MensajeError = oSiCitasDetalle.MensajeError: GoTo TerminarIIMAGA
            End If
            If Not oSiCitas.Eliminar(oDoSiCitas) Then
                    bProcesoOK = False: ms_MensajeError = oSiCitas.MensajeError: GoTo TerminarIIMAGA
            End If
        Else
            If mi_opcion = sghAgregar Then
                Set oRsTmp1 = siCitasXidprogramacion(oDoSiCitas.idProgramacion, oConexion)
                lnCupos = oRsTmp1.RecordCount
                If oRsTmp1.RecordCount > 0 Then
                   lnCupos = oRsTmp1!cupo
                End If
                oDoSiCitas.cupo = lnCupos + 1
                oDoSiCitas.llaveTicket = Format(Now, "ddmmyyhhmmss") & oDoSiCitas.HoraInicio & Trim(Str(oDoSiCitas.IdSala))
                If Not oSiCitas.Insertar(oDoSiCitas) Then
                        bProcesoOK = False: ms_MensajeError = oSiCitas.MensajeError: GoTo TerminarIIMAGA
                End If
            Else
                If Not oSiCitas.modificar(oDoSiCitas) Then
                        bProcesoOK = False: ms_MensajeError = oSiCitas.MensajeError: GoTo TerminarIIMAGA
                End If
            End If
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoSiCitas.IdUsuarioAuditoria, _
                                       IIf(mi_opcion = sghAgregar, "A", IIf(mi_opcion = sghModificar, "M", "E")), _
                                       oDoSiCitas.idCitaSI, "SiCitas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")
            oDoSiCitasDetalle.idCitaSI = oDoSiCitas.idCitaSI
            oDoSiCitasDetalle.IdUsuarioAuditoria = oDoSiCitas.IdUsuarioAuditoria
            If mi_opcion = sghModificar Then
               If Not oSiCitasDetalle.Eliminar(oDoSiCitasDetalle) Then
                      bProcesoOK = False: ms_MensajeError = oSiCitasDetalle.MensajeError: GoTo TerminarIIMAGA
               End If
            End If
            rsFacturacionProductosCPT.MoveFirst
            Do While Not rsFacturacionProductosCPT.EOF
               oDoSiCitasDetalle.llaveTicket = oDoSiCitas.llaveTicket
               oDoSiCitasDetalle.cantidad = rsFacturacionProductosCPT.Fields!cantidad
               oDoSiCitasDetalle.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
               oDoSiCitasDetalle.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
               oDoSiCitasDetalle.observaciones = IIf(IsNull(rsFacturacionProductosCPT!poliza), "", rsFacturacionProductosCPT!poliza)
               If Not oSiCitasDetalle.Insertar(oDoSiCitasDetalle) Then
                      bProcesoOK = False: ms_MensajeError = oSiCitasDetalle.MensajeError: GoTo TerminarIIMAGA
               End If
               rsFacturacionProductosCPT.MoveNext
             Loop
         End If
     End If
     '
     If (mi_opcion = sghAgregar Or mi_opcion = sghModificar) And oDoSiCitas.idPaciente > 0 Then
       If oDoSiCitas.TELEFONO <> "" Or oDoSiCitas.Direccion <> "" Then
            oDoPaciente.idPaciente = oDoSiCitas.idPaciente
            If oPacientes.SeleccionarPorId(oDoPaciente) Then
               If Len(oDoSiCitas.TELEFONO) > 0 Then
                  oDoPaciente.TELEFONO = oDoSiCitas.TELEFONO
               End If
               If Len(oDoSiCitas.Direccion) > 0 Then
                  oDoPaciente.DireccionDomicilio = oDoSiCitas.Direccion
               End If
               If oPacientes.modificar(oDoPaciente, False) Then
               End If
               If siCitasActualizaDatos(oDoSiCitas.Paciente, oDoSiCitas.TELEFONO, oDoSiCitas.Direccion, oConexion) = True Then
               End If
            End If
       End If
     End If
     '
TerminarIIMAGA:
    If bProcesoOK Then
        SICitasMantenimiento = True
        oConexion.CommitTrans
    Else
        SICitasMantenimiento = False
        ms_MensajeError = ms_MensajeError & " - " & Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oSiCitas = Nothing
    Set oDoSiCitasDetalle = Nothing
    Set oSiCitasDetalle = Nothing

End Function



Function SiCitasSeleccionarPorIdPuntoCArgaYFecha(lnIdPuntoCarga As Long, ldFechaCita As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasSeleccionarPorIdPuntoCArgaYFecha = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiCitasSeleccionarPorIdPuntoCArgaYFecha"
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@fechaCita", adDBTimeStamp, adParamInput, 8, ldFechaCita): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiCitasSeleccionarPorIdPuntoCArgaYFecha = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function SiCitasLeerPorMedicoYMes(lnIdSala As Long, iMes As Integer, iAnio As Integer) As Collection
Dim oSiCitas As New SIGHDatos.SiCitas
Dim rsProgramaciones As ADODB.Recordset
Dim oProgramaciones As New Collection
Dim oProgramacion As SIGHcomun.DoSiCitas
    
    Set rsProgramaciones = SiCitasPorGrupoYMes(lnIdSala, iMes, iAnio)

   
    Do While Not rsProgramaciones.EOF
       Set oProgramacion = New DoSiCitas
       oProgramacion.IdPuntoCarga = rsProgramaciones!IdPuntoCarga
       oProgramacion.IdSala = rsProgramaciones!IdSala
       oProgramacion.fecha = rsProgramaciones!fecha
       oProgramacion.horaFinal = rsProgramaciones!horaFinal
       oProgramacion.HoraInicio = rsProgramaciones!HoraInicio
       oProgramacion.idCitaSI = rsProgramaciones!idCitaSI
       oProgramacion.Paciente = rsProgramaciones!Paciente
       oProgramacion.IdEstado = rsProgramaciones!IdEstado
       oProgramacion.IdMovimiento = IIf(IsNull(rsProgramaciones!IdMovimiento), 0, rsProgramaciones!IdMovimiento)
       oProgramacion.cupo = IIf(IsNull(rsProgramaciones!cupo), 0, rsProgramaciones!cupo)
       oProgramaciones.Add oProgramacion
       rsProgramaciones.MoveNext
    Loop
    Set SiCitasLeerPorMedicoYMes = oProgramaciones
End Function


Function SiCitasPorGrupoYMes(lnIdSala As Long, iMes As Integer, iAnio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
 
   Set SiCitasPorGrupoYMes = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "SiCitasPorGrupoYMes"
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SiCitasPorGrupoYMes = oRecordset
 
   ms_MensajeError = ""
   oConexion.Close
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SiCitasDetallePorIdentificador(lnIdCitaSI As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
 
   Set SiCitasDetallePorIdentificador = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "SiCitasDetallePorIdentificador"
        Set oParameter = .CreateParameter("@IdCitaSI", adInteger, adParamInput, 0, lnIdCitaSI): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SiCitasDetallePorIdentificador = oRecordset
 
   ms_MensajeError = ""
   oConexion.Close
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SiCitasFiltroPorPaciente(lcPacienteBuscar As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
 
   Set SiCitasFiltroPorPaciente = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "SiCitasFiltroPorPaciente"
        Set oParameter = .CreateParameter("@filtro", adVarChar, adParamInput, 100, Left(lcPacienteBuscar, 100)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SiCitasFiltroPorPaciente = oRecordset
 
   ms_MensajeError = ""
   oConexion.Close
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SiCitasSeleccionarPorId(lnIdCita As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasSeleccionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiCitasSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdSiCita", adInteger, adParamInput, 0, lnIdCita): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiCitasSeleccionarPorId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function SiCitasSalasSeleccionarTodas() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasSalasSeleccionarTodas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiCitasSalasSeleccionarTodas"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiCitasSalasSeleccionarTodas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SiCitasResumenPorDia(lnIdSala As Long, ldFechaCita As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasResumenPorDia = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiCitasResumenPorDia"
        Set oParameter = .CreateParameter("@fechaCita", adDBTimeStamp, adParamInput, 8, ldFechaCita): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiCitasResumenPorDia = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function SiCitasDetalleSeleccionarPorLlaveTicket(lcLlaveTicket As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasDetalleSeleccionarPorLlaveTicket = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiCitasDetalleSeleccionarPorLlaveTicket"
        Set oParameter = .CreateParameter("@IdSiCita", adVarChar, adParamInput, 20, lcLlaveTicket): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiCitasDetalleSeleccionarPorLlaveTicket = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SiCitasPorDia(lnIdSala As Long, ldFechaCita As Date, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiCitasPorDia = Nothing
    ms_MensajeError = ""
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
        .CommandText = "SiCitasPorDia"
        Set oParameter = .CreateParameter("@fechaCita", adDBTimeStamp, adParamInput, 8, ldFechaCita): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set SiCitasPorDia = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oConexion = Nothing
   End If
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SiFechasNoLaborablesXmesAnio(lnIdSala As Long, lnMes As Integer, lnAnio As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SiFechasNoLaborablesXmesAnio = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiFechasNoLaborablesXmesAnio"
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAnio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SiFechasNoLaborablesXmesAnio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SiFechasNoLaborablesEliminaXmesAnio(lnIdSala As Long, lnMes As Integer, lnAnio As Integer, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    SiFechasNoLaborablesEliminaXmesAnio = False
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiFechasNoLaborablesEliminaXmesAnio"
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAnio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        .Execute
   End With
   SiFechasNoLaborablesEliminaXmesAnio = True
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SiFechasNoLaborablesAgregaFeriado(lnIdSala As Long, ldFeriado As Date, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    SiFechasNoLaborablesAgregaFeriado = False
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SiFechasNoLaborablesAgregaFeriado"
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@feriado", adDBTimeStamp, adParamInput, 8, ldFeriado): .Parameters.Append oParameter
        .Execute
   End With
   SiFechasNoLaborablesAgregaFeriado = True
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function siCitasActualizaDatos(lcPaciente As String, lcTelefono As String, lcDireccion As String, oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    siCitasActualizaDatos = False
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "siCitasActualizaDatos"
        Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, lcPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, lcTelefono): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Direccion", adVarChar, adParamInput, 100, lcDireccion): .Parameters.Append oParameter
        .Execute
   End With
   siCitasActualizaDatos = True
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function siCitasSeleccionarPorllaveTicket(lcLlaveTicket As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set siCitasSeleccionarPorllaveTicket = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "siCitasSeleccionarPorllaveTicket"
        Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, lcLlaveTicket): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set siCitasSeleccionarPorllaveTicket = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function SiProgramacionAgregar(oProgramaciones As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New SiProgramacion
Dim oDoProgramacion As New DOSiProgramacion
Dim bEstadoTransaccion As Boolean

    
    SiProgramacionAgregar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgramaciones.Count > 0 Then
        For Each oDoProgramacion In oProgramaciones
            If oProgMedica.Insertar(oDoProgramacion) Then
                '
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoProgramacion.IdUsuarioAuditoria, "A", oDoProgramacion.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
                bEstadoTransaccion = True
            Else
                bEstadoTransaccion = False
                Exit For
            End If
        Next
    Else
        bEstadoTransaccion = True
    End If
    
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        SiProgramacionAgregar = True
    Else
        oConexion.RollbackTrans
        SiProgramacionAgregar = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function SiProgramacionModificar(oDOProgramacionMedica As DOSiProgramacion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgramacionMedica As New SiProgramacion
    
    SiProgramacionModificar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
      
    oConexion.BeginTrans
    Set oProgramacionMedica.Conexion = oConexion
    If oProgramacionMedica.modificar(oDOProgramacionMedica) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOProgramacionMedica.IdUsuarioAuditoria, "M", oDOProgramacionMedica.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
        SiProgramacionModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oProgramacionMedica.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oProgramacionMedica = Nothing
    
End Function

'MODIFICADO FRANKLIN CACHAY 10/10/2013 - se agrego oConexion.CommandTimeout = 300
Function SiProgramacionEliminar(oDoProgramacion As DOSiProgramacion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New SiProgramacion
Dim bEstadoTransaccion As Boolean
    
    SiProgramacionEliminar = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgMedica.Eliminar(oDoProgramacion) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoProgramacion.IdUsuarioAuditoria, "E", oDoProgramacion.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
        bEstadoTransaccion = True
    Else
        bEstadoTransaccion = False
    End If
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        SiProgramacionEliminar = True
    Else
        oConexion.RollbackTrans
        SiProgramacionEliminar = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

Function SiProgramacionMedicaSeleccionarPorResponsableYMes(ByVal lIdResponsable As Long, iMes As Integer, _
                                                           iAnio As Integer, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SiProgramacionMedicaSeleccionarPorResponsableYMes = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "SiProgramacionMedicaSeleccionarPorResponsableYMes"
        Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, lIdResponsable): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, iMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, iAnio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SiProgramacionMedicaSeleccionarPorResponsableYMes = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SiProgramacionMedicaLeerPorResponsableYMes(IdResponsable As Long, iMes As Integer, iAnio As Integer) As Collection
Dim oConexion As New ADODB.Connection
Dim oProgramacionMedica As New SiProgramacion
Dim rsProgramaciones As ADODB.Recordset
Dim oProgramaciones As New Collection
Dim oProgramacion As DOSiProgramacion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oProgramacionMedica.Conexion = oConexion
    Set rsProgramaciones = SiProgramacionMedicaSeleccionarPorResponsableYMes(IdResponsable, iMes, iAnio, oConexion)
    oConexion.Close
    ms_MensajeError = oProgramacionMedica.MensajeError
   
    Do While Not rsProgramaciones.EOF
       Set oProgramacion = New DOSiProgramacion
       oProgramacion.idProgramacion = IIf(IsNull(rsProgramaciones!idProgramacion), 0, rsProgramaciones!idProgramacion)
       oProgramacion.IdResponsable = IIf(IsNull(rsProgramaciones!IdResponsable), 0, rsProgramaciones!IdResponsable)
       oProgramacion.fecha = IIf(IsNull(rsProgramaciones!fecha), "", rsProgramaciones!fecha)
       oProgramacion.HoraInicio = IIf(IsNull(rsProgramaciones!HoraInicio), "", rsProgramaciones!HoraInicio)
       oProgramacion.horaFin = IIf(IsNull(rsProgramaciones!horaFin), "", rsProgramaciones!horaFin)
       oProgramacion.IdTurno = IIf(IsNull(rsProgramaciones!IdTurno), 0, rsProgramaciones!IdTurno)
       oProgramacion.IdSala = IIf(IsNull(rsProgramaciones!IdSala), 0, rsProgramaciones!IdSala)
       oProgramaciones.Add oProgramacion
       rsProgramaciones.MoveNext
    Loop
    
    Set SiProgramacionMedicaLeerPorResponsableYMes = oProgramaciones
    
    
End Function

Function SiProgramacionXmedicoFechaHOra(lnIdResponsable As Long, lcFecha As String, lcHora As String) As Recordset


'    Dim lcSql As String
    Dim oRs As New Recordset
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
            .CommandText = "SiProgramacionXmedicoFechaHOra"
            Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, lnIdResponsable): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Hora", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
            Set oRs = .Execute
            Set oRs.ActiveConnection = Nothing
    End With
        
    Set SiProgramacionXmedicoFechaHOra = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


Function SiProgramacionEliminarVarios(oDoProgramacion As DOSiProgramacion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcMedicoServicio As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oProgMedica As New SiProgramacion
Dim bEstadoTransaccion As Boolean
    
    SiProgramacionEliminarVarios = False
    bEstadoTransaccion = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oProgMedica.Conexion = oConexion
    
    If oProgMedica.Eliminar(oDoProgramacion) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoProgramacion.IdUsuarioAuditoria, "E", oDoProgramacion.idProgramacion, "ProgramacionMedica", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcMedicoServicio)        'ListBarItems.idListItem
        bEstadoTransaccion = True
    Else
        bEstadoTransaccion = False
    End If
    
    If bEstadoTransaccion Then
        oConexion.CommitTrans
        SiProgramacionEliminarVarios = True
    Else
        oConexion.RollbackTrans
        SiProgramacionEliminarVarios = False
        ms_MensajeError = oProgMedica.MensajeError
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oProgMedica = Nothing
End Function

Function SiProgramacionXsalaFecha(lnIdSala As Long, lcFecha As Date) As Recordset


'    Dim lcSql As String
    Dim oRs As New Recordset
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
            .CommandText = "SiProgramacionXsalaFecha"
            Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
            Set oRs = .Execute
            Set oRs.ActiveConnection = Nothing
    End With
        
    Set SiProgramacionXsalaFecha = oRs
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function



Sub SiProgramacionLlenaComboCuposMaximo(oCtlResponsable As Object, ml_idSala As Long, ml_FechaCita As Date, _
                                        ByRef ml_NroCuposMaximosXdia As Long, lbMuestraMensaje As Boolean, _
                                        lnOpcion As sghSiCuposProgamados, ByRef oRsProgramacion As Recordset)
    Dim lnCuposMaximos As Long, lnCuposQueQuedan As Long
    Set oRsProgramacion = SiProgramacionXsalaFecha(ml_idSala, ml_FechaCita)
    ml_NroCuposMaximosXdia = 0
    If oRsProgramacion.RecordCount = 0 Then
       If lbMuestraMensaje = True Then
          MsgBox "No hay programación para ese día", vbInformation, ""
       End If
       If lnOpcion = sghSiProgramacionEnDataCombo Or lnOpcion = sghSiProgramacionEnListBox Then
          oCtlResponsable.Clear
       End If
    Else
       If lnOpcion = sghSiProgramacionEnDataCombo Or lnOpcion = sghSiProgramacionEnListBox Then
          oCtlResponsable.Clear
       End If
       oRsProgramacion.MoveFirst
       Do While Not oRsProgramacion.EOF
          lnCuposMaximos = DateDiff("s", CDate(Date & " " & oRsProgramacion!HoraInicio), CDate(Date & " " & oRsProgramacion!horaFin)) / 60
          lnCuposMaximos = Round(lnCuposMaximos / oRsProgramacion!TiempoPromedioAtencion, 0)
          lnCuposQueQuedan = lnCuposMaximos - oRsProgramacion!cuposLlenos
          ml_NroCuposMaximosXdia = ml_NroCuposMaximosXdia + lnCuposMaximos
          If lnOpcion = sghSiProgramacionEnDataCombo Or lnOpcion = sghSiProgramacionEnListBox Then
             oCtlResponsable.AddItem Trim(oRsProgramacion!responsable) & " (" & oRsProgramacion!HoraInicio & " - " & _
                                 oRsProgramacion!horaFin & ") (Le quedan: " & Trim(Str(lnCuposQueQuedan)) & " Cupos)"
          End If
          oRsProgramacion.MoveNext
       Loop
       If lnOpcion = sghSiProgramacionEnDataCombo Then
            If oRsProgramacion.RecordCount = 1 Then
               oCtlResponsable.ListIndex = 0
            End If
       End If
    End If

End Sub

Function siCitasXidprogramacion(lnIdProgrmacion As Long, Optional oConexion1 As Connection) As Recordset
    Dim oRs As New Recordset
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
            .CommandText = "siCitasXidprogramacion"
            Set oParameter = .CreateParameter("@IdProgrmacion", adInteger, adParamInput, 0, lnIdProgrmacion): .Parameters.Append oParameter
            Set oRs = .Execute
            If oConexion1 Is Nothing Then
               Set oRs.ActiveConnection = Nothing
            End If
    End With
        
    Set siCitasXidprogramacion = oRs
    If oConexion1 Is Nothing Then
        oConexion.Close
        Set oConexion = Nothing
    End If
    Set oCommand = Nothing
End Function
Function SiProgramacionXresponsableYfechaYsala(oDOProgramacionMedica As DOSiProgramacion) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set SiProgramacionXresponsableYfechaYsala = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SiProgramacionXresponsableYfechaYsala"
            Set oParameter = .CreateParameter("@idSala", adInteger, adParamInput, 0, oDOProgramacionMedica.IdSala): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, oDOProgramacionMedica.IdResponsable): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oDOProgramacionMedica.fecha = 0, Null, oDOProgramacionMedica.fecha)): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set SiProgramacionXresponsableYfechaYsala = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function


Function siProgramacionXsalaYfecha(lnIdSala As Long, ldFecha As Date) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set siProgramacionXsalaYfecha = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "siProgramacionXsalaYfecha"
            Set oParameter = .CreateParameter("@idSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, ldFecha): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set siProgramacionXsalaYfecha = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function SiCitasSeleccionarPorSalaYFecha(lnIdSala As Long, ldFecha As Date) As Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set SiCitasSeleccionarPorSalaYFecha = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SiCitasSeleccionarPorSalaYFecha"
            Set oParameter = .CreateParameter("@idSala", adInteger, adParamInput, 0, lnIdSala): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, ldFecha): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
        End With

        Set SiCitasSeleccionarPorSalaYFecha = oRsTmp1
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function


Function SiCitasDetalleActualizaLlave(lnIdCitaSI As Long, lcLlaveTicket As String, oConexion As Connection) As Boolean
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        SiCitasDetalleActualizaLlave = False
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SiCitasDetalleActualizaLlave"
            Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, lcLlaveTicket): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idCitaSi", adInteger, adParamInput, 0, lnIdCitaSI): .Parameters.Append oParameter
             .Execute
        End With
        SiCitasDetalleActualizaLlave = True
        Set oCommand = Nothing
End Function


