VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New sighdatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Const AnchoPagina = 120

Property Let MensajeError(sValue As String)
  ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
  MensajeError = ms_MensajeError
End Property

Function LabMovimientoSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  
  Set LabMovimientoSeleccionarPorFechasPuntoCarga = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSeleccionarPorFechasPuntoCarga"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoSeleccionarPorFechasPuntoCarga = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoLaboratorioAgregar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, oDOPAciente As doPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, IdCuentaAtencion As Long, rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoLaboratorio As New LabMovimientoLaboratorio
  Dim oPacientes As New Pacientes
  Dim oFactOrdenesServicio As New FactOrdenServicio
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
  Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
  Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
  Dim oFacturacionServicioPagos As New FacturacionServicioPagos
  Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
  Dim oFacturacionServicioFinan As New FacturacionServicioFin
  Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
  Dim oFactServicioPagos As New FactOrdenServicioPagos
  Dim oDoLabMovimientoCPT As New DoLabMovimientoCPT
  Dim oLabMovimientoCPT As New LabMovimientoCPT
  Dim oHistoria As New HistoriasClinicas
  Dim bProcesoOK As Boolean
  Dim rsPacientes As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoLaboratorio.Conexion = oConexion
  Set oPacientes.Conexion = oConexion
  Set oFactOrdenesServicio.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  Set oFacturacionServicioDespacho.Conexion = oConexion
  Set oFacturacionServicioPagos.Conexion = oConexion
  Set oFacturacionServicioFinan.Conexion = oConexion
  Set oFactServicioPagos.Conexion = oConexion
  Set oHistoria.Conexion = oConexion
  Set oLabMovimientoCPT.Conexion = oConexion
  
  If Not oLabMovimiento.Insertar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarILabA
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "A", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                  'ListBarItems.idListItem
  
  If IdCuentaAtencion > 0 Then
    'Paciente registrado ingresando Nro Cuenta
    If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
    oDoLabMovimientoLaboratorio.IdOrden = oDoFactOrdenServ.IdOrden
  Else
    'Paciente registrado ingresando Nro Boleta
    If oDOPAciente.idPaciente = 0 Then
      Set rsPacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDOPAciente)
      If rsPacientes.RecordCount > 0 Then
        oDOPAciente.idPaciente = rsPacientes.Fields!idPaciente
      Else
        oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
        If Not oPacientes.Insertar(oDOPAciente) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarILabA
      End If
    End If
    oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
    If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
    oDoFactOrdenServ.idPaciente = oDOPAciente.idPaciente
    oDoFactOrdenServ.IdPuntoCarga = oDOLabMovimiento.IdPuntoCarga
    If Not oFactOrdenesServicio.Modificar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
  End If
  
  oDoLabMovimientoLaboratorio.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoLaboratorio.Insertar(oDoLabMovimientoLaboratorio) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoLaboratorio.MensajeError: GoTo TerminarILabA
  
  If IdCuentaAtencion > 0 Then
    'Paciente registrado ingresando Nro Cuenta
    If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento) = True Then
      'If oDoFactOrdenServ.IdTipoFinanciamiento = 1 Or oDoFactOrdenServ.IdTipoFinanciamiento = 5 Then
      '***********Contado ***********
      With oDOFactOrdenServicioPagos
        .FechaCreacion = oDoFactOrdenServ.FechaCreacion
        .IdEstadoFacturacion = 1   'Registrado
        .IdOrden = oDoFactOrdenServ.IdOrden
        .IdUsuario = oDoFactOrdenServ.IdUsuarioAuditoria
        .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
      End With
      If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFactServicioPagos.MensajeError: GoTo TerminarILabA
    End If
  End If
  
  'Agregar Productos (insumos)
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  If Not (rsFacturacionProductos.EOF = True And rsFacturacionProductos.BOF = True) Then
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
      oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
      oDoLabMovimientoDetalle.idProductoCPT = rsFacturacionProductos.Fields!idProductoCPT
      oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
      oDoLabMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
      If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarILabA
      rsFacturacionProductos.MoveNext
    Loop
  End If
  
  'Agregar Productos (cpt)
  oDoLabMovimientoCPT.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  rsFacturacionProductosCPT.MoveFirst
  Do While Not rsFacturacionProductosCPT.EOF
    oDoLabMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
    oDoLabMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!IdProducto
    oDoLabMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
    oDoLabMovimientoCPT.importe = rsFacturacionProductosCPT.Fields!totalPorPagar
    If Not oLabMovimientoCPT.Insertar(oDoLabMovimientoCPT) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoCPT.MensajeError: GoTo TerminarILabA
    
    If IdCuentaAtencion > 0 Then
      'Paciente registrado ingresando Nro Cuenta
      oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
      oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
      oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
      oDoFacturacionServicioDesp.IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
      oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
      oDoFacturacionServicioDesp.total = rsFacturacionProductosCPT.Fields!totalPorPagar
      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarILabA
      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento) = True Then
        '***********Contado ***********
        With oDoFacturacionServicioPagos
          .cantidad = rsFacturacionProductosCPT.Fields!cantidad
          .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
          .IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
          .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
          .Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
          .total = rsFacturacionProductosCPT.Fields!totalPorPagar
        End With
        If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:                    GoTo TerminarILabA
      Else
        '***********Sis,Soat,exoneracion, convenio ***********
        With oDoFacturacionServicioFinanc
          .CantidadFinanciada = rsFacturacionProductosCPT.Fields!cantidad
          .FechaAutoriza = oDoFactOrdenServ.FechaCreacion
          .IdOrden = oDoFactOrdenServ.IdOrden
          .IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
          .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
          .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
          .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
          .PrecioFinanciado = rsFacturacionProductosCPT.Fields!PrecioUnitario
          .TotalFinanciado = rsFacturacionProductosCPT.Fields!totalPorPagar
          .IdFuenteFinanciamiento = oDoFactOrdenServ.IdFuenteFinanciamiento
        End With
        If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarILabA
      End If
    End If
    rsFacturacionProductosCPT.MoveNext
  Loop
  
TerminarILabA:
  If bProcesoOK Then
    LabMovimientoLaboratorioAgregar = True
    oConexion.CommitTrans
  Else
    LabMovimientoLaboratorioAgregar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoLaboratorio = Nothing
  Set oPacientes = Nothing
  Set oFactOrdenesServicio = Nothing
  Set oLabMovimientoDetalle = Nothing
  Set oDoLabMovimientoDetalle = Nothing
  Set oDoFacturacionServicioDesp = Nothing
  Set oFacturacionServicioDespacho = Nothing
  Set oDoFacturacionServicioPagos = Nothing
  Set oFacturacionServicioPagos = Nothing
  Set oDoFacturacionServicioFinanc = Nothing
  Set oFacturacionServicioFinan = Nothing
  Set oFactServicioPagos = Nothing
  Set mo_AdminAdmision = Nothing
  Set oDoLabMovimientoCPT = Nothing
  Set oLabMovimientoCPT = Nothing
  Set oHistoria = Nothing
End Function

Function LabMovimientoLaboratorioModificar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, oDOPAciente As doPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, IdCuentaAtencion As Long, rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoLaboratorio As New LabMovimientoLaboratorio
  Dim oPacientes As New Pacientes
  Dim oFactOrdenesServicio As New FactOrdenServicio
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
  Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
  Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
  Dim oFacturacionServicioPagos As New FacturacionServicioPagos
  Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
  Dim oFacturacionServicioFinan As New FacturacionServicioFin
  Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
  Dim oFactServicioPagos As New FactOrdenServicioPagos
  Dim oDoLabMovimientoCPT As New DoLabMovimientoCPT
  Dim oLabMovimientoCPT As New LabMovimientoCPT
  Dim oHistoria As New HistoriasClinicas
  Dim bProcesoOK As Boolean
  Dim oRsTmp1 As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoLaboratorio.Conexion = oConexion
  Set oPacientes.Conexion = oConexion
  Set oFactOrdenesServicio.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  Set oFacturacionServicioDespacho.Conexion = oConexion
  Set oFacturacionServicioPagos.Conexion = oConexion
  Set oFacturacionServicioFinan.Conexion = oConexion
  Set oFactServicioPagos.Conexion = oConexion
  Set oHistoria.Conexion = oConexion
  Set oLabMovimientoCPT.Conexion = oConexion
  
  If IdCuentaAtencion = 0 Then
    If oDOPAciente.idPaciente = 0 Then
      oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
      If Not oPacientes.Insertar(oDOPAciente) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
    Else
      If Not oPacientes.Modificar(oDOPAciente, True) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
    End If
  End If
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
  If Not oLabMovimientoLaboratorio.Modificar(oDoLabMovimientoLaboratorio) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoLaboratorio.MensajeError: GoTo TerminarIMMAGA
  oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
  If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
  oDoFactOrdenServ.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
  If oDOPAciente.idPaciente > 0 And oDoFactOrdenServ.idPaciente <> oDOPAciente.idPaciente Then
    oDoFactOrdenServ.idPaciente = oDOPAciente.idPaciente
    If Not oFactOrdenesServicio.Modificar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
  End If
   
  'Elimina Productos
  If IdCuentaAtencion > 0 Then
    Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden)
    If oRsTmp1.RecordCount > 0 Then
      If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
        oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
        oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
        oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
        If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIMMAGA
      End If
    End If
    oDoFacturacionServicioFinanc.IdOrden = oDoFactOrdenServ.IdOrden
    oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
    If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIMMAGA
    If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIMMAGA
  End If
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoDetalle.Eliminar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarIMMAGA
  oDoLabMovimientoCPT.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoCPT.Eliminar(oDoLabMovimientoCPT) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoCPT.MensajeError: GoTo TerminarIMMAGA
  
  'Agregar Productos (insumos)
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  If Not (rsFacturacionProductos.EOF = True And rsFacturacionProductos.BOF = True) Then
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
      oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
      oDoLabMovimientoDetalle.idProductoCPT = rsFacturacionProductos.Fields!idProductoCPT
      oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
      oDoLabMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
      If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarIMMAGA
      rsFacturacionProductos.MoveNext
    Loop
  End If
  
  'Agregar Productos (cpt)
  oDoLabMovimientoCPT.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  rsFacturacionProductosCPT.MoveFirst
  Do While Not rsFacturacionProductosCPT.EOF
    oDoLabMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
    oDoLabMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!IdProducto
    oDoLabMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
    oDoLabMovimientoCPT.importe = rsFacturacionProductosCPT.Fields!totalPorPagar
    If Not oLabMovimientoCPT.Insertar(oDoLabMovimientoCPT) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoCPT.MensajeError: GoTo TerminarIMMAGA
    
    If IdCuentaAtencion > 0 Then
      'Paciente registrado ingresando Nro Cuenta
      oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
      oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
      oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
      oDoFacturacionServicioDesp.IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
      oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
      oDoFacturacionServicioDesp.total = rsFacturacionProductosCPT.Fields!totalPorPagar
      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIMMAGA
      If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento) = True Then
        '***********Contado ***********
        With oDoFacturacionServicioPagos
          .cantidad = rsFacturacionProductosCPT.Fields!cantidad
          .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
          .IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
          .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
          .Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
          .total = rsFacturacionProductosCPT.Fields!totalPorPagar
        End With
        If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:                    GoTo TerminarIMMAGA
      Else
        '***********Sis,Soat,exoneracion, convenio ***********
        With oDoFacturacionServicioFinanc
          .CantidadFinanciada = rsFacturacionProductosCPT.Fields!cantidad
          .FechaAutoriza = oDoFactOrdenServ.FechaCreacion
          .IdOrden = oDoFactOrdenServ.IdOrden
          .IdProducto = rsFacturacionProductosCPT.Fields!IdProducto
          .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
          .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
          .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
          .PrecioFinanciado = rsFacturacionProductosCPT.Fields!PrecioUnitario
          .TotalFinanciado = rsFacturacionProductosCPT.Fields!totalPorPagar
          .IdFuenteFinanciamiento = oDoFactOrdenServ.IdFuenteFinanciamiento
        End With
        If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIMMAGA
      End If
    End If
    rsFacturacionProductosCPT.MoveNext
  Loop
  
TerminarIMMAGA:
  If bProcesoOK Then
    LabMovimientoLaboratorioModificar = True
    oConexion.CommitTrans
  Else
    LabMovimientoLaboratorioModificar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoLaboratorio = Nothing
  Set oPacientes = Nothing
  Set oFactOrdenesServicio = Nothing
  Set oLabMovimientoDetalle = Nothing
  Set oDoLabMovimientoDetalle = Nothing
  Set oDoFacturacionServicioDesp = Nothing
  Set oFacturacionServicioDespacho = Nothing
  Set oDoFacturacionServicioPagos = Nothing
  Set oFacturacionServicioPagos = Nothing
  Set oDoFacturacionServicioFinanc = Nothing
  Set oFacturacionServicioFinan = Nothing
  Set oFactServicioPagos = Nothing
  Set mo_AdminAdmision = Nothing
  Set oDoLabMovimientoCPT = Nothing
  Set oLabMovimientoCPT = Nothing
  Set oHistoria = Nothing
End Function

Function LabMovimientoLaboratorioAnular(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, oDOPAciente As doPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, IdCuentaAtencion As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoLaboratorio As New LabMovimientoLaboratorio
  Dim oPacientes As New Pacientes
  Dim oFactOrdenesServicio As New FactOrdenServicio
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
  Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
  Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
  Dim oFacturacionServicioPagos As New FacturacionServicioPagos
  Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
  Dim oFacturacionServicioFinan As New FacturacionServicioFin
  Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
  Dim oFactServicioPagos As New FactOrdenServicioPagos
  Dim oHistoria As New HistoriasClinicas
  Dim bProcesoOK As Boolean
  Dim oRsTmp1 As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoLaboratorio.Conexion = oConexion
  Set oPacientes.Conexion = oConexion
  Set oFactOrdenesServicio.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  Set oFacturacionServicioDespacho.Conexion = oConexion
  Set oFacturacionServicioPagos.Conexion = oConexion
  Set oFacturacionServicioFinan.Conexion = oConexion
  Set oFactServicioPagos.Conexion = oConexion
  Set oHistoria.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0   'Anulado
  If Not oLabMovimiento.Modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarIAMAGA
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "E", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  'Anula Cabecera Productos
  If IdCuentaAtencion > 0 Then
    oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
    If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
    oDoFactOrdenServ.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
    oDoFactOrdenServ.IdEstadoFacturacion = 9  'anulado
    If Not oFactOrdenesServicio.Modificar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
    Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden)
    If oRsTmp1.RecordCount > 0 Then
      If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
        oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
        If Not oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
        oDOFactOrdenServicioPagos.IdEstadoFacturacion = 9 'anulado
        oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
        If Not oFactServicioPagos.Modificar(oDOFactOrdenServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
      End If
    End If
  End If
  
TerminarIAMAGA:
  If bProcesoOK Then
    LabMovimientoLaboratorioAnular = True
    oConexion.CommitTrans
  Else
    LabMovimientoLaboratorioAnular = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoLaboratorio = Nothing
  Set oPacientes = Nothing
  Set oFactOrdenesServicio = Nothing
  Set oLabMovimientoDetalle = Nothing
  Set oDoLabMovimientoDetalle = Nothing
  Set oDoFacturacionServicioDesp = Nothing
  Set oFacturacionServicioDespacho = Nothing
  Set oDoFacturacionServicioPagos = Nothing
  Set oFacturacionServicioPagos = Nothing
  Set oDoFacturacionServicioFinanc = Nothing
  Set oFacturacionServicioFinan = Nothing
  Set oFactServicioPagos = Nothing
  Set mo_AdminAdmision = Nothing
  Set oHistoria = Nothing
End Function

Function LabMovimientoSeleccionarPorId(IdMovimiento As Long) As DOLabMovimiento
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oDOLabMovimiento As New DOLabMovimiento
  Set LabMovimientoSeleccionarPorId = Nothing
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set oLabMovimiento.Conexion = oConexion
  oDOLabMovimiento.IdMovimiento = IdMovimiento
  If oLabMovimiento.SeleccionarPorId(oDOLabMovimiento) Then
    Set LabMovimientoSeleccionarPorId = oDOLabMovimiento
  Else
    ms_MensajeError = oLabMovimiento.MensajeError
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oDOLabMovimiento = Nothing
End Function

Function LabMovimientoLaboratorioSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorIdMovimiento"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorIdMovimiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoDetalleSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoDetalleSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoDetalleSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoDetalleSeleccionarPorIdMovimiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoCPTSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoCPTSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoCPTSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoCPTSeleccionarPorIdMovimiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoLaboratorioSeleccionarPorIdComprobantePago(LnIdComprobantePago As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorIdComprobantePago = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorIdComprobantePago"
    Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, LnIdComprobantePago): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorIdComprobantePago = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoIngresosSeleccionarPorId(IdMovimiento As Long) As DOLabMovimientoIngresos
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimientoIngresos As New LabMovimientoIngresos
  Dim oDoLabMovimientoIngresos As New DOLabMovimientoIngresos
  Set LabMovimientoIngresosSeleccionarPorId = Nothing
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set oLabMovimientoIngresos.Conexion = oConexion
  oDoLabMovimientoIngresos.IdMovimiento = IdMovimiento
  If oLabMovimientoIngresos.SeleccionarPorId(oDoLabMovimientoIngresos) Then
    Set LabMovimientoIngresosSeleccionarPorId = oDoLabMovimientoIngresos
  Else
    ms_MensajeError = oLabMovimientoIngresos.MensajeError
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimientoIngresos = Nothing
  Set oDoLabMovimientoIngresos = Nothing
End Function

Function LabMovimientoSalidasSeleccionarPorId(IdMovimiento As Long) As DOLabMovimientoSalidas
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimientoSalidas As New LabMovimientoSalidas
  Dim oDoLabMovimientoSalidas As New DOLabMovimientoSalidas
  Set LabMovimientoSalidasSeleccionarPorId = Nothing
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set oLabMovimientoSalidas.Conexion = oConexion
  oDoLabMovimientoSalidas.IdMovimiento = IdMovimiento
  If oLabMovimientoSalidas.SeleccionarPorId(oDoLabMovimientoSalidas) Then
    Set LabMovimientoSalidasSeleccionarPorId = oDoLabMovimientoSalidas
  Else
    ms_MensajeError = oLabMovimientoSalidas.MensajeError
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimientoSalidas = Nothing
  Set oDoLabMovimientoSalidas = Nothing
End Function

Function LabMovimientoSeleccionarPorFechasPuntoCargaIngresos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSeleccionarPorFechasPuntoCargaIngresos = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSeleccionarPorFechasPuntoCargaIngresos"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoSeleccionarPorFechasPuntoCargaIngresos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoSeleccionarPorFechasPuntoCargaSalidas(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSeleccionarPorFechasPuntoCargaSalidas = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSeleccionarPorFechasPuntoCargaSalidas"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@lnIdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoSeleccionarPorFechasPuntoCargaSalidas = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoIngresosAgregar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoIngresos As DOLabMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoIngresos As New LabMovimientoIngresos
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoIngresos.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimiento.Insertar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabIng
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "A", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
  oDoLabMovimientoIngresos.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoIngresos.Insertar(oDoLabMovimientoIngresos) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoIngresos.MensajeError: GoTo TerminarLabIng
  
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
    If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabIng
    rsFacturacionProductos.MoveNext
  Loop
   
TerminarLabIng:
  If bProcesoOK Then
    LabMovimientoIngresosAgregar = True
    oConexion.CommitTrans
  Else
    LabMovimientoIngresosAgregar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoIngresos = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMovimientoIngresosModificar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoIngresos As DOLabMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoIngresos As New LabMovimientoIngresos
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoIngresos.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimientoIngresos.Modificar(oDoLabMovimientoIngresos) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoIngresos.MensajeError: GoTo TerminarLabMod
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
   
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  If Not oLabMovimientoDetalle.Eliminar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabMod
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
    If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabMod
    rsFacturacionProductos.MoveNext
  Loop
  
TerminarLabMod:
  If bProcesoOK Then
    LabMovimientoIngresosModificar = True
    oConexion.CommitTrans
  Else
    LabMovimientoIngresosModificar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoIngresos = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMovimientoIngresosAnular(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoIngresos As DOLabMovimientoIngresos, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoIngresos As New LabMovimientoIngresos
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoIngresos.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0 'Anulado
  If Not oLabMovimiento.Modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabAnu
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "E", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
TerminarLabAnu:
  If bProcesoOK Then
    LabMovimientoIngresosAnular = True
    oConexion.CommitTrans
  Else
    LabMovimientoIngresosAnular = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoIngresos = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMovimientoSalidasAgregar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoSalidas As DOLabMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoSalidas As New LabMovimientoSalidas
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoSalidas.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimiento.Insertar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabSalIng
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "A", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")               'ListBarItems.idListItem
  
  oDoLabMovimientoSalidas.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoSalidas.Insertar(oDoLabMovimientoSalidas) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoSalidas.MensajeError: GoTo TerminarLabSalIng
  
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
    If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabSalIng
    rsFacturacionProductos.MoveNext
  Loop
  
TerminarLabSalIng:
  If bProcesoOK Then
    LabMovimientoSalidasAgregar = True
    oConexion.CommitTrans
  Else
    LabMovimientoSalidasAgregar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoSalidas = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMovimientoSalidasModificar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoSalidas As DOLabMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoSalidas As New LabMovimientoSalidas
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoSalidas.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimientoSalidas.Modificar(oDoLabMovimientoSalidas) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoSalidas.MensajeError: GoTo TerminarLabSalMod
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  If Not oLabMovimientoDetalle.Eliminar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabSalMod
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.IdProducto = rsFacturacionProductos.Fields!IdProducto
    If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabSalMod
    rsFacturacionProductos.MoveNext
  Loop
  
TerminarLabSalMod:
  If bProcesoOK Then
    LabMovimientoSalidasModificar = True
    oConexion.CommitTrans
  Else
    LabMovimientoSalidasModificar = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoSalidas = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMovimientoSalidasAnular(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoSalidas As DOLabMovimientoSalidas, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oLabMovimientoSalidas As New LabMovimientoSalidas
  Dim oLabMovimientoDetalle As New LabMovimientoDetalle
  Dim oDoLabMovimientoDetalle As New DoLabMovimientoDetalle
  Dim bProcesoOK As Boolean
  
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoSalidas.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0   'Anulado
  If Not oLabMovimiento.Modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabSalAnul
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "E", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem

TerminarLabSalAnul:
  If bProcesoOK Then
    LabMovimientoSalidasAnular = True
    oConexion.CommitTrans
  Else
    LabMovimientoSalidasAnular = False
    ms_MensajeError = ms_MensajeError & " - " & Err.Description
    oConexion.RollbackTrans
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oLabMovimiento = Nothing
  Set oLabMovimientoSalidas = Nothing
  Set oLabMovimientoDetalle = Nothing
End Function

Function LabMotivoSalidasSeleccionarTodos() As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim ms_MensajeError  As String
  oRecordset.Open "select * from LabMotivoSalidas order by motivo", SIGHComun.CadenaConexion, adOpenKeyset, adLockOptimistic
  Set LabMotivoSalidasSeleccionarTodos = oRecordset
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabRecojeExamenSeleccionarTodos() As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim ms_MensajeError  As String
  oRecordset.Open "select * from LabRecojeExamen order by RecojeExamen", SIGHComun.CadenaConexion, adOpenKeyset, adLockOptimistic
  Set LabRecojeExamenSeleccionarTodos = oRecordset
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoCPTSeleccionarPorFechas(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
    
  Set LabMovimientoCPTSeleccionarPorFechas = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoCPTSeleccionarPorFechas"
    Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoCPTSeleccionarPorFechas = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LaboratorioPruebasPorEmpleado(ldFechaIni As Date, ldFechaFin As Date, idEmpleado As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
    
  Set LaboratorioPruebasPorEmpleado = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LaboratorioPruebasPorEmpleado"
    Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, idEmpleado): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LaboratorioPruebasPorEmpleado = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabFactCatalogoServiciosHospFiltraPorPuntoCarga(lcFiltro As String, lnIdPuntoCarga As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabFactCatalogoServiciosHospFiltraPorPuntoCarga = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabFactCatalogoServiciosHospFiltraPorPuntoCarga"
    Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Filtro ", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabFactCatalogoServiciosHospFiltraPorPuntoCarga = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnIdProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
'    MsgBox 1
    .CommandText = "LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter    '3->Por Codigo     4->Por fecha
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoIngresosSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoIngresosSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoIngresosSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoIngresosSeleccionarPorIdMovimiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoSalidasSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSalidasSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSalidasSeleccionarPorIdMovimiento"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoSalidasSeleccionarPorIdMovimiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabRecojeExamenSeleccionaPorID(lnIdRecojeExamen As Long) As String
  Dim oRsTmp As New Recordset
  oRsTmp.Open "select * from LabRecojeExamen where idRecojeExamen=" & lnIdRecojeExamen, SIGHComun.CadenaConexion, adOpenKeyset, adLockOptimistic
  If oRsTmp.RecordCount > 0 Then
    LabRecojeExamenSeleccionaPorID = oRsTmp.Fields!RecojeExamen
  Else
    LabRecojeExamenSeleccionaPorID = ""
  End If
  Set oRsTmp = Nothing
End Function


Function LabMovimientoLaboratorioSeleccionarPorFechasPToInsumos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorFechasPToInsumos = Nothing
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasPtoInsumos"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter    '3->Por Codigo     4->Por fecha
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorFechasPToInsumos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Public Function LabRecuperaResultados_Res(CAnalisis As String, Orden As Double) As String
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  LabRecuperaResultados_Res = ""
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idorden='" & Orden & "'"
    Set oTabla = New ADODB.Recordset
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      If oTabla!ResultadoAnalisis <> "" Then LabRecuperaResultados_Res = oTabla!ResultadoAnalisis
    End If
    oTabla.Close
    oConexSQL.Close
  End If
End Function

Public Function LabRecuperaResultados_ReaP(CAnalisis As String, Orden As Double) As Long
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idorden='" & Orden & "'"
    Set oTabla = New ADODB.Recordset
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      If oTabla!ResultadoAnalisis <> "" Then LabRecuperaResultados_ReaP = oTabla!realizaAnalisis
    End If
    oTabla.Close
    oConexSQL.Close
  End If
End Function

Public Sub LabImprimeCabeceraResultados(CodPac As String, NomPac As String, HisPac As String, Fec_ As String, NomMed As String)
  Dim Mensaje, TempY
  If Fec_ = "" Then Fec_ = Format(Now, "dd/mm/yyyy hh:mm:ss")
  If HisPac = "" Then HisPac = "-"
  Printer.Font = "Verdana"
  Printer.FontSize = 8
  'Printer.PrintQuality = -1
  'Printer.Orientation = vbPRORPortrait 'vbPRORLandscape
  Printer.ScaleMode = 7 'vbCentimeters 'Medida en Centimetros
  Printer.CurrentY = 1.2
  Mensaje = Format(Fec_, "dd/mmmm/yyyy")
  Mensaje = Mensaje & ", " & Format(Fec_, "hh:mm:ss AMPM")
  TempY = Printer.CurrentY
  Printer.FontBold = True
  Printer.Print Tab(8); "Id Paciente:"; Tab(65); "Nombres:"
  Printer.FontBold = False
  Printer.CurrentY = TempY
  Printer.Print Tab(25); CodPac; Tab(85); NomPac
  TempY = Printer.CurrentY
  Printer.FontBold = True
  Printer.Print Tab(8); "Nº Historia:"; Tab(68); "Fecha:"
  Printer.FontBold = False
  Printer.CurrentY = TempY
  Printer.Print Tab(25); HisPac; Tab(85); Mensaje
  
  TempY = Printer.CurrentY
  Printer.FontBold = True
  Printer.Print Tab(12); "Médico:"
  Printer.FontBold = False
  Printer.CurrentY = TempY
  If NomMed = "Ninguno" Or NomMed = "" Then NomMed = "-"
  Printer.Print Tab(25); NomMed
  Printer.CurrentY = Printer.CurrentY + 0.1 'Printer.Print
  Printer.FontSize = 12
  Printer.FontUnderline = True
  Printer.FontBold = True
  TempY = Printer.CurrentY
  Mensaje = "RESULTADOS"
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2 'Obtiene la mitad del ancho
  Printer.Print Mensaje
  Printer.FontBold = False
  Printer.FontUnderline = False
  Printer.FontSize = 8
  Printer.Print Tab(5); "================================================================================="
End Sub

Public Sub ImprimeLínea()
  Printer.FontSize = 8
  Printer.Print Tab(5); "================================================================================="
End Sub

Public Sub ImprimeVariasLineas(Inicio As Single, Mens As String, AnchoPagina As Integer)
  Dim TempMens As String
  Do While Mens <> ""
    Mens = Trim(Mens)
    If Len(Mens) > (AnchoPagina) Then
      TempMens = Left(Mens, AnchoPagina)
      Mens = Right(Mens, Len(Mens) - AnchoPagina)
    Else
      TempMens = Mens
      Mens = ""
    End If
    Printer.Print Tab(Inicio); TempMens
  Loop
End Sub

Public Sub LabImprimePieResultados()
  Dim Mensaje
  Printer.CurrentY = Printer.CurrentY + 0.1
  Mensaje = "Atentamente,"
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.EndDoc
End Sub

Public Sub LabImprimeObservaciones(Obs As String)
  Dim TempY
  Printer.CurrentY = Printer.CurrentY + 0.05
  Printer.FontUnderline = True
  Printer.FontBold = True
  TempY = Printer.CurrentY
  Printer.Print Tab(7); "OBSERVACIONES:"
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.CurrentY = TempY
  ImprimeVariasLineas 28, Obs, AnchoPagina
End Sub

Public Sub LabImprimeRealizaPrueba(ReaP As String)
  Dim TempY
  Printer.CurrentY = Printer.CurrentY + 0.05
  'Printer.FontUnderline = True
  Printer.FontBold = True
  TempY = Printer.CurrentY
  Printer.Print Tab(17); "PERSONAL QUE REALIZÓ PRUEBA  ------> "
  'Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.CurrentY = TempY
  ImprimeVariasLineas 66, ReaP, AnchoPagina
End Sub

Public Function LabImprimeResultadosBQM(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long)     'Bioquímica
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
    
  If Not (ml_idPrueba = "BQM001") Then
    Printer.FontSize = 9
    Mensaje = "ANÁLISIS: " & UCase(NomPru)
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  End If
  If ml_idPrueba = "BQM001" Then 'Glucosa
    'BQM001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    
    Printer.FontSize = 9
    Mensaje = "ANÁLISIS: " & UCase(NomPru) & " ( " & UCase(Res01) & " )"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res02 & " " & Res03; Tab(73); Res04 & " " & Res03; Tab(118); Res05
  ElseIf ml_idPrueba = "BQM003" Then  'Colesterol total
    'BQM003
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM004" Then 'Colesterol fraccionado
    'BQM004
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(13); "DOSAJE DE"; Tab(46); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(99); "MÉTODO"
    Printer.FontBold = False
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Colesterol HDL"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res01 & " " & Res02; Tab(80); Res03 & " " & Res02; Tab(110); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Colesterol LDL"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res05 & " " & Res06; Tab(80); Res07 & " " & Res06; Tab(110); Res08
  ElseIf ml_idPrueba = "BQM005" Then 'Triglicéridos
    'BQM005
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM006" Then 'transaminasa GOT/GPT
    'BQM006
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM007" Then 'transaminasa GOT/GPT
    'BQM007
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM008" Or ml_idPrueba = "BQM009" Then 'Bilirrubina Total
    'BQM008
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "1095" Then 'Billirubina total y fraccionada
    'BQM009
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(13); "DOSAJE DE"; Tab(46); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(100); "MÉTODO"
    Printer.FontBold = False
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Bilirrubina Total"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res01 & " " & Res02; Tab(80); Res03 & " " & Res02; Tab(110); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Bilirrubina Directa"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res05 & " " & Res06; Tab(80); Res07 & " " & Res06; Tab(110); Res08
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "3. Billirrubina Indirecta"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res09 & " " & Res10; Tab(80); Res11 & " " & Res10; Tab(110); Res12
  ElseIf ml_idPrueba = "BQM010" Then 'Fosfatasa alcalina
    'BQM010
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM011" Then 'Proteina total
    'BQM011
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM012" Then 'Proteina total y fraccionada
    'BQM012
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(13); "DOSAJE DE"; Tab(46); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(99); "MÉTODO"
    Printer.FontBold = False
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Proteína Total"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res01 & " " & Res02; Tab(80); Res03 & " " & Res02; Tab(110); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Albúmina"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res05 & " " & Res06; Tab(80); Res07 & " " & Res06; Tab(110); Res08
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "3. Globulinas"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res09 & " " & Res10; Tab(80); Res11 & " " & Res10; Tab(110); Res12
  ElseIf ml_idPrueba = "BQM014" Then 'Amilasa en orina
    'BQM014
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM015" Then 'Úrea
    'BQM015
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
  ElseIf ml_idPrueba = "BQM016" Then  'Creatinina
    'BQM016
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(37); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(69); Res03 & " --> " & Res04 & " " & Res02; Tab(118); Res05
  ElseIf ml_idPrueba = "BQM017" Then 'Ácido Úrico
    'BQM017
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(37); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(69); Res03 & " --> " & Res04 & " " & Res02; Tab(118); Res05
  ElseIf ml_idPrueba = "BQM018" Then 'ADA
    'BQM018
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(37); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(69); Res03 & " --> " & Res04 & " " & Res02; Tab(118); Res05
  ElseIf ml_idPrueba = "BQM019" Then  'Calcio
    'BQM019
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(37); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(69); Res03 & " --> " & Res04 & " " & Res02; Tab(118); Res05
  ElseIf ml_idPrueba = "BQM021" Then  'Depuración de Creatinina
    'BQM021
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(13); "DOSAJE DE"; Tab(46); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(99); "MÉTODO"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Creatinina en Suero"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res01 & " " & Res02; Tab(80); Res03 & " " & Res02; Tab(110); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Creatinina en Orina"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res05 & " " & Res06; Tab(80); Res07 & " " & Res06; Tab(110); Res08
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "3. Dep. Creatinina Endógena"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res09 & " " & Res10; Tab(80); Res11 & " " & Res10; Tab(110); Res12
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "4. Volumen de Orina"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res13 & " " & Res14; Tab(80); Res15 & " " & Res14; Tab(110); Res16
  ElseIf ml_idPrueba = "BQM030" Then 'transaminasa GOT/GPT
    'BQM030
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(13); "DOSAJE DE"; Tab(46); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(99); "MÉTODO"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Trasnsaminasa GPT"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res01 & " " & Res02; Tab(80); Res03 & " " & Res02; Tab(110); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Transaminasa GOT"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(52); Res05 & " " & Res06; Tab(80); Res07 & " " & Res06; Tab(110); Res08
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosCPA(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long) 'Citopatología
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String, Res30 As String
  Dim Res31 As String, Res32 As String, Res33 As String, Res34 As String, Res35 As String, Res36 As String, Res37 As String, Res38 As String, Res39 As String, Res40 As String, Res41 As String, Res42 As String, Res43 As String, Res44 As String, Res45 As String, Res46 As String, Res47 As String, Res48 As String, Res49 As String, Res50 As String, Res51 As String, Res52 As String, Res53 As String, Res54 As String, Res55 As String, Res56 As String, Res57 As String, Res58 As String, Res59 As String, Res60 As String
  Dim Combo(1 To 30) As String
  Dim I As Integer, J As Integer, k As Integer
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "CPA001" Then  'Papanicolau
    'CPA001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
    Res17 = SeparaCadena(Resu)
    Res18 = SeparaCadena(Resu)
    Res19 = SeparaCadena(Resu)
    Res20 = SeparaCadena(Resu)
    Res21 = SeparaCadena(Resu)
    Res22 = SeparaCadena(Resu)
    Res23 = SeparaCadena(Resu)
    Res24 = SeparaCadena(Resu)
    Res25 = SeparaCadena(Resu)
    Res26 = SeparaCadena(Resu)
    Res27 = SeparaCadena(Resu)
    Res28 = SeparaCadena(Resu)
    Res29 = SeparaCadena(Resu)
    Res30 = SeparaCadena(Resu)
    Res31 = SeparaCadena(Resu)
    Res32 = SeparaCadena(Resu)
    Res33 = SeparaCadena(Resu)
    Res34 = SeparaCadena(Resu)
    Res35 = SeparaCadena(Resu)
    Res36 = SeparaCadena(Resu)
    Res37 = SeparaCadena(Resu)
    Res38 = SeparaCadena(Resu)
    Res39 = SeparaCadena(Resu)
    Res40 = SeparaCadena(Resu)
    Res41 = SeparaCadena(Resu)
    Res42 = SeparaCadena(Resu)
    Res43 = SeparaCadena(Resu)
    Res44 = SeparaCadena(Resu)
    Res45 = SeparaCadena(Resu)
    Res46 = SeparaCadena(Resu)
    Res47 = SeparaCadena(Resu)
    Res48 = SeparaCadena(Resu)
    Res49 = SeparaCadena(Resu)
    Res50 = SeparaCadena(Resu)
    Res51 = SeparaCadena(Resu)
    Res52 = SeparaCadena(Resu)
    Res53 = SeparaCadena(Resu)
    Res54 = SeparaCadena(Resu)
    Res55 = SeparaCadena(Resu)
    
    Res60 = "TOMA DE MUESTRA"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Res60) / 2
    Printer.Print Res60
    Printer.FontUnderline = False
    Printer.FontBold = False
      
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "SERVICIO DE PROCEDENCIA:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(41); Res01
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "1. CITOLOGÍA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    If Res02 = "True" Then Printer.Print Tab(10); "Primera vez"
    If Res03 = "True" Then
      Printer.Print Tab(10); "Subsecuente"
      Printer.FontBold = False
      TempY = Printer.CurrentY
      Printer.Print Tab(12); "Diagnóstico último PAP:"
      Printer.CurrentY = TempY
      ImprimeVariasLineas 38, Res04, AnchoPagina
    End If
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "2. SITUACIÓN GINECO-OBSTÉTRICA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(10); "GP:"; Tab(50); "Fecha Última Regla:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(16); Res05; Tab(78); Res06
        
    If Res07 = "True" Then Printer.Print Tab(10); "Premenopausia"
    If Res08 = "True" Then Printer.Print Tab(10); "Postmenopausia"
    If Res09 = "True" Then Printer.Print Tab(10); "Uso de Hormonas"
    If Res10 = "True" Then Printer.Print Tab(10); "DIU"
    If Res11 = "True" Then Printer.Print Tab(10); "Histerectomía"
    If Res12 = "True" Then Printer.Print Tab(10); "Embarazo actual"
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "3. OBSERVACIONES:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    
    For I = 1 To 10
      Combo(I) = ""
    Next I
    If Res13 = "1" Then Combo(1) = "Aparentemente Sano"
    If Res14 = "1" Then Combo(2) = "Cervicitis"
    If Res15 = "1" Then Combo(3) = "Erosión"
    If Res16 = "1" Then Combo(4) = "Anormal"
    If Res17 = "1" Then Combo(5) = "Sangrado Anormal"
    If Res18 = "1" Then Combo(6) = "Prolapso"
    If Res19 = "1" Then Combo(7) = "Leucorrea"
    If Res20 = "1" Then Combo(8) = "Prolapso"
    For k = 1 To 10
      For I = 1 To 9
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 10
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      'Printer.FontBold = True
      'Printer.Print Tab(7); "MUESTRA"
      'Printer.FontBold = False
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(10); Combo(k * 5 + 1); Tab(36); Combo(k * 5 + 2); Tab(62); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(115); Combo(k * 5 + 5)
    Next k
       
    If Res21 <> "" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Comentarios:"
      Printer.CurrentY = TempY
      ImprimeVariasLineas 25, Res21, AnchoPagina
    End If
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    TempY = Printer.CurrentY
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "4. RESPONSABLE DE TOMA DE PAP:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(48); Res22
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    Res60 = "RESULTADOS"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Res60) / 2
    Printer.Print Res60
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "1. CARACTERÍSTICAS DE LA MUESTRA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    If Res23 = "True" Or Res24 = "True" Or Res25 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Característica:"
      Printer.CurrentY = TempY
      If Res23 = "True" Then Printer.Print Tab(25); "Adecuada"
      If Res24 = "True" Then Printer.Print Tab(25); "Limitada"
      If Res25 = "True" Then Printer.Print Tab(25); "Inadecuada"
    End If
    If Res26 = "True" Or Res27 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Motivo:"
      Printer.CurrentY = TempY
      If Res26 = "True" Then Printer.Print Tab(18); "Muestra con sangre"
      If Res27 = "True" Then Printer.Print Tab(18); "Muestra mal Fijada"
    End If
    If Res28 = "True" Or Res29 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Células Endocervicales:"
      Printer.CurrentY = TempY
      If Res28 = "True" Then Printer.Print Tab(34); "Presentes"
      If Res29 = "True" Then Printer.Print Tab(34); "Ausentes"
    End If
    If Res30 <> "" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Otros:"
      Printer.CurrentY = TempY
      ImprimeVariasLineas 18, Res30, AnchoPagina
    End If
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "2. EVALUACIÓN CITOLÓGICA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    If Res31 = "True" Or Res32 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Estado de Células:"
      Printer.CurrentY = TempY
      If Res31 = "True" Then Printer.Print Tab(30); "Negativo a Cáncer (Normal)"
      If Res32 = "True" Then Printer.Print Tab(30); "Anormal"
    End If
    
    If Res33 = "1" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Grado de Proceso Inflamatorio:"
      Printer.CurrentY = TempY
      Printer.Print Tab(42); Res34
    End If
    If Res35 <> "" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Células Endometriales en mayores a 50 años:"
      Printer.CurrentY = TempY
      Printer.Print Tab(64); Res35
    End If
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "3. HALLAZGOS ADICIONALES:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    For I = 1 To 10
      Combo(I) = ""
    Next I
    If Res36 = "1" Then Combo(1) = "Cambios Reactivos"
    If Res37 = "1" Then Combo(2) = "Sospechoso de PVH"
    If Res38 = "1" Then Combo(3) = "Sospechoso V. Herpes"
    If Res39 = "1" Then Combo(4) = "Atrofia"
    If Res40 = "1" Then Combo(5) = "Bacterias"
    If Res41 = "1" Then Combo(6) = "Trichomonas"
    If Res42 = "1" Then Combo(7) = "Hongos"
    If Res43 = "1" Then Combo(8) = "Otros: " & Res44
    For k = 1 To 10
      For I = 1 To 9
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 10
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      'Printer.FontBold = True
      'Printer.Print Tab(7); "MUESTRA"
      'Printer.FontBold = False
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(10); Combo(k * 5 + 1); Tab(36); Combo(k * 5 + 2); Tab(62); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(115); Combo(k * 5 + 5)
    Next k
    
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "4. DIAGNÓSTICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    If Res45 = "True" Then Printer.Print Tab(10); "Negativo a neoplasia maligna"
    If Res46 = "True" Then Printer.Print Tab(10); "ASC-US"
    If Res47 = "True" Then Printer.Print Tab(10); "ASC-H"
    If Res48 = "True" Then Printer.Print Tab(10); "AGUS"
    
    If Res49 = "True" Or Res50 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Lesión intraepitelial de:"
      Printer.CurrentY = TempY
      If Res49 = "True" Then Printer.Print Tab(34); "Bajo Grado"
      If Res50 = "True" Then Printer.Print Tab(34); "Alto Grado"
    End If
    
    If Res51 = "True" Or Res52 = "True" Or Res53 = "True" Then
      TempY = Printer.CurrentY
      Printer.Print Tab(10); "Sospechoso CA:"
      Printer.CurrentY = TempY
      If Res51 = "True" Then Printer.Print Tab(28); "Invasor"
      If Res52 = "True" Then Printer.Print Tab(28); "Adenocarcinoma"
      If Res53 = "True" Then Printer.Print Tab(28); "Maligno no especificado"
    End If
    
    If Res54 <> "" Then
      Printer.CurrentY = Printer.CurrentY + 0.1
      Printer.FontUnderline = True
      Printer.FontBold = True
      Printer.Print Tab(7); "5. RECOMENDACIONES:"
      Printer.FontUnderline = False
      Printer.FontBold = False
      ImprimeVariasLineas 10, Res54, AnchoPagina
    End If
  ElseIf ml_idPrueba = "CPA002" Then
    'BAAF 'Impronta 'PAP de liquidos y Fluidos
    'CPA002
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    
    Mensaje = "TOMA DE MUESTRA"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
  
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE TOMA DE MUESTRA: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(43); Res01
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "SERVICIO DE PROCEDENCIA: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(43); Res02
    
    Printer.FontBold = True
    Printer.Print Tab(7); "TIPO DE MUESTRA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res03, AnchoPagina
    
    Printer.FontBold = True
    Printer.Print Tab(7); "ANTECEDENTES DE IMPORTANCIA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res04, AnchoPagina
    
    Printer.FontBold = True
    Printer.Print Tab(7); "HISTORIA BREVE: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res05, AnchoPagina
    Printer.Print
    Mensaje = "RESULTADO"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE REALIZACIÓN DEL PROCEDIMIENTO: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(61); Res06
    Printer.FontBold = True
    Printer.Print Tab(7); "PROCEDIMIENTO REALIZADO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res07, AnchoPagina
    ImprimeVariasLineas 10, Res08, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "DIAGNÓSTICO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res09, AnchoPagina
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosHEM(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long) 'Hematología
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "HEM001" Then 'Hemograma completo
    'HEM001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    
    Printer.CurrentY = Printer.CurrentY + 0.05   ' Printer.Print
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "HEMOGLOBINA:"; Tab(70); "HEMATOCRITO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res01 & " g %"; Tab(97); Res02 & " %"
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "HEMATÍES:"; Tab(70); "PLAQUETAS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(22); Res03 & " mm3"; Tab(94); Res04 & " mm3"
    
    Printer.CurrentY = Printer.CurrentY + 0.05   ' Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "FÓRMULA LEUCOCITARIA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "LEUCOCITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res05 & " x mm3"
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "EOSINÓFILOS:"; Tab(38); "BASÓFILOS:"; Tab(68); "LINFOCITOS:"; Tab(96); "MONOCITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(28); Res06 & " %"; Tab(57); Res07 & " %"; Tab(92); Res08 & " %"; Tab(123); Res09 & " %"
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "NEUTRÓFILOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(30); Res10 & " %  ---->"
    
    'TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(45); "MIELOCITOS:"; Tab(85); "METAMIEOLCITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(67); Res11 & " %"; Tab(118); Res12 & " %"
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(45); "ABASTONADOS:"; Tab(85); "SEGMENTADOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(70); Res13 & " %"; Tab(115); Res14 & " %"
  ElseIf ml_idPrueba = "HEM002" Then 'Hemoglobina
    'HEM002
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(11); "RESULTADO"; Tab(57); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
    
    Printer.Print Tab(14); Res01 & " gr  %"; Tab(61); Res02 & " --> " & Res03; Tab(111); Res04
  ElseIf ml_idPrueba = "HEM003" Then 'VSG
    'HEM003
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(40); "VALOR DE REFERENCIA:"; Tab(97); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " mm / hora"; Tab(72); Res02; Tab(119); Res03
  ElseIf ml_idPrueba = "HEM004" Then  'Recuento de reticulocitos
    'HEM004
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(39); "VALOR DE REFERENCIA:"; Tab(80); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " %"; Tab(71); Res02; Tab(100); Res03
  ElseIf ml_idPrueba = "HEM005" Then 'Recuento de plaquetas
    'HEM005
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(39); "VALOR DE REFERENCIA:"; Tab(87); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " mm3"; Tab(71); Res02; Tab(109); Res03
  ElseIf ml_idPrueba = "HEM006" Then 'Tiempo de Coagulación y sangría
    'HEM006
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(11); "DOSAJE DE"; Tab(45); "RESULTADO"; Tab(67); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "1. Tiempo de Coagulación"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(50); Res01 & " min " & Res02 & " seg"; Tab(82); Res03; Tab(114); Res04
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "2. Tiempo de Sangría"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(50); Res05 & " min " & Res06 & " seg"; ; Tab(82); Res07; Tab(114); Res08
  ElseIf ml_idPrueba = "HEM008" Then 'Leishmania
    'HEM008
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & ". "; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(33); Res02; Tab(107); Res03
  ElseIf ml_idPrueba = "HEM009" Then 'Células L.E.
    'Hem009
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & ". "; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(33); Res02; Tab(107); Res03
  ElseIf ml_idPrueba = "HEM010" Then 'lámina periférica
    'HEM010
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
        
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "SERIE BLANCA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "LEUCOCITOS TÍPICOS:"; Tab(50); "LEUCOCITOS ATÍPICOS:"; Tab(90); "BLASTOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(36); Res01 & " %"; Tab(83); Res02 & " %"; Tab(112); Res03 & " %"
        
    Printer.CurrentY = Printer.CurrentY + 0.05   ' Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "SERIE ROJA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "COLOR:"; Tab(50); "TAMAÑO:"; Tab(90); "FORMA:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(19); Res04 & ""; Tab(67); Res05 & ""; Tab(110); Res06 & ""
  ElseIf ml_idPrueba = "HEM011" Then 'Hematocrito
    'HEM011
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(11); "RESULTADO"; Tab(57); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
        
    Printer.Print Tab(15); Res01 & " %"; Tab(62); Res02 & " --> " & Res03; Tab(110); Res04
  ElseIf ml_idPrueba = "HEM012" Then 'tiempo de protombina
    'HEM012
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(11); "RESULTADO"; Tab(57); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
    Printer.Print Tab(7); "Tiempo " & Format(Res01, "00") & " seg"; Tab(70); Res02; Tab(105); Res03
    
    Printer.Print Tab(7); "Protombina " & Res04 & " %"; Tab(50); "INR " & Res05
  ElseIf ml_idPrueba = "HEM013" Then 'grupo y factor sanguíneo
    'HEM013
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "GRUPO SANGUÍNEO: "; Tab(86); "FACTOR RH:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(32); Res01; Tab(110); Res02
  ElseIf ml_idPrueba = "HEM030" Then 'tiempo de coagulación
    'HEM030
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(11); "RESULTADO"; Tab(57); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
    Printer.Print Tab(11); Res01 & " min " & Res02 & " seg"; Tab(72); Res03; Tab(114); Res04
  ElseIf ml_idPrueba = "HEM031" Then 'tiempo de sangria
    'HEM031
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    Printer.FontBold = True
    Printer.Print Tab(11); "RESULTADO"; Tab(57); "VALOR DE REFERENCIA"; Tab(101); "MÉTODO"
    Printer.FontBold = False
    Printer.Print Tab(11); Res01 & " min " & Res02 & " seg"; Tab(72); Res03; Tab(114); Res04
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosINM(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long) 'Inmunoserología
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "INM001" Then  'Diagnóstico de embarazo
    'INM001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN CUALITATIVO: " & Res01; Tab(50); "SENSIBILIDAD:"; Tab(85); "TIPO DE PRUEBA:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(74); Res02; Tab(115); Res03
  ElseIf ml_idPrueba = "INM002" Then 'Sub unidad beta cuantitativo
    'INM002
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    If Res02 <> "" Then Res02 = Res02 & " mUI / ml"
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & ". " & Res02; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res03
  ElseIf ml_idPrueba = "INM003" Then 'VIH-SIDA
    'INM003
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & ". "; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res02
  ElseIf ml_idPrueba = "INM006" Then 'Elisa HBSAg (Antígeno australiano)
    'INM006
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(42); "VALOR DE REFERENCIA:"; Tab(95); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01; Tab(75); Res02; Tab(118); Res03
  ElseIf ml_idPrueba = "INM008" Then 'Aglutinaciones febriles
    'INM008
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "TÍFICO " & Chr(34) & "O" & Chr(34) & ":"; Tab(50); "TÍFICO " & Chr(34) & "H" & Chr(34) & ":"; Tab(90); "TÍFICO" & Chr(34) & "A" & Chr(34) & ":"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(25); Res01 & " " & Res02; Tab(71); Res03 & " " & Res04; Tab(114); Res05 & " " & Res06
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "PARATÍFICO " & Chr(34) & "B" & Chr(34) & ":"; Tab(50); "BRUCELLA :"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(31); Res07 & " " & Res08; Tab(71); Res09 & " " & Res10
  ElseIf ml_idPrueba = "INM009" Then  'RPR o VDRL (Sífilis)
    'INM009
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & ". "; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res02
  ElseIf ml_idPrueba = "INM010" Then  'PCR
    'INM010
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & " " & Res02 & " mg /dl"; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res03
  ElseIf ml_idPrueba = "INM011" Then 'Prueba de Latex
    'INM011
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & " " & Res02 & " mg /dl"; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res03
  ElseIf ml_idPrueba = "INM012" Then 'ASO
    'INM012
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM021" Then 'Toxicológico
    'INM021
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(27); "RESULTADO "; Tab(83); "MÉTODO"
    Printer.FontBold = False
    
    Printer.Print Tab(7); "Marihuana: "; Tab(27); Res01; Tab(86); Res02
    Printer.Print Tab(7); "Cocaína: "; Tab(27); Res03; Tab(86); Res04
  ElseIf ml_idPrueba = "INM030" Then 'VIH Prueba Rápida
    'INM030
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO: " & Res01 & "."; Tab(86); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(107); Res02
  ElseIf ml_idPrueba = "INM031" Then 'ELISA HTLV 1-2
    'INM031
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM032" Then 'Elisa Hepatitis C
    'INM032
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM033" Then  'Elisa Anticore HBC
    'INM033
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM034" Then  'Coombs directo
    'INM034
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM035" Then  'Coombs indirecto
    'INM035
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(38); "VALOR DE REFERENCIA:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(70); Res02; Tab(110); Res03
  ElseIf ml_idPrueba = "INM036" Then  'Antígeno Australiano (HBsAg)
    'INM036
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(89); "MÉTODO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & "."; Tab(110); Res02
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosMIC(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long)    'Microbiologia
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String, Res30 As String
  Dim Res31 As String, Res32 As String, Res33 As String, Res34 As String, Res35 As String, Res36 As String, Res37 As String, Res38 As String, Res39 As String, Res40 As String, Res41 As String, Res42 As String, Res43 As String, Res44 As String, Res45 As String, Res46 As String, Res47 As String, Res48 As String, Res49 As String, Res50 As String, Res51 As String, Res52 As String, Res53 As String, Res54 As String, Res55 As String, Res56 As String, Res57 As String, Res58 As String, Res59 As String, Res60 As String
  Dim Res61 As String, Res62 As String, Res63 As String, Res64 As String, Res65 As String, Res66 As String, Res67 As String, Res68 As String, Res69 As String, Res70 As String, Res71 As String, Res72 As String, Res73 As String, Res74 As String, Res75 As String, Res76 As String, Res77 As String, Res78 As String, Res79 As String, Res80 As String
  Dim I As Integer, J As Integer, k As Integer
  Dim Combo(1 To 30) As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "MIC032" Then 'Parasitologico Directo
    'MIC032
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "PROTOZOARIOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res01
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "HELMINTOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res02
    
    If Res03 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "MÉTODO:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 27, Res03, AnchoPagina
    End If
  ElseIf ml_idPrueba = "MIC033" Then 'Parasitológico seriado
    'MIC033
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "PROTOZOARIOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res01
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "HELMINTOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res02
    
    If Res03 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "MÉTODO:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 27, Res03, AnchoPagina
    End If
  ElseIf ml_idPrueba = "MIC006" Then  'Espermatograma
    'MIC006
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
    Res17 = SeparaCadena(Resu)
    Res18 = SeparaCadena(Resu)
    Res19 = SeparaCadena(Resu)
    Res20 = SeparaCadena(Resu)
    Res21 = SeparaCadena(Resu)
    Res22 = SeparaCadena(Resu)
    Res23 = SeparaCadena(Resu)
    Res24 = SeparaCadena(Resu)
    Res25 = SeparaCadena(Resu)
    Res26 = SeparaCadena(Resu)
    Res27 = SeparaCadena(Resu)
    Res28 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(7); "DATOS DE LA MUESTRA"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "MUESTRA Nº"; Tab(47); "CÓDIGO"; Tab(87); "FECHA"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res01; Tab(57); Res02; Tab(97); Res03
    Printer.FontBold = True
    TempY = Printer.CurrentY
    Printer.Print Tab(7); "ABSTINENCIA"; Tab(47); "HORA TOMA"; Tab(87); "HORA PROCESAMIENTO"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res04; Tab(67); Res05; Tab(107); Res06
    
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN MACROSCÓPICO"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "VOLUMEN"; Tab(47); "PH"; Tab(87); "COLOR"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res07 & " ml"; Tab(57); Res08; Tab(97); Res09
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "ASPECTO"; Tab(47); "T.LICUEFACCIÓN"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res10; Tab(57); Res11 & " min"
    
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN MICROSCÓPICO"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "LEUCOCITOS"; Tab(47); "GLÓBULOS ROJOS"; Tab(87); "GÉRMENES"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res12 & " x c"; Tab(57); Res13 & " x c"; Tab(97); Res14
    If Res15 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "OTROS"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 17, Res15, AnchoPagina
    End If
    
    Printer.FontBold = True
    Printer.Print Tab(7); "MOTILIDAD"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "G3"; Tab(37); "G2"; Tab(67); "G1"; Tab(97); "G0"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res16 & " %"; Tab(47); Res17 & " %"; Tab(77); Res18 & " %"; Tab(107); Res19 & " %";
    
    Printer.FontBold = True
    Printer.Print Tab(7); "CONCENTRACIÓN"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "ESPERMATOZOIDES / ml"; Tab(67); "ESPERMATOZOIDES EYACULADO /ml"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res20 & " x 10"; Tab(87); Res21 & " x 10"
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "ESPERMATOZOIDES MÓTILES / ml"; Tab(67); "ESPERMATOZOIDES EYECTADOS / ml"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res22 & " x 10"; Tab(87); Res23 & " x 10"
    
    Printer.FontBold = True
    Printer.Print Tab(7); "MORFOLOGÍA"
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "NORMALES"; Tab(47); "INMADUROS"; Tab(87); "ANOMALÍAS EN CABEZA"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(17); Res24 & " %"; Tab(57); Res25 & " %"; Tab(97); Res26 & " %"
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(7); "ANOMALÍAS EN SEGMENTO INTERMEDIO"; Tab(87); "ANOMALÍAS EN COLA"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res27 & " %"; Tab(87); Res28 & " %"
    
  ElseIf ml_idPrueba = "MIC031" Then
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
    Res17 = SeparaCadena(Resu)
    Res18 = SeparaCadena(Resu)
    Res19 = SeparaCadena(Resu)
    Res20 = SeparaCadena(Resu)
    Res21 = SeparaCadena(Resu)
    Res22 = SeparaCadena(Resu)
    Res23 = SeparaCadena(Resu)
    Res24 = SeparaCadena(Resu)
    Res25 = SeparaCadena(Resu)
    Res26 = SeparaCadena(Resu)
    Res27 = SeparaCadena(Resu)
    Res28 = SeparaCadena(Resu)
    Res29 = SeparaCadena(Resu)
    Res30 = SeparaCadena(Resu)
    Res31 = SeparaCadena(Resu)
    Res32 = SeparaCadena(Resu)
    Res33 = SeparaCadena(Resu)
    Res34 = SeparaCadena(Resu)
    Res35 = SeparaCadena(Resu)
    Res36 = SeparaCadena(Resu)
    Res37 = SeparaCadena(Resu)
    Res38 = SeparaCadena(Resu)
    Res39 = SeparaCadena(Resu)
    Res40 = SeparaCadena(Resu)
    Res41 = SeparaCadena(Resu)
    Res42 = SeparaCadena(Resu)
    Res43 = SeparaCadena(Resu)
    Res44 = SeparaCadena(Resu)
    Res45 = SeparaCadena(Resu)
    Res46 = SeparaCadena(Resu)
    Res47 = SeparaCadena(Resu)
    Res48 = SeparaCadena(Resu)
    Res49 = SeparaCadena(Resu)
    Res50 = SeparaCadena(Resu)
    Res51 = SeparaCadena(Resu)
    Res52 = SeparaCadena(Resu)
    Res53 = SeparaCadena(Resu)
    Res54 = SeparaCadena(Resu)
    Res55 = SeparaCadena(Resu)
    Res56 = SeparaCadena(Resu)
    Res57 = SeparaCadena(Resu)
    Res58 = SeparaCadena(Resu)
    Res59 = SeparaCadena(Resu)
    Res60 = SeparaCadena(Resu)
    Res61 = SeparaCadena(Resu)
    Res62 = SeparaCadena(Resu)
    Res63 = SeparaCadena(Resu)
    Res64 = SeparaCadena(Resu)
    Res65 = SeparaCadena(Resu)
    Res66 = SeparaCadena(Resu)
    Res67 = SeparaCadena(Resu)
    Res68 = SeparaCadena(Resu)
    Res69 = SeparaCadena(Resu)
    Res70 = SeparaCadena(Resu)
    Res71 = SeparaCadena(Resu)
    Res72 = SeparaCadena(Resu)
    Res73 = SeparaCadena(Resu)
    Res74 = SeparaCadena(Resu)
    Res75 = SeparaCadena(Resu)
    Res76 = SeparaCadena(Resu)
    Res77 = SeparaCadena(Resu)
    Res78 = SeparaCadena(Resu)
    Res79 = SeparaCadena(Resu)
    Res80 = SeparaCadena(Resu)
    
    For I = 1 To 25
      Combo(I) = ""
    Next I
    If Res01 = "1" Then Combo(1) = "Aspirado Gástrico"
    If Res02 = "1" Then Combo(2) = "Cat. Venoso central"
    If Res03 = "1" Then Combo(3) = "Esperma"
    If Res04 = "1" Then Combo(4) = "Esputo"
    If Res05 = "1" Then Combo(5) = "Heces"
    If Res06 = "1" Then Combo(6) = "Hisopado Rectal"
    If Res07 = "1" Then Combo(7) = "L. Ascítico-Peritoneal"
    If Res08 = "1" Then Combo(8) = "L. Biliar"
    If Res09 = "1" Then Combo(9) = "L. Cefalorraquídeo"
    If Res10 = "1" Then Combo(10) = "L. Pericárdico"
    If Res11 = "1" Then Combo(11) = "L. Pleural"
    If Res12 = "1" Then Combo(12) = "L. Sinovial"
    If Res13 = "1" Then Combo(13) = "Orina"
    If Res14 = "1" Then Combo(14) = "Raspado de Piel"
    If Res15 = "1" Then Combo(15) = "Sangre"
    If Res16 = "1" Then Combo(16) = "S. Cerviz"
    If Res17 = "1" Then Combo(17) = "S. Conjuntiva"
    If Res18 = "1" Then Combo(18) = "S. de Herida"
    If Res19 = "1" Then Combo(19) = "S. Faríngea"
    If Res20 = "1" Then Combo(20) = "S. Nasal"
    If Res21 = "1" Then Combo(21) = "S. Ótica"
    If Res22 = "1" Then Combo(22) = "S. Uretral"
    If Res23 = "1" Then Combo(23) = "S. Vaginal"
    If Res24 = "1" Then Combo(24) = "Otro: " & Res25
    For k = 1 To 24
      For I = 1 To 23
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 25
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.FontBold = True
      Printer.Print Tab(7); "MUESTRA"
      Printer.FontBold = False
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(7); Combo(k * 5 + 1); Tab(34); Combo(k * 5 + 2); Tab(61); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(115); Combo(k * 5 + 5)
    Next k
    
    For I = 1 To 25
      Combo(I) = ""
    Next I
    If Res26 = "1" Then Combo(1) = "Coma"
    If Res27 = "1" Then Combo(2) = "Insuficiencia Renal"
    If Res28 = "1" Then Combo(3) = "Neoplasia"
    If Res29 = "1" Then Combo(4) = "Enf. Pulmonar"
    If Res30 = "1" Then Combo(5) = "Cirrosis"
    If Res31 = "1" Then Combo(6) = "Desnutrición"
    If Res32 = "1" Then Combo(7) = "Diabetes Mellitus"
    If Res33 = "1" Then Combo(8) = "Infección por VIH"
    For k = 1 To 8
      For I = 1 To 7
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 9
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.FontBold = True
      Printer.Print Tab(7); "FACTORES INTRÍNSECOS EN INFECCIONES INTRAHOSPITALARIAS"
      Printer.FontBold = False
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(7); Combo(k * 5 + 1); Tab(34); Combo(k * 5 + 2); Tab(61); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(115); Combo(k * 5 + 5)
    Next k
        
    For I = 1 To 25
      Combo(I) = ""
    Next I
    If Res34 = "1" Then Combo(1) = "Sonda Vesical"
    If Res35 = "1" Then Combo(2) = "Cat. Venoso Central"
    If Res36 = "1" Then Combo(3) = "Traqueostomía"
    If Res37 = "1" Then Combo(4) = "Sonda Nasogastrica"
    If Res38 = "1" Then Combo(5) = "Ventilación Mecánica"
    If Res39 = "1" Then Combo(6) = "Cirugía Previa"
    If Res40 = "1" Then Combo(7) = "Cirugía Contaminada"
    For k = 1 To 7
      For I = 1 To 6
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 8
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.FontBold = True
      Printer.Print Tab(7); "FACTORES EXTRÍNSECOS EN INFECCIONES INTRAHOSPITALARIAS"
      Printer.FontBold = False
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(7); Combo(k * 5 + 1); Tab(34); Combo(k * 5 + 2); Tab(61); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(115); Combo(k * 5 + 5)
    Next k
        
    If Res41 <> "" Or Res42 <> "" Or Res43 <> "" Or Res44 <> "" Or Res45 <> "" Or Res46 <> "" Then
      Printer.Print
      Mensaje = "INFORME DE RESULTADO DE EXAMEN MICROBIOLÓGICO"
      Printer.FontUnderline = True
      Printer.FontBold = True
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
      Printer.Print Mensaje
      Printer.FontUnderline = False
      Printer.FontBold = False
    End If
    
    If Res41 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "SEDIMENTO URINARIO / EXAMEN DIRECTO:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 57, Res41, AnchoPagina
    End If
    If Res42 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "GRAM / KOH / OTROS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 35, Res42, AnchoPagina
    End If
    If Res43 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "RECUENTO DE COLONIAS (UFC x ml):"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 50, Res43, AnchoPagina
    End If
    If Res44 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(7); "PRUEBA DE ANTIBIÓTICOS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 40, Res44, AnchoPagina
    End If
    If Res45 <> "" Or Res46 <> "" Then
      Printer.FontBold = True
      Printer.Print Tab(7); "AISLAMIENTO DE GÉRMENES:"
      Printer.FontBold = False
    End If
    If Res45 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(10); "GÉRMEN 1:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 25, Res45, AnchoPagina
    End If
    If Res46 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(10); "GÉRMEN 2:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 25, Res46, AnchoPagina
    End If
        
    For I = 1 To 30
      Combo(I) = ""
    Next I
    If Res47 = "1" Then Combo(1) = "Ac. Nalidixico"
    If Res48 = "1" Then Combo(2) = "Amikacina"
    If Res49 = "1" Then Combo(3) = "Amoxicilina"
    If Res50 = "1" Then Combo(4) = "Amox/Ac.Clavulanico"
    If Res51 = "1" Then Combo(5) = "Ampicilina"
    If Res52 = "1" Then Combo(6) = "Aztreonam"
    If Res53 = "1" Then Combo(7) = "Cefadroxil"
    If Res54 = "1" Then Combo(8) = "Cefalotina"
    If Res55 = "1" Then Combo(9) = "Cefaclor"
    If Res56 = "1" Then Combo(10) = "Cefepime"
    If Res57 = "1" Then Combo(11) = "Cefotaxima"
    If Res58 = "1" Then Combo(12) = "Ceftazidima"
    If Res59 = "1" Then Combo(13) = "Ceftriaxona"
    If Res60 = "1" Then Combo(14) = "Ciprofloxacina"
    If Res61 = "1" Then Combo(15) = "Clindamicina"
    If Res62 = "1" Then Combo(16) = "Cloranfenicol"
    If Res63 = "1" Then Combo(17) = "Cotrimoxazol"
    If Res64 = "1" Then Combo(18) = "Dicloxacilina"
    If Res65 = "1" Then Combo(19) = "Eritromicina"
    If Res66 = "1" Then Combo(20) = "Estreptomicina"
    If Res67 = "1" Then Combo(21) = "Gentamicina"
    If Res68 = "1" Then Combo(22) = "Impenem"
    If Res69 = "1" Then Combo(23) = "Lincomicina"
    If Res70 = "1" Then Combo(24) = "Meropenem"
    If Res71 = "1" Then Combo(25) = "Nitropurantoina"
    If Res72 = "1" Then Combo(26) = "Norfloxacina"
    If Res73 = "1" Then Combo(27) = "Sulfametosaxol"
    If Res74 = "1" Then Combo(28) = "Tetraciclina"
    If Res75 = "1" Then Combo(29) = "Vancomicina"
    For k = 1 To 29
      For I = 1 To 28
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 30
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.Print
      Mensaje = "SENSIBILIDAD"
      Printer.FontUnderline = True
      Printer.FontBold = True
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
      Printer.Print Mensaje
      Printer.FontUnderline = False
      Printer.FontBold = False
    End If
    If J Mod 6 = 0 Then
      I = J \ 6
    Else
      I = J \ 6 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(7); Combo(k * 6 + 1); Tab(29); Combo(k * 6 + 2); Tab(51); Combo(k * 6 + 3); Tab(73); Combo(k * 6 + 4); Tab(95); Combo(k * 6 + 5); Tab(117); Combo(k * 6 + 6)
    Next k
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosBS(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long)   'Banco de Sangre
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String, Res30 As String
  Dim Res31 As String, Res32 As String, Res33 As String, Res34 As String, Res35 As String, Res36 As String, Res37 As String, Res38 As String, Res39 As String, Res40 As String, Res41 As String, Res42 As String, Res43 As String, Res44 As String, Res45 As String, Res46 As String, Res47 As String, Res48 As String, Res49 As String, Res50 As String, Res51 As String, Res52 As String, Res53 As String, Res54 As String, Res55 As String, Res56 As String, Res57 As String, Res58 As String, Res59 As String, Res60 As String
  Dim Res61 As String, Res62 As String, Res63 As String, Res64 As String, Res65 As String, Res66 As String, Res67 As String, Res68 As String, Res69 As String, Res70 As String, Res71 As String, Res72 As String, Res73 As String, Res74 As String, Res75 As String, Res76 As String, Res77 As String, Res78 As String, Res79 As String, Res80 As String, Res81 As String, Res82 As String, Res83 As String, Res84 As String, Res85 As String, Res86 As String, Res87 As String, Res88 As String, Res89 As String, Res90 As String
  Dim Res91 As String, Res92 As String, Res93 As String, Res94 As String, Res95 As String, Res96 As String, Res97 As String, Res98 As String, Res99 As String, Res100 As String, Res101 As String, Res102 As String, Res103 As String, Res104 As String, Res105 As String, Res106 As String, Res107 As String, Res108 As String, Res109 As String, Res110 As String, Res111 As String, Res112 As String, Res113 As String, Res114 As String, Res115 As String, Res116 As String, Res117 As String, Res118 As String, Res119 As String, Res120 As String
  Dim Res121 As String, Res122 As String, Res123 As String, Res124 As String, Res125 As String, Res126 As String, Res127 As String, Res128 As String, Res129 As String, Res130 As String, Res131 As String, Res132 As String, Res133 As String
  Dim Combo(1 To 25) As String
  Dim I As Integer, J As Integer, k As Integer
  
  Printer.Font = "Verdana"
  Printer.FontSize = 8
  Printer.ScaleMode = 7 'vbCentimeters 'Medida en Centimetros
  Printer.CurrentY = 3.9
  Printer.FontSize = 9
  Mensaje = "EG05-FR01: FORMATO DE SELECCIÓN DEL POSTULANTE"
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
    
  If ml_idPrueba = "BAS001" Then 'Tamizaje
    'BS001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
    Res17 = SeparaCadena(Resu)
    Res18 = SeparaCadena(Resu)
    Res19 = SeparaCadena(Resu)
    Res20 = SeparaCadena(Resu)
    Res21 = SeparaCadena(Resu)
    Res22 = SeparaCadena(Resu)
    Res23 = SeparaCadena(Resu)
    Res24 = SeparaCadena(Resu)
    Res25 = SeparaCadena(Resu)
    Res26 = SeparaCadena(Resu)
    Res27 = SeparaCadena(Resu)
    Res28 = SeparaCadena(Resu)
    Res29 = SeparaCadena(Resu)
    Res30 = SeparaCadena(Resu)
    Res31 = SeparaCadena(Resu)
    Res32 = SeparaCadena(Resu)
    Res33 = SeparaCadena(Resu)
    Res34 = SeparaCadena(Resu)
    Res35 = SeparaCadena(Resu)
    Res36 = SeparaCadena(Resu)
    Res37 = SeparaCadena(Resu)
    Res38 = SeparaCadena(Resu)
    Res39 = SeparaCadena(Resu)
    Res40 = SeparaCadena(Resu)
    Res41 = SeparaCadena(Resu)
    Res42 = SeparaCadena(Resu)
    Res43 = SeparaCadena(Resu)
    Res44 = SeparaCadena(Resu)
    Res45 = SeparaCadena(Resu)
    Res46 = SeparaCadena(Resu)
    Res47 = SeparaCadena(Resu)
    Res48 = SeparaCadena(Resu)
    Res49 = SeparaCadena(Resu)
    Res50 = SeparaCadena(Resu)
    Res51 = SeparaCadena(Resu)
    Res52 = SeparaCadena(Resu)
    Res53 = SeparaCadena(Resu)
    Res54 = SeparaCadena(Resu)
    Res55 = SeparaCadena(Resu)
    Res56 = SeparaCadena(Resu)
    Res57 = SeparaCadena(Resu)
    Res58 = SeparaCadena(Resu)
    Res59 = SeparaCadena(Resu)
    Res60 = SeparaCadena(Resu)
    Res61 = SeparaCadena(Resu)
    Res62 = SeparaCadena(Resu)
    Res63 = SeparaCadena(Resu)
    Res64 = SeparaCadena(Resu)
    Res65 = SeparaCadena(Resu)
    Res66 = SeparaCadena(Resu)
    Res67 = SeparaCadena(Resu)
    Res68 = SeparaCadena(Resu)
    Res69 = SeparaCadena(Resu)
    Res70 = SeparaCadena(Resu)
    Res71 = SeparaCadena(Resu)
    Res72 = SeparaCadena(Resu)
    Res73 = SeparaCadena(Resu)
    Res74 = SeparaCadena(Resu)
    Res75 = SeparaCadena(Resu)
    Res76 = SeparaCadena(Resu)
    Res77 = SeparaCadena(Resu)
    Res78 = SeparaCadena(Resu)
    Res79 = SeparaCadena(Resu)
    Res80 = SeparaCadena(Resu)
    Res81 = SeparaCadena(Resu)
    Res82 = SeparaCadena(Resu)
    Res83 = SeparaCadena(Resu)
    Res84 = SeparaCadena(Resu)
    Res85 = SeparaCadena(Resu)
    Res86 = SeparaCadena(Resu)
    Res87 = SeparaCadena(Resu)
    Res88 = SeparaCadena(Resu)
    Res89 = SeparaCadena(Resu)
    Res90 = SeparaCadena(Resu)
    Res91 = SeparaCadena(Resu)
    Res92 = SeparaCadena(Resu)
    Res93 = SeparaCadena(Resu)
    Res94 = SeparaCadena(Resu)
    Res95 = SeparaCadena(Resu)
    Res96 = SeparaCadena(Resu)
    Res97 = SeparaCadena(Resu)
    Res98 = SeparaCadena(Resu)
    Res99 = SeparaCadena(Resu)
    Res100 = SeparaCadena(Resu)
    Res101 = SeparaCadena(Resu)
    Res102 = SeparaCadena(Resu)
    Res103 = SeparaCadena(Resu)
    Res104 = SeparaCadena(Resu)
    Res105 = SeparaCadena(Resu)
    Res106 = SeparaCadena(Resu)
    Res107 = SeparaCadena(Resu)
    Res108 = SeparaCadena(Resu)
    Res109 = SeparaCadena(Resu)
    Res110 = SeparaCadena(Resu)
    Res111 = SeparaCadena(Resu)
    Res112 = SeparaCadena(Resu)
    Res113 = SeparaCadena(Resu)
    Res114 = SeparaCadena(Resu)
    Res115 = SeparaCadena(Resu)
    Res116 = SeparaCadena(Resu)
    Res117 = SeparaCadena(Resu)
    Res118 = SeparaCadena(Resu)
    Res119 = SeparaCadena(Resu)
    Res120 = SeparaCadena(Resu)
    Res121 = SeparaCadena(Resu)
    Res122 = SeparaCadena(Resu)
    Res123 = SeparaCadena(Resu)
    Res124 = SeparaCadena(Resu)
    Res125 = SeparaCadena(Resu)
    Res126 = SeparaCadena(Resu)
    Res127 = SeparaCadena(Resu)
    Res128 = SeparaCadena(Resu)
    Res129 = SeparaCadena(Resu)
    Res130 = SeparaCadena(Resu)
    Res131 = SeparaCadena(Resu)
    Res132 = SeparaCadena(Resu)
    Res133 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(8); "GRUPO SANGUÍNEO:"; Tab(45); "FACTOR RH:"; Tab(85); "CÓDIGO DE DONANTE:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(33); Res04; Tab(64); Res05; Tab(121); Res01
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(8); "FECHA REGISTRO:"; Tab(45); "TIPO DE POSTULANTE:"; Tab(85); "CÓDIGO DE POSTULANTE:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(30); Res03; Tab(76); Res06; Tab(124); Res02
    
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "I. DATOS PERSONALES:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "NOMBRE:"; Tab(55); "SEXO:"; Tab(90); Right(Res12, Len(Res12) - 4) & ":"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(21); Res07 & " " & Res08 & ", "; Res09; Tab(69); Res10; Tab(106); Res13
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "DOMICILIO:"; Tab(55); "ESTADO CIVIL:"; Tab(90); "TELÉFONO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(24); Res15; Tab(79); Right(Res11, Len(Res11) - 4); Tab(114); Res20
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "L.NACIMIENTO:"; Tab(55); "F.NACIMIENTO:"; Tab(90); "EDAD:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(28); Res14; Tab(80); Res16; Tab(108); Trim(Str(EdadActual(CDate(Res16), Now))) & " años"
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "L.PROCEDENCIA:"; Tab(55); "OCUPACIÓN:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Res17 = SepararCadenaDerecha(Res17, "= ")
    Printer.Print Tab(30); Res19; Tab(77); Res17
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "CENTRO DE TRABAJO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(35); Res18
    
    If Res21 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "OBSERVACIONES:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 31, Res21, AnchoPagina
    End If
    
    If Res22 = "1" Then
      Printer.Print
      Mensaje = "DONACIÓN POR REPOSICIÓN"
      Printer.FontBold = True
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
      Printer.Print Mensaje
      TempY = Printer.CurrentY
      Printer.Print Tab(9); "RECEPTOR:"; Tab(55); "Nº H.C.:"; Tab(85); "Nº CAMA:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      Printer.Print Tab(23); Res23; Tab(71); Res27; Tab(107); Res28
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "HOSPITALIZA EN:"; Tab(55); "PARENTESCO:"; Tab(85); "T.ATENCIÓN:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      Printer.Print Tab(31); Res24; Tab(78); Res26; Tab(110); Res29
    End If
    
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "II. PROTOCOLO DE SELECCIÓN AL DONANTE DE SANGRE:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    If Res30 = "True" Then
      Printer.Print Tab(9); "1. Si ha donado sangre alguna vez."
    Else
      If Res31 = "True" Then Printer.Print Tab(9); "1. No ha donado sangre alguna vez."
    End If
    If Res32 = "True" Then
      Printer.Print Tab(9); "2. Donó sangre los últimos 3 meses."
    Else
      If Res33 = "True" Then Printer.Print Tab(9); "2. No donó sangre los últimos 3 meses."
    End If
    If Res34 = "True" Then
      Printer.Print Tab(9); "3. Se puso nervioso cuando donó sangre."
    Else
      If Res35 = "True" Then Printer.Print Tab(9); "3. No se puso nervioso cuando donó sangre."
    End If
    If Res36 = "True" Then
      Printer.Print Tab(9); "4. Ha sido operado en los últimos 6 meses."
    Else
      If Res37 = "True" Then Printer.Print Tab(9); "4. No ha sido operado en los últimos 6 meses."
    End If
    If Res38 = "True" Then
      Printer.Print Tab(9); "5. Ha recibido sangre, tranplante de órganos ó tejidos."
    Else
      If Res39 = "True" Then Printer.Print Tab(9); "5. No ha recibido sangre, tranplante de órganos ó tejidos."
    End If
    If Res40 = "True" Then
      Printer.Print Tab(9); "6. Ha viajado a zona endémica de paludismo."
    Else
      If Res41 = "True" Then Printer.Print Tab(9); "6. No ha viajado a zona endémica de paludismo."
    End If
    If Res42 = "True" Then
      Printer.Print Tab(9); "7. Viajó fuera del país en los últimos años."
    Else
      If Res43 = "True" Then Printer.Print Tab(9); "7. No viajó fuera del país en los últimos años."
    End If
    If Res44 = "True" Then
      Printer.Print Tab(9); "8. Ha usado drogas ilegales."
    Else
      If Res45 = "True" Then Printer.Print Tab(9); "8. No ha usado drogas ilegales."
    End If
    If Res46 = "True" Then
      Printer.Print Tab(9); "9. Está tomando alguna medicina."
    Else
      If Res47 = "True" Then Printer.Print Tab(9); "9. No está tomando alguna medicina."
    End If
    If Res48 = "True" Then
      Printer.Print Tab(9); "10. Ha sido tatuado."
    Else
      If Res49 = "True" Then Printer.Print Tab(9); "10. No ha sido tatuado."
    End If
    If Res50 = "True" Then
      Printer.Print Tab(9); "11. Ha sido sometido a punción de piel (aretes, acupunturas)."
    Else
      If Res51 = "True" Then Printer.Print Tab(9); "11. No ha sido sometido a punción de piel (aretes, acupunturas)."
    End If
    For I = 1 To 25
      Combo(I) = ""
    Next I
    If Res62 = "1" Then Combo(1) = "Hepatitis"
    If Res63 = "1" Then Combo(2) = "Tuberculosis (5a)"
    If Res64 = "1" Then Combo(3) = "Fiebre Tifoidea (2a)"
    If Res65 = "1" Then Combo(4) = "Fiebre Malta (3a)"
    If Res66 = "1" Then Combo(5) = "Fiebre Amarilla (1a)"
    If Res67 = "1" Then Combo(6) = "Paludismo"
    If Res68 = "1" Then Combo(7) = "Chagas (Rp)"
    If Res69 = "1" Then Combo(8) = "Bartolenosis"
    If Res70 = "1" Then Combo(9) = "Cardiopatías (Rp)"
    If Res71 = "1" Then Combo(10) = "Hipertención Arterial"
    If Res72 = "1" Then Combo(11) = "Convulsiones (Rp)"
    If Res73 = "1" Then Combo(12) = "Hemorragias"
    If Res74 = "1" Then Combo(13) = "Cáncer (Rp)"
    If Res75 = "1" Then Combo(14) = "Diabetes (Rp)"
    If Res76 = "1" Then Combo(15) = "Asma"
    If Res77 = "1" Then Combo(16) = "Fiebre Reumática (Rp)"
    If Res78 = "1" Then Combo(17) = "Enf. Venéreas (3a)"
    If Res79 = "1" Then Combo(18) = "Transt. de Coagulación"
    If Res80 = "1" Then Combo(19) = "Dengue (1a)"
    If Res81 = "1" Then Combo(20) = "Hipertiroidismo"
    If Res82 = "1" Then Combo(21) = "Amebiasis"
    If Res84 = "1" Then Combo(22) = "Mononucleosis"
    If Res84 = "1" Then Combo(23) = "Osteomielitis"
    If Res85 = "1" Then Combo(24) = "Glomerulonefritis"
    For k = 1 To 24
      For I = 1 To 23
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 25
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.Print Tab(9); "12. Ha tenido o tiene alguna(s) de las siguientes enfermedades:"
    Else
      Printer.Print Tab(9); "12. No ha tenido ni tiene ninguna enfermedad."
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(13); Combo(k * 5 + 1); Tab(38); Combo(k * 5 + 2); Tab(63); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(113); Combo(k * 5 + 5)
    Next k
    If Res52 = "True" Then
      Printer.Print Tab(9); "13. Ha tenido contacto directo con personas que tengan hepatitis."
    Else
      If Res53 = "True" Then Printer.Print Tab(9); "13. No ha tenido contacto directo con personas que tengan hepatitis."
    End If
    If Res54 = "True" Then
      Printer.Print Tab(9); "14. Ha tenido contacto directo con personas que tengan ictericia."
    Else
      If Res55 = "True" Then Printer.Print Tab(9); "14. No ha tenido contacto directo con personas que tengan ictericia."
    End If
    For I = 1 To 5
      Combo(I) = ""
    Next I
    If Res88 = "1" Then Combo(1) = "Homosexual"
    If Res89 = "1" Then Combo(2) = "Bisexual"
    If Res90 = "1" Then Combo(3) = "Promiscuo"
    If Res91 = "1" Then Combo(4) = "Prostituta"
    For k = 1 To 4
      For I = 1 To 3
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 25
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.Print Tab(9); "15. Pertence a algún grupo de riesgo:"
    Else
      Printer.Print Tab(9); "15. No pertence a ningún grupo de riesgo."
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(13); Combo(k * 5 + 1); Tab(38); Combo(k * 5 + 2); Tab(63); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(113); Combo(k * 5 + 5)
    Next k
        
    For I = 1 To 5
      Combo(I) = ""
    Next I
    If Res94 = "1" Then Combo(1) = "Homosexual"
    If Res95 = "1" Then Combo(2) = "Bisexual"
    If Res96 = "1" Then Combo(3) = "Promiscuo"
    If Res97 = "1" Then Combo(4) = "Prostituta"
    For k = 1 To 4
      For I = 1 To 3
        If Combo(I) = "" Then
          Combo(I) = Combo(I + 1)
          Combo(I + 1) = ""
        End If
      Next I
    Next k
    For I = 1 To 25
      If Combo(I) = "" Then
        J = I - 1
        Exit For
      End If
    Next I
    If J > 0 Then
      Printer.Print Tab(9); "16. Ha tenido contacto sexual con algún grupo de riesgo:"
    Else
      Printer.Print Tab(9); "16. No ha tenido contacto sexual con ningún grupo de riesgo."
    End If
    If J Mod 5 = 0 Then
      I = J \ 5
    Else
      I = J \ 5 + 1
    End If
    For k = 0 To I - 1
      Printer.Print Tab(13); Combo(k * 5 + 1); Tab(38); Combo(k * 5 + 2); Tab(63); Combo(k * 5 + 3); Tab(88); Combo(k * 5 + 4); Tab(113); Combo(k * 5 + 5)
    Next k
    If Res56 = "True" Then
      Printer.Print Tab(9); "17. Tuvo contacto sexual con más de una persona en los últimos 3 años."
    Else
      If Res57 = "True" Then Printer.Print Tab(9); "17. No tuvo contacto sexual con más de una persona en los últimos 3 años."
    End If
    If Res58 = "True" Then
      Printer.Print Tab(9); "18. Tiene SIDA o ha tenido alguna prueba positiva de SIDA."
    Else
      If Res59 = "True" Then Printer.Print Tab(9); "18. No tiene SIDA o no ha tenido alguna prueba positiva de SIDA."
    End If
    If Res60 = "True" Then
      Printer.Print Tab(9); "19. Ha sido excluido como donante anteriormente."
    Else
      If Res61 = "True" Then Printer.Print Tab(9); "19. No ha sido excluido como donante anteriormente."
    End If
    If Res10 = "Femenino" Then
      Printer.Print Tab(9); "20. Fecha de su última regla " & Res100 & "."
      Printer.Print Tab(9); "21. Menstrúa " & Res101 & " días."
      If Res102 = "True" Then
        Printer.Print Tab(9); "22. En su menstruación el sangrado es abundante."
      Else
        If Res103 = "True" Then
          Printer.Print Tab(9); "22. En su menstruación el sangrado es moderado."
        Else
          If Res104 = "True" Then Printer.Print Tab(9); "22. En su menstruación el sangrado es escaso."
        End If
      End If
      If Res105 = "True" Then
        Printer.Print Tab(9); "23. Está gestando."
      Else
        If Res106 = "True" Then Printer.Print Tab(9); "23. No está gestando."
      End If
      Printer.Print Tab(9); "24. Fecha del último parto " & Res107 & "."
      If Res108 = "True" Then
        Printer.Print Tab(9); "25. Está dando de lactar."
      Else
        If Res109 = "True" Then Printer.Print Tab(9); "25. No está dando de lactar."
      End If
    End If
    
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "III. EXAMEN CLÍNICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "PESO:"; Tab(35); "TALLA:"; Tab(59); "PRESIÓN ARTERIAL:"; Tab(98); "PULSO:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(18); Res110 & " Kg"; Tab(48); Res111 & " mt"; Tab(89); Res112 & " mmHg"; Tab(118); Res113 & " pul/min"
    If Res114 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "ESTADO ACCESOS VENOSOS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 42, Res114, AnchoPagina
    End If
    If Res115 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "OBSERVACIONES:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 31, Res115, AnchoPagina
    End If
    
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "IV. EXÁMENENES COMPLEMENTARIOS:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "HEMATOCRITO:"; Tab(49); "HEMOGLOBINA:"; Tab(89); "HbsAg:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(28); Res116; Tab(73); Res117; Tab(108); Res118
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "VDRL/RPR:"; Tab(49); "G.SANGUÍNEO:"; Tab(89); "FACTOR RH:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(24); Res119; Tab(72); Res120; Tab(114); Res121
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "FENOTIPO RH:"; Tab(49); "ANTI HTLV:"; Tab(89); "ANTICORE VHB:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res122; Tab(69); Res123; Tab(118); Res124
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "ANTI CHAGAS:"; Tab(49); "ANTI HIV:"; Tab(89); "ANTI VHC:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res125; Tab(67); Res126; Tab(112); Res127
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "MALARIA:"; Tab(49); "BARTONELLA:"; Tab(89); "VARIANTE DU:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(22); Res128; Tab(71); Res129; Tab(116); Res130
    
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "V. CALIFICACIÓN DEL DONANTE:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontBold = True
    If Res131 = "True" Then
      Printer.Print Tab(9); "APTO."
    Else
      If Res132 = "True" Then
        Printer.Print Tab(9); "NO APTO TEMPORALMENTE."
      Else
        If Res133 = "True" Then Printer.Print Tab(9); "NO APTO PERMANENTEMENTE."
      End If
    End If
    Printer.FontBold = False
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosCOA(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long)  'CoproAnálisis
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "COA003" Then  'Test de Graham
    'COA003
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); Res01
    Printer.FontBold = False
    Printer.CurrentY = TempY
    ImprimeVariasLineas 20, Res02, AnchoPagina
  ElseIf ml_idPrueba = "COA004" Then 'Investigación de Criptosporidium y Coccidias
    'COA004
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); Res01
    Printer.FontBold = False
    If Res02 <> "" Then
      Printer.CurrentY = TempY
      ImprimeVariasLineas 20, Res02, AnchoPagina
    End If
  ElseIf ml_idPrueba = "COA031" Then 'Reacción Inflamatoria
    'COA031
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "REACCIÓN INFLAMATORIA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(40); Res01
        
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "LEUCOCITOS:"; Tab(51); "PMN:"; Tab(88); "MN:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(26); Res02 & " x c"; Tab(64); Res03 & " %"; Tab(105); Res04 & " %"
  ElseIf ml_idPrueba = "COA005" Then 'Coprofuncional de Heces
    'COA005
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
        
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN FÍSICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "ASPECTO:"; Tab(45); "COLOR:"; Tab(81); "OLOR:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(22); Res01; Tab(59); Res02; Tab(98); Res03
        
    Printer.CurrentY = Printer.CurrentY + 0.05  'Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN QUÍMICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "AZÚCARES REDUCTORES (BENEDICT):"; Tab(72); "SANGRE OCULTA (THEVENON):"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(53); Res04; Tab(115); Res05 & " " & Res06
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "GRASAS (SUDAN):"; Tab(72); "REACCIÓN (pH):"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(32); Res07; Tab(100); Res08
        
    Printer.CurrentY = Printer.CurrentY + 0.05   ' Printer.Print
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN MICROSCÓPICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "LEUCOCITOS:"; Tab(81); "PARÁSITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res09 & " x c"; Tab(105); Res10
    If Res11 <> "" Then
      TempY = Printer.CurrentY
      Printer.CurrentY = TempY
      Printer.FontBold = True
      Printer.Print Tab(9); "OTROS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 20, Res11, AnchoPagina
    End If
        
    Printer.CurrentY = Printer.CurrentY + 0.05   ' Printer.Print
    TempY = Printer.CurrentY
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "REACCIÓN INFLAMATORIA:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(40); Res12
        
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "LEUCOCITOS:"; Tab(45); "PMN:"; Tab(81); "MN:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(26); Res13 & " x c"; Tab(57); Res14 & " %"; Tab(96); Res15 & " %"
  ElseIf ml_idPrueba = "COA030" Then 'Sangre oculta
    'COA030
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    
    TempY = Printer.CurrentY
    Printer.Print Tab(7); "Resultado: " & Res01 & " " & Res02
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosPAQ(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long) 'Patología Quirúrquica
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
  
  If ml_idPrueba = "PAQ001" Then 'Pieza operatoria mayor, Biopsia Quirurgica, Pieza operatoria Mediana, Pieza Operatoria pequeña
    'PAQ001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    
    Mensaje = "TOMA DE MUESTRA"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
  
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE TOMA DE MUESTRA: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(43); Res01
    
    Printer.FontBold = True
    Printer.Print Tab(7); "ANTECEDENTES DE IMPORTANCIA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res02, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "HISTORIA BREVE: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res03, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "MUESTRA O ESPECIMEN MOTIVO DE ESTUDIO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res04, AnchoPagina
    Printer.Print
    Mensaje = "RESULTADO"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE REALIZACIÓN DEL PROCEDIMIENTO: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(61); Res05
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN CLÍNICO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res06, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "PROCEDIMIENTO REALIZADO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res07, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "HALLAZGOS OPERATORIOS: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res08, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "IMPRESIÓN DIAGNÓSTICA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res09, AnchoPagina
  ElseIf ml_idPrueba = "PAQ002" Then 'Líquidos con BlockCell
    'PAQ002
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    
    Mensaje = "TOMA DE MUESTRA"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
  
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE TOMA DE MUESTRA: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(43); Res01
    
    Printer.FontBold = True
    Printer.Print Tab(7); "TIPO DE MUESTRA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res02, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "ANTECEDENTES DE IMPORTANCIA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res03, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "HISTORIA BREVE: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res04, AnchoPagina
    Printer.Print
    Mensaje = "RESULTADO"
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
    Printer.Print Mensaje
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "FECHA DE REALIZACIÓN DEL PROCEDIMIENTO: "
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(61); Res05
    Printer.FontBold = True
    Printer.Print Tab(7); "IMPRESIÓN DIAGNÓSTICA: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res06, AnchoPagina
    Printer.FontBold = True
    Printer.Print Tab(7); "PROCEDIMIENTO REALIZADO: "
    Printer.FontBold = False
    ImprimeVariasLineas 10, Res07, AnchoPagina
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabImprimeResultadosANA(Resu As String, ml_idPrueba As String, NomPru As String, ml_observacion As String, ml_realiza As Long)  'Urianálisis
  Dim TempY As Single, TempX As Single, Mensaje As String
  Dim Res01 As String, Res02 As String, Res03 As String, Res04 As String, Res05 As String, Res06 As String, Res07 As String, Res08 As String, Res09 As String, Res10 As String, Res11 As String, Res12 As String, Res13 As String, Res14 As String, Res15 As String, Res16 As String, Res17 As String, Res18 As String, Res19 As String, Res20 As String, Res21 As String, Res22 As String, Res23 As String, Res24 As String, Res25 As String, Res26 As String, Res27 As String, Res28 As String, Res29 As String
  
  Printer.FontSize = 9
  Mensaje = "ANÁLISIS: " & UCase(NomPru)
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.2 'Printer.Print
    
  If ml_idPrueba = "ANA001" Then 'Examen completo de orina
    'ANA001
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
    Res11 = SeparaCadena(Resu)
    Res12 = SeparaCadena(Resu)
    Res13 = SeparaCadena(Resu)
    Res14 = SeparaCadena(Resu)
    Res15 = SeparaCadena(Resu)
    Res16 = SeparaCadena(Resu)
    Res17 = SeparaCadena(Resu)
    Res18 = SeparaCadena(Resu)
    Res19 = SeparaCadena(Resu)
        
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN FÍSICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "COLOR:"; Tab(45); "ASPECTO:"; Tab(81); "DENSIDAD:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(19); Res01; Tab(64); Res02; Tab(104); Res03
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "OLOR:"; Tab(45); "REACCIÓN:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(19); Res04; Tab(64); Res05
    
    Printer.CurrentY = Printer.CurrentY + 0.05
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN QUÍMICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "GLUCOSA:"; Tab(45); "UROBILINOGENO:"; Tab(81); "HEMOGLOBINA:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(22); Res06; Tab(71); Res07; Tab(109); Res08
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "PROTEINAS:"; Tab(45); "CUERPOS CETÓNICOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(25); Res09; Tab(76); Res10
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "BILIRRUBINA:"; Tab(45); "NITRITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(27); Res11; Tab(63); Res12
    
    Printer.CurrentY = Printer.CurrentY + 0.05
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.Print Tab(7); "EXAMEN MICROSCÓPICO:"
    Printer.FontUnderline = False
    Printer.FontBold = False
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "CEL.EPITELIALES:"; Tab(45); "HEMATÍES:"; Tab(81); "LEUCOCITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(31); Res13 & " x c"; Tab(65); Res14 & " x c"; Tab(107); Res16 & " x c"
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "PUS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(18); Res18 & " x c"
        
    If Res17 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "CILINDROS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 25, Res17, AnchoPagina
    End If
        
    If Res19 <> "" Then
      TempY = Printer.CurrentY
      Printer.FontBold = True
      Printer.Print Tab(9); "CRISTALES:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 25, Res19, AnchoPagina
    End If
        
    If Res15 <> "" Then
      TempY = Printer.CurrentY
      Printer.CurrentY = TempY
      Printer.FontBold = True
      Printer.Print Tab(9); "OTROS:"
      Printer.FontBold = False
      Printer.CurrentY = TempY
      ImprimeVariasLineas 20, Res15, AnchoPagina
    End If
  ElseIf ml_idPrueba = "ANA003" Then 'proteína en orina de 24 horas
    'ANA003
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    
    Printer.FontBold = True
    Printer.Print Tab(14); "RESULTADO:"; Tab(50); "VALOR DE REFERENCIA:"; Tab(97); "MÉTODO:"
    Printer.FontBold = False
    Printer.Print Tab(14); Res01 & " gr / l / 24 h"; Tab(53); Res02; Tab(103); Res03
  ElseIf ml_idPrueba = "ANA010" Then 'prueba de Addis
    'ANA010
    Res01 = SeparaCadena(Resu)
    Res02 = SeparaCadena(Resu)
    Res03 = SeparaCadena(Resu)
    Res04 = SeparaCadena(Resu)
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.FontBold = True
    Printer.Print Tab(9); "VOLUMEN:"; Tab(45); "HEMATÍES:"; Tab(81); "LEUCOCITOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(24); Res01; Tab(64); Res02; Tab(107); Res03
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(9); "CEL.EPITELIALES:"; Tab(45); "CIL.HIALINOS:"; Tab(81); "CIL.GRANULOSOS:"
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(32); Res04; Tab(68); Res05; Tab(113); Res06
  End If
  If ml_realiza <> 0 Then LabImprimeRealizaPrueba LabEmpleado(ml_realiza)
  If ml_observacion <> "" Then LabImprimeObservaciones ml_observacion
  Printer.FontSize = 8
  Printer.CurrentY = Printer.CurrentY + 0.1
  Printer.Print Tab(5); "================================================================================="
End Function

Public Function LabRecuperaResultados_Obs(CAnalisis As String, Orden As Double) As String
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  LabRecuperaResultados_Obs = ""
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idorden='" & Orden & "'"
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      If oTabla!ObservacionResultado <> "" Then LabRecuperaResultados_Obs = oTabla!ObservacionResultado
    End If
    oTabla.Close
    oConexSQL.Close
  End If
End Function

Public Function LabIngresaResultados(CAnalisis As String, Orden As Double, Resultado As String, Observacion As String, Usuario As Double, RealizaPrueba As Long, Detalleorden As ADODB.Recordset, NroOrden As Long)
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idOrden='" & Orden & "'"
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      oTabla.Close
      Busca = "UPDATE labresultado SET idOrden='" & Orden & "', ResultadoAnalisis='" & Resultado & "', ObservacionResultado='" & Observacion & "', idUsuario='" & Usuario & "', realizaAnalisis='" & RealizaPrueba & "' WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idOrden='" & Orden & "'"
      oTabla.Open Busca, oConexSQL
      MsgBox "Resultado de Prueba actualizado correctamente.", vbInformation, "SIGH GalenHos"
    Else
      oTabla.Close
      Busca = "INSERT INTO labresultado (idAnalisis, idOrden, ResultadoAnalisis, ObservacionResultado, idusuario, realizaAnalisis) VALUES ('" & CAnalisis & "', '" & Orden & "', '" & Resultado & "','" & Observacion & "','" & Usuario & "', '" & RealizaPrueba & "')"
      oTabla.Open Busca, oConexSQL
      MsgBox "Resultado de Prueba grabado correctamente.", vbInformation, "SIGH GalenHos"
    End If
    'oTabla.Close
    oConexSQL.Close
    If LabAveriguaSiSeHanRegistradoTodosLosResultados(Detalleorden, NroOrden) = True Then
      LabActualizaEstadoFacturacion Orden
      LabActualizaEstadoLaboratorio CDbl(NroOrden)
    End If
'    Unload Me
  End If
End Function

Public Function SeparaCadena(Texto As String) As String
  Dim E As Integer
  E = InStr(1, Texto, "\", vbTextCompare)
  If E = 0 Then
    SeparaCadena = Texto
    Texto = ""
    If SeparaCadena = "Verdadero" Then SeparaCadena = "True"
    If SeparaCadena = "Falso" Then SeparaCadena = "False"
  Else
    SeparaCadena = Trim(Left(Texto, E - 1))
    Texto = Trim(Right(Texto, Len(Texto) - E))
    If SeparaCadena = "Verdadero" Then SeparaCadena = "True" 'Or SeparaCadena = "1"
    If SeparaCadena = "Falso" Then SeparaCadena = "False" 'Or SeparaCadena = "0"
  End If
End Function

Public Function SepararCadenaDerecha(Texto As String, Busca As String) As String
  Dim E As Integer
  E = InStr(1, Texto, Busca, vbTextCompare)
  If E = 0 Then
    SepararCadenaDerecha = Right(Texto, Len(Texto) - Len(Busca))
  Else
    SepararCadenaDerecha = Trim(Right(Texto, Len(Texto) - E))
  End If
End Function

Public Function SepararCadenaIzquierda(Texto As String, Busca As String) As String
  Dim E As Integer
  E = InStr(1, Texto, Busca, vbTextCompare)
  If E = 0 Then
    SepararCadenaIzquierda = ""
  Else
    SepararCadenaIzquierda = Trim(Left(Texto, E - 1))
  End If
End Function

Public Function SeparaCaracteres(Texto As String) As String
  Texto = Trim(Texto)
  SeparaCaracteres = Left(Texto, 1)
  Texto = Right(Texto, Len(Texto) - 1)
End Function

Public Function LabImprimeResultados(CAnalisis As Double, Orden As Double, Resultado As String, Observacion As String)
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & CAnalisis & "'"
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      oTabla.Close
      Busca = "UPDATE labresultado SET idOrden='" & Orden & "', ResultadoAnalisis='" & Resultado & "', ObservacionResultado='" & Observacion & "', idUsuario='" & Usuario & "' WHERE idAnalisis= '" & CAnalisis & "'"
      oTabla.Open Busca, oConexSQL
      MsgBox "Datos de prueba actualizadas correctamente", vbInformation, "GalenHos"
    Else
      oTabla.Close
      Busca = "INSERT INTO labresultado (idAnalisis, idOrden, ResultadoAnalisis, ObservacionResultado, idusuario) VALUES ('" & CAnalisis & "', '" & Orden & "', '" & Resultado & "','" & Observacion & "','" & Usuario & "')"
      oTabla.Open Busca, oConexSQL
      MsgBox "Datos de prueba grabadas correctamente", vbInformation, "GalenHos"
    End If
    oConexSQL.Close
  End If
End Function

Function EmpleadosDeLab(lcFiltro As Long) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set EmpleadosDeLab = Nothing
    sSQL1 = "SELECT dbo.empleados.IdEmpleado, LTRIM(dbo.empleados.ApellidoPaterno) + ' ' + LTRIM(dbo.empleados.ApellidoMaterno) + ' ' + dbo.empleados.Nombres AS ApNom, dbo.empleadosCargos.idCargo " & _
            " FROM dbo.usuariosRoles, dbo.empleados LEFT OUTER JOIN dbo.empleadosCargos ON dbo.empleados.IdEmpleado = dbo.empleadosCargos.idEmpleado " & _
            " WHERE dbo.empleados.idEmpleado=dbo.usuariosRoles.idEmpleado AND dbo.usuariosRoles.idRol= " & lcFiltro & _
            " ORDER BY LTRIM(dbo.empleados.ApellidoPaterno) + ' ' + LTRIM(dbo.empleados.ApellidoMaterno) + ' ' + dbo.empleados.Nombres"
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = sSQL1
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set EmpleadosDeLab = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
    
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

Function TodosEmpleadosDeLab() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String: Dim sSQL2 As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set TodosEmpleadosDeLab = Nothing
    sSQL1 = "SELECT dbo.Empleados.IdEmpleado, LTRIM(dbo.Empleados.ApellidoPaterno) + ' ' + LTRIM(dbo.Empleados.ApellidoMaterno) + ' ' + dbo.Empleados.Nombres AS ApNom, dbo.EmpleadosCargos.idCargo " & _
            " FROM dbo.usuariosRoles, dbo.Empleados LEFT OUTER JOIN dbo.EmpleadosCargos ON dbo.Empleados.IdEmpleado = dbo.EmpleadosCargos.idEmpleado " & _
            " WHERE dbo.empleados.idEmpleado=dbo.usuariosRoles.idEmpleado AND (dbo.usuariosRoles.idRol=69 OR dbo.usuariosRoles.idRol=70)" & _
            " ORDER BY LTRIM(dbo.Empleados.ApellidoPaterno) + ' ' + LTRIM(dbo.Empleados.ApellidoMaterno) + ' ' + dbo.Empleados.Nombres"
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = sSQL1
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set TodosEmpleadosDeLab = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
    
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

Public Function LabEmpleado(idEmpleado As Long) As String
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  LabEmpleado = ""
  Busca = "SELECT * FROM empleados WHERE idEmpleado= '" & idEmpleado & "';"
  oTabla.CursorLocation = adUseClient
  oTabla.LockType = adLockOptimistic
  oTabla.Open Busca, oConexSQL
  If Not (oTabla.EOF = True And oTabla.BOF = True) Then
    If oTabla!ApellidoPaterno <> "" Then LabEmpleado = oTabla!ApellidoPaterno
    If oTabla!ApellidoMaterno <> "" Then LabEmpleado = LabEmpleado & " " & oTabla!ApellidoMaterno
    If oTabla!Nombres <> "" Then LabEmpleado = LabEmpleado & " " & oTabla!Nombres
  End If
  oTabla.Close
  oConexSQL.Close
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
  oConexion.Open SIGHComun.CadenaConexion
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

Function ServiciosDondeEstuvoElPacientePorId(idPaciente As Long, idServicio As Long, idAtencion As Long) As ADODB.Recordset
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
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ServiciosDondeEstuvoElPacientePorId"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ServiciosDondeEstuvoElPacientePorId = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function ServiciosDondeSeAtendioElPaciente(idPaciente As Long, idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar el Servicio que se atendió un paciente por Consulta externa o Emergencia
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ServiciosDondeSeAtendioElPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ServiciosDondeSeAtendioElPaciente = oRecordset
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
  'Procedimiento para averiguar las atenciones SIS que tuvo un paciente
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
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

Function AtencionesQueTuvoElPacienteEnGeneral(idPaciente As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar las atenciones SIS que tuvo un paciente
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AtencionesQueTuvoElPacienteEnGeneral"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set AtencionesQueTuvoElPacienteEnGeneral = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function AtencionesQueTuvoElPacienteConstancias(idPaciente As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar las atenciones que tuvo un paciente para constancias
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AtencionesQueTuvoElPacienteConstancias"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set AtencionesQueTuvoElPacienteConstancias = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function AtencionesQueTuvoElPacienteConstanciasPorId(idPaciente As Long, idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar la atención registrada en una constancia
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AtencionesQueTuvoElPacienteConstanciasPorId"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set AtencionesQueTuvoElPacienteConstanciasPorId = oRecordset
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
  oConexion.Open SIGHComun.CadenaConexion
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

Function SeleccionaFarmaciaPorCuentaYServicio(idAtencion As Long, idServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Farmacia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
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

Function SeleccionaFarmaciaPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Farmacia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaFarmaciaPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaImagenologiaPorCuentaYServicio(idAtencion As Long, idServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Imageneologia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaImagenologiaPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
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

Function SeleccionaImagenologiaPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Imageneologia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaImagenologiaPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaImagenologiaPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaInsumosPorCuentaYServicio(idAtencion As Long, idServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
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

Function SeleccionaInsumosPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaInsumosPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function SeleccionaLaboratorioPorCuentaYServicio(idAtencion As Long, idServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Laboratorio de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaLaboratorioPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
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

Function SeleccionaLaboratorioPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Laboratorio de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaLaboratorioPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaLaboratorioPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function
Function SeleccionaProcedimientosPorCuentaYServicio(idAtencion As Long, idServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaProcedimientosPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, idServicio): .Parameters.Append oParameter
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

Function SeleccionaProcedimientosPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaProcedimientosPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaProcedimientosPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LaboratorioYaRegistroPrueba(IdCuentaAtencion As Long, fechaDespacho As Date, idProductoCPT As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar si ya registro la prueba en las ultimas x horas 'x= parametros hora
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LaboratorioYaRegistroPrueba"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IdCuentaAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@fechaDespacho", adDate, adParamInput, 0, fechaDespacho): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, idProductoCPT): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LaboratorioYaRegistroPrueba = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LaboratorioSeleccionarRol(idEmpleado As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar si ya registro la prueba en las ultimas 12 horas
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LaboratorioSeleccionarRol"
    Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, idEmpleado): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LaboratorioSeleccionarRol = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabMovimientoLaboratorioSeleccionarPorFechas(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar si ya registro la prueba en las ultimas 12 horas
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechas"
    Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorFechas = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

Function LabSeleccionaMedicos() As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim sSQL1 As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set LabSeleccionaMedicos = Nothing
    sSQL1 = "SELECT empleados.idEmpleado, UPPER(LTRIM(empleados.ApellidoPaterno)) + ' ' + UPPER(LTRIM(empleados.ApellidoMaterno)) + ' ' + empleados.nombres AS ApNom " & _
            " FROM (medicos LEFT JOIN empleados ON medicos.idEmpleado = empleados.idEmpleado) " & _
            " ORDER BY apnom "
    With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = sSQL1
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set LabSeleccionaMedicos = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
    
ManejadorDeError:
    ms_MensajeError = Err.Description
End Function

Function LabOrdenConResultados(Detalle As ADODB.Recordset, NroOrden As Long) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  LabOrdenConResultados = False
  oCommand.CommandType = adCmdText
  Set oCommand.ActiveConnection = oConexion
  Detalle.MoveFirst
  Do While Not Detalle.EOF
    sSQL1 = "SELECT * FROM labmovimientolaboratorio, labresultado WHERE labmovimientolaboratorio.idorden=labresultado.idorden AND labresultado.idAnalisis=" & Detalle!codigo & " AND labmovimientolaboratorio.idmovimiento=" & NroOrden
    oCommand.CommandText = sSQL1
    Set oRecordset = oCommand.Execute
    If oRecordset.EOF = True And oRecordset.BOF = True Then
    
    Else
      LabOrdenConResultados = True
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
      Set oRecordset = Nothing
      Exit Function
    End If
    Detalle.MoveNext
  Loop
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

Function LabAveriguaCodigoPrueba(idPrueba As String) As String
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  LabAveriguaCodigoPrueba = ""
  oCommand.CommandType = adCmdText
  Set oCommand.ActiveConnection = oConexion
  sSQL1 = "SELECT * FROM labPruebas WHERE codigoCPT='" & Trim(idPrueba) & "'"
  oCommand.CommandText = sSQL1
  Set oRecordset = oCommand.Execute
  If oRecordset.EOF = True And oRecordset.BOF = True Then
    
  Else
    LabAveriguaCodigoPrueba = oRecordset!codigoPrueba
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Exit Function
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

Function LabAveriguaSiSeHanRegistradoTodosLosResultados(Detalle As ADODB.Recordset, NroOrden As Long) As Boolean
  LabAveriguaSiSeHanRegistradoTodosLosResultados = False
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  Dim TotReg As Integer, TotRes As Integer
  TotReg = 0: TotRes = 0
  oConexion.Open SIGHComun.CadenaConexion
  oConexion.CursorLocation = adUseClient
  LabAveriguaSiSeHanRegistradoTodosLosResultados = False
  oCommand.CommandType = adCmdText
  Set oCommand.ActiveConnection = oConexion
  Detalle.MoveFirst
  Do While Not Detalle.EOF
    TotReg = TotReg + 1
    sSQL1 = "SELECT * FROM labmovimientolaboratorio, labresultado WHERE labmovimientolaboratorio.idorden=labresultado.idorden AND labresultado.idAnalisis=" & Detalle!codigo & " AND labmovimientolaboratorio.idmovimiento=" & NroOrden
    oCommand.CommandText = sSQL1
    Set oRecordset = oCommand.Execute
    If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then TotRes = TotRes + 1
    Detalle.MoveNext
  Loop
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  If TotReg = TotRes Then LabAveriguaSiSeHanRegistradoTodosLosResultados = True
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

Public Function LabActualizaEstadoFacturacion(Orden As Double)
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  Busca = "UPDATE factOrdenServicio SET idEstadoFacturacion='1' WHERE idOrden='" & Orden & "'"
  oTabla.CursorLocation = adUseClient
  oTabla.LockType = adLockOptimistic
  oTabla.Open Busca, oConexSQL
End Function

Public Function LabActualizaEstadoLaboratorio(Orden As Double)
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  Busca = "UPDATE LabMovimiento SET idLabEstado='1' WHERE idMovimiento='" & Orden & "'"
  oTabla.CursorLocation = adUseClient
  oTabla.LockType = adLockOptimistic
  oTabla.Open Busca, oConexSQL
End Function

Public Function PruebaTieneResultado(CAnalisis As String, Orden As Double) As Boolean
  Dim Busca As String
  Dim oConexSQL As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  oConexSQL.Open SIGHComun.CadenaConexion
  
  PruebaTieneResultado = False
  If Val(CAnalisis) <> 0 Then
    Busca = "SELECT * FROM labresultado WHERE idAnalisis= '" & Trim(CAnalisis) & "' AND idorden='" & Orden & "'"
    Set oTabla = New ADODB.Recordset
    oTabla.CursorLocation = adUseClient
    oTabla.LockType = adLockOptimistic
    oTabla.Open Busca, oConexSQL
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then PruebaTieneResultado = True
    oTabla.Close
    oConexSQL.Close
  End If
End Function
