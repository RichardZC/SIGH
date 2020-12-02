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
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para LABORATORIO
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Const AnchoPagina = 120
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

'<Agregado por: WABG el: 11/29/2020-12:22:22 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
Function SeleccionarProcedimientosTamizajeSegunTipoFinanciamiento(lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SeleccionarProcedimientosTamizajeSegunTipoFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_LabBanSangrTamizajeSIS"
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarProcedimientosTamizajeSegunTipoFinanciamiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'</Agregado por: WABG el: 11/29/2020-12:22:22 en el equipo: SISGALENPLUS-PC><CAMBIO 44>
'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  
  Set LabMovimientoSeleccionarPorFechasPuntoCarga = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioAgregar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, _
                   oDoPaciente As DOPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, _
                            idCuentaAtencion As Long, rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                                    mo_lcNombrePc As String, lnIdReceta As Long, lcBoleta As String, _
                                                        lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
     lcDx As String, lcDxCodigo As String, lcUPS As String, lnEpsPorcentaje As Double, lnSiIdCita As Long) As Boolean
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
  Dim oDoSiCitas As New DoSiCitas
  Dim oSiCitas As New SiCitas
  Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
  Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
  Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
  Dim bProcesoOK As Boolean
  Dim rsPacientes As New Recordset
  Dim oRsTmp1  As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  Dim lcSql As String, lnImporteTotal As Double
  Dim lnIdComprobantePago As Long
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  
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
  Set oRecetaCabecera.Conexion = oConexion
  Set oRecetaDetalleItem.Conexion = oConexion
  Set oSiCitas.Conexion = oConexion
  
  If Not oLabMovimiento.Insertar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarILabA
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "A", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                  'ListBarItems.idListItem
  
  If idCuentaAtencion > 0 Then
    'Paciente registrado ingresando Nro Cuenta
    If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
    oDoLabMovimientoLaboratorio.IdOrden = oDoFactOrdenServ.IdOrden
  Else
    'Paciente registrado ingresando Nro Boleta
    If oDoPaciente.idPaciente = 0 Then
      Set rsPacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDoPaciente)
      If rsPacientes.RecordCount > 0 Then
        oDoPaciente.idPaciente = rsPacientes.Fields!idPaciente
      Else
        'oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
        'If Not oPacientes.Insertar(oDOPAciente) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarILabA
      End If
    End If
    oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
'    If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
'       bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
'    End If
    oDoFactOrdenServ.idPaciente = oDoPaciente.idPaciente
    oDoFactOrdenServ.IdPuntoCarga = oDOLabMovimiento.IdPuntoCarga
    oDoFactOrdenServ.fechaDespacho = oDOLabMovimiento.fecha
    If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
       bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarILabA
    End If
    
    'actualiza datos del Paciente en BOLETA
    If oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
    

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "ActualizaDatosPacienteBoleta"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDoPaciente.idPaciente): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, Left(Trim(oDoPaciente.ApellidoPaterno) & " " & Trim(oDoPaciente.ApellidoMaterno) & " " & Trim(oDoPaciente.PrimerNombre), 50)): .Parameters.Append oParameter
          .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing
       
    End If
    
    'Actualiza Paquetes (si fuera el caso)
    If oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
        lnIdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
        rsFacturacionProductosCPT.MoveFirst
        Do While Not rsFacturacionProductosCPT.EOF
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaFacturacionPaquetesEnCaso"
              Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oDOLabMovimiento.IdMovimiento): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, rsFacturacionProductosCPT.Fields!idProducto): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
            End With
            Set oParameter = Nothing
            Set oCommand = Nothing
           
            rsFacturacionProductosCPT.MoveNext
        Loop
    End If
  End If
  
  oDoLabMovimientoLaboratorio.IdMovimiento = oDOLabMovimiento.IdMovimiento
  If Not oLabMovimientoLaboratorio.Insertar(oDoLabMovimientoLaboratorio) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoLaboratorio.MensajeError: GoTo TerminarILabA
  
  If idCuentaAtencion > 0 Then
    'Paciente registrado ingresando Nro Cuenta
    If mo_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
      'If oDoFactOrdenServ.IdTipoFinanciamiento = 1 Or oDoFactOrdenServ.IdTipoFinanciamiento = 5 Then
      '***********Contado ***********
      With oDOFactOrdenServicioPagos
        .FechaCreacion = oDoFactOrdenServ.FechaCreacion
        .IdEstadoFacturacion = 1   'Registrado
        .IdOrden = oDoFactOrdenServ.IdOrden
        .idUsuario = oDoFactOrdenServ.IdUsuarioAuditoria
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
      oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
      oDoLabMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
      If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarILabA
      rsFacturacionProductos.MoveNext
    Loop
  End If
  
  'Agregar Productos (cpt)
  oDoLabMovimientoCPT.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  
  If rsFacturacionProductosCPT.RecordCount > 0 Then rsFacturacionProductosCPT.MoveFirst ' 2308 Frank
  lnImporteTotal = 0
  Do While Not rsFacturacionProductosCPT.EOF
    If rsFacturacionProductosCPT.Fields!idProducto > 0 Then
    oDoLabMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
    oDoLabMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
    oDoLabMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
    oDoLabMovimientoCPT.Importe = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
    oDoLabMovimientoCPT.resultadoAutomatico = IIf(rsFacturacionProductosCPT!resultadoAutomatico = True, 1, 0)
    If Not oLabMovimientoCPT.Insertar(oDoLabMovimientoCPT) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoCPT.MensajeError: GoTo TerminarILabA
    
    If idCuentaAtencion > 0 Then
      'Paciente registrado ingresando Nro Cuenta
      oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
      oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
      oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
      oDoFacturacionServicioDesp.idProducto = rsFacturacionProductosCPT.Fields!idProducto
      oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
      oDoFacturacionServicioDesp.total = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarILabA
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
        If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:                    GoTo TerminarILabA
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
          .TotalFinanciado = rsFacturacionProductosCPT.Fields!totalPorPagar
          .idFuenteFinanciamiento = oDoFactOrdenServ.idFuenteFinanciamiento
        End With
        If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarILabA
        End If
        lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
    End If
    End If
    rsFacturacionProductosCPT.MoveNext
  Loop
  
  'Receta automatica
  'debb-18/05/2016 (inicio)
  If oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
        'Paciente con Boleta
        oDoRecetaCabecera.IdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
        If oRecetaCabecera.SeleccionarPorIdComprobantePagoRecetaDetalleItem(oDoRecetaCabecera) = True Then
           lnIdReceta = oDoRecetaCabecera.idReceta
        End If
  End If
  If lnIdReceta > 0 Then
        oDoRecetaCabecera.idReceta = lnIdReceta
        If oRecetaCabecera.SeleccionarPorId(oDoRecetaCabecera) = True Then
             oDoRecetaCabecera.IdEstado = 2
             oDoRecetaCabecera.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
             If oRecetaCabecera.modificar(oDoRecetaCabecera) = False Then
                bProcesoOK = False: ms_MensajeError = oRecetaCabecera.MensajeError: GoTo TerminarILabA
             End If
             Set oRecetaDetalle.Conexion = oConexion
             rsFacturacionProductosCPT.MoveFirst
             Do While Not rsFacturacionProductosCPT.EOF
                 oDoRecetaDetalle.idReceta = lnIdReceta
                 oDoRecetaDetalle.idItem = rsFacturacionProductosCPT.Fields!idProducto
                 oDoRecetaDetalle.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
                 oDoRecetaDetalle.idEstadoDetalle = 2
                 If oRecetaDetalle.RecetaDetalleActualizaCantDespachada(oDoRecetaDetalle) = False Then
                     bProcesoOK = False: ms_MensajeError = oRecetaDetalle.MensajeError: GoTo TerminarILabA
                 End If
                 oDoRecetaDetalleItem.idReceta = lnIdReceta
                 oDoRecetaDetalleItem.CantidadDespachada = rsFacturacionProductosCPT.Fields!cantidad
                 oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDOLabMovimiento.IdMovimiento))
                 oDoRecetaDetalleItem.idItem = rsFacturacionProductosCPT.Fields!idProducto
                 oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
                 If oDoLabMovimientoLaboratorio.IdComprobantePago = 0 Then
                    If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
                       bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarILabA
                    End If
                 Else
                    oDoRecetaDetalleItem.IdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
                    oDoRecetaDetalleItem.IdEstado = sghRecetaEstados.sighRecetaDespachada
                    If oRecetaDetalleItem.modificar(oDoRecetaDetalleItem) = False Then
                       bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarILabA
                    End If
                 End If
                 rsFacturacionProductosCPT.MoveNext
             Loop
        End If
  End If
  'debb-18/05/2016 (fin)
  If LaboratorioActualizaInterconexionSistemaLaboratorio(sghAgregar, oDoPaciente, rsFacturacionProductosCPT, _
                                                         oDoLabMovimientoLaboratorio, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                         lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion) = False Then
     bProcesoOK = False
  End If
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
  If lnSiIdCita > 0 Then
   oDoSiCitas.idCitaSI = lnSiIdCita
   oDoSiCitas.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
   If oSiCitas.SeleccionarPorId(oDoSiCitas) = True Then
      oDoSiCitas.IdMovimiento = oDOLabMovimiento.IdMovimiento
      oDoSiCitas.IdEstado = sghSiCitasEstados.sghSiCitaConTomaMuestra
      If oSiCitas.modificar(oDoSiCitas) = True Then
      End If
   End If
  End If
  
  '
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
  Set oDoSiCitas = Nothing
  Set oSiCitas = Nothing
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
  Set oRsTmp1 = Nothing
  Set oRecetaCabecera = Nothing
  Set oDoRecetaCabecera = Nothing
  Set oRecetaDetalle = Nothing
  Set oDoRecetaDetalle = Nothing
  Set oRecetaDetalleItem = Nothing
  Set oDoRecetaDetalleItem = Nothing
End Function

Function LaboratorioActualizaInterconexionSistemaLaboratorio(mi_opcion As sghOpciones, oDoPacientes As DOPaciente, _
                                                        rsFacturacionProductos As Recordset, _
                                                        oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, _
                                                        oDoFactOrdenServ As DoFactOrdenServ, lcBoleta As String, _
                                                        lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
                                                        lcDx As String, lcDxCodigo As String, lcUPS As String, _
                                                        oConexion As Connection) As Boolean
    On Error GoTo LabActInter
    Dim TodoOk As Boolean
    Dim oDOInterLabCabecera As New DOInterLabCabecera
    Dim oInterLabCabecera As New InterLabCabecera
    Dim oDOInterLabDetalle As New DOInterLabDetalle
    Dim oInterLabDetalle As New InterLabDetalle
    Dim oDOPacientes1 As New DOPaciente
    Dim oPacientes1 As New Pacientes
    Dim oConexionExterna As New Connection
    TodoOk = False
    LaboratorioActualizaInterconexionSistemaLaboratorio = True
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.BeginTrans
    Set oInterLabDetalle.Conexion = oConexionExterna
    Set oInterLabCabecera.Conexion = oConexionExterna
    Set oPacientes1.Conexion = oConexion
    oDOInterLabCabecera.identificador = oDoLabMovimientoLaboratorio.IdOrden
    oDOInterLabCabecera.IdUsuarioAuditoria = sighentidades.Usuario
    oDOInterLabDetalle.identificador = oDoLabMovimientoLaboratorio.IdOrden
    oDOInterLabDetalle.IdUsuarioAuditoria = sighentidades.Usuario
    If oDoPacientes.idPaciente > 0 Then
        oDOPacientes1.idPaciente = oDoPacientes.idPaciente
        If oPacientes1.SeleccionarPorId(oDOPacientes1) Then
        End If
    End If
    If mi_opcion <> sghAgregar Then
        If oInterLabCabecera.SeleccionarPorId(oDOInterLabCabecera) = False Then
           'GoTo LabActInter
           mi_opcion = sghAgregar
        Else
           If oDOInterLabCabecera.IdEstado <> sghInterLabEstados.sghIntRegistradoEnGalenhos And _
              oDOInterLabCabecera.IdEstado <> sghInterLabEstados.sghIntAnulado Then
              LaboratorioActualizaInterconexionSistemaLaboratorio = False
              MsgBox "No se podrá Modificar/Eliminar porque ya se leyó  en el SISTEMA DE LABORATORIO" & Chr(13) & _
                     "IDENTIFICADOR: " & Trim(Str(oDOInterLabCabecera.identificador)), vbInformation, ""
              GoTo LabActInter
           End If
        End If
    End If
    oDOInterLabCabecera.cama = lcCAma
    oDOInterLabCabecera.dx = lcDx
    oDOInterLabCabecera.dxCodigo = lcDxCodigo
    oDOInterLabCabecera.fechaHoraMuestra = oDoFactOrdenServ.FechaHoraRealizaCpt
    If oDoPacientes.idPaciente = 0 Then
       oDOInterLabCabecera.Fnacimiento = IIf(oDoLabMovimientoLaboratorio.FechaNacimiento = 0, CDate("01/01/1990"), oDoLabMovimientoLaboratorio.FechaNacimiento)
    Else
       oDOInterLabCabecera.Fnacimiento = oDOPacientes1.FechaNacimiento
    End If
    oDOInterLabCabecera.Historia = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oDOPacientes1.NroHistoriaClinica)), False)
    oDOInterLabCabecera.IdEstado = sghInterLabEstados.sghIntRegistradoEnGalenhos
    oDOInterLabCabecera.Medico = lcMedico
    oDOInterLabCabecera.medicoDNI = lcMedicoDNI
    oDOInterLabCabecera.num_orden = IIf(lcBoleta = "", Trim(Str(oDoLabMovimientoLaboratorio.idCuentaAtencion)), lcBoleta)
    oDOInterLabCabecera.Paciente = oDoLabMovimientoLaboratorio.Paciente
    oDOInterLabCabecera.PacienteDocumento = IIf(oDoPacientes.idPaciente = 0, "00000000", oDOPacientes1.NroDocumento)
    oDOInterLabCabecera.PacienteEmail = IIf(oDoPacientes.idPaciente = 0, "@", oDOPacientes1.email)
    oDOInterLabCabecera.PacienteFiliacion = IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 1, "P", _
                                          IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 2, "S", _
                                          IIf(oDoFactOrdenServ.IdTipoFinanciamiento = 4, "C", "O")))
                                          
    If oDoPacientes.idPaciente = 0 Then
       oDOInterLabCabecera.PacienteId = 9000000 + oDoLabMovimientoLaboratorio.IdOrden
    Else
       oDOInterLabCabecera.PacienteId = oDoPacientes.idPaciente
    End If
    
    oDOInterLabCabecera.PacienteTipoDocumento = IIf(oDoPacientes.idPaciente = 0, "1", oDOPacientes1.idDocIdentidad)
    oDOInterLabCabecera.Sexo = IIf(oDoLabMovimientoLaboratorio.idTipoSexo = 1, "M", "F")
    oDOInterLabCabecera.servicioCodigo = lcUPS
    oDOInterLabCabecera.medicoColegiatura = oDoLabMovimientoLaboratorio.colegiatura
    If mi_opcion = sghAgregar Then
        If oInterLabCabecera.Insertar(oDOInterLabCabecera) = False Then
           GoTo LabActInter
        End If
    Else
        If mi_opcion = sghModificar Then
        If oInterLabCabecera.modificar(oDOInterLabCabecera) = False Then
            If oInterLabCabecera.Insertar(oDOInterLabCabecera) = False Then
               GoTo LabActInter
            End If
        End If
        End If
        If oInterLabDetalle.Eliminar(oDOInterLabDetalle) = False Then
            GoTo LabActInter
        End If
        If mi_opcion = sghEliminar Then
            If oInterLabCabecera.Eliminar(oDOInterLabCabecera) = False Then
               GoTo LabActInter
            End If
        End If
    End If
    If mi_opcion <> sghEliminar Then
        rsFacturacionProductos.Filter = "ResultadoAutomatico=true"
        If rsFacturacionProductos.RecordCount = 0 Then
            If oInterLabCabecera.Eliminar(oDOInterLabCabecera) = False Then
               GoTo LabActInter
            End If
        Else
            rsFacturacionProductos.MoveFirst
            Do While Not rsFacturacionProductos.EOF
               oDOInterLabDetalle.cpt = rsFacturacionProductos!nombreProducto
               oDOInterLabDetalle.cptCodigo = rsFacturacionProductos!Codigo
               oDOInterLabDetalle.IdEstado = sghInterLabEstados.sghIntRegistradoEnGalenhos
               If oInterLabDetalle.Insertar(oDOInterLabDetalle) = False Then
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
    Set oDOInterLabCabecera = Nothing
    Set oInterLabCabecera = Nothing
    Set oDOInterLabDetalle = Nothing
    Set oInterLabDetalle = Nothing
    Set oConexionExterna = Nothing
    Set oDOPacientes1 = Nothing
    Set oPacientes1 = Nothing
    Exit Function
    Resume
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioModificar(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, _
                              oDoPaciente As DOPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, _
                              idCuentaAtencion As Long, rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                              mo_lcNombrePc As String, lnIdReceta As Long, lcBoleta As String, _
                                                        lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
                                     lcDx As String, lcDxCodigo As String, lcUPS As String, lnEpsPorcentaje As Double, _
                                     lnIdOrdenPago As Long) As Boolean
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
  Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
  Dim oRecetaDetalle As New RecetaDetalle, oDoRecetaDetalle As New DoRecetaDetalle
  Dim oRecetaDetalleItem As New RecetaDetalleItem, oDoRecetaDetalleItem As New DoRecetaDetalleItem
  Dim bProcesoOK As Boolean
  Dim oRsTmp1 As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
  Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
  Dim lcSql As String, lnImporteTotal As Double
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  On Error GoTo TerminarIMMAGA
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexionShape ' sighentidades.CadenaConexion
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
  Set oRecetaCabecera.Conexion = oConexion
  Set oRecetaDetalleItem.Conexion = oConexion
  
'  If IdCuentaAtencion = 0 Then
'    If oDoPaciente.idPaciente = 0 Then
'      'oDOPAciente.NroHistoriaClinica = oHistoria.GenerarNroHistoria(oDOPAciente.IdTipoNumeracion)
'      'If Not oPacientes.Insertar(oDOPAciente) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
'    Else
'      If Not oPacientes.modificar(oDoPaciente, True) Then bProcesoOK = False: ms_MensajeError = oPacientes.MensajeError: GoTo TerminarIMMAGA
'    End If
'  End If
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
  If Not oLabMovimientoLaboratorio.modificar(oDoLabMovimientoLaboratorio) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoLaboratorio.MensajeError: GoTo TerminarIMMAGA
  '
  oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
  'If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then
  '   bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIMMAGA
  'End If
  oDoFactOrdenServ.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
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
    If oRsTmp1.State = 1 Then
       oRsTmp1.Close
    End If
    Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden, oConexion)
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
    oDoFacturacionServicioFinanc.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
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
      oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
      oDoLabMovimientoDetalle.cantidadFallada = rsFacturacionProductos.Fields!cantidadFallada
      If Not oLabMovimientoDetalle.Insertar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarIMMAGA
      rsFacturacionProductos.MoveNext
    Loop
  End If
  
  'Agregar Productos (cpt)
  oDoLabMovimientoCPT.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoCPT.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
  If rsFacturacionProductosCPT.RecordCount = 0 Then
     bProcesoOK = False: ms_MensajeError = "NO existe Items": GoTo TerminarIMMAGA
  End If
  lnImporteTotal = 0
  rsFacturacionProductosCPT.MoveFirst
  Do While Not rsFacturacionProductosCPT.EOF
    If rsFacturacionProductosCPT.Fields!idProducto > 0 Then
    oDoLabMovimientoCPT.resultadoAutomatico = IIf(rsFacturacionProductosCPT!resultadoAutomatico = True, 1, 0)
    oDoLabMovimientoCPT.cantidad = rsFacturacionProductosCPT.Fields!cantidad
    oDoLabMovimientoCPT.idProductoCPT = rsFacturacionProductosCPT.Fields!idProducto
    oDoLabMovimientoCPT.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
    oDoLabMovimientoCPT.Importe = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) 'rsFacturacionProductosCPT.Fields!totalPorPagar
    If Not oLabMovimientoCPT.Insertar(oDoLabMovimientoCPT) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoCPT.MensajeError: GoTo TerminarIMMAGA
    
    If idCuentaAtencion > 0 Then
      'Paciente registrado ingresando Nro Cuenta
      oDoFacturacionServicioDesp.IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
      oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
      oDoFacturacionServicioDesp.cantidad = rsFacturacionProductosCPT.Fields!cantidad
      oDoFacturacionServicioDesp.idProducto = rsFacturacionProductosCPT.Fields!idProducto
      oDoFacturacionServicioDesp.Precio = rsFacturacionProductosCPT.Fields!PrecioUnitario
      oDoFacturacionServicioDesp.total = Round(rsFacturacionProductosCPT.Fields!cantidad * rsFacturacionProductosCPT.Fields!PrecioUnitario, 2) ' rsFacturacionProductosCPT.Fields!totalPorPagar
      If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo TerminarIMMAGA
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
        If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError:                    GoTo TerminarIMMAGA
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
        lnImporteTotal = lnImporteTotal + oDoFacturacionServicioFinanc.TotalFinanciado
        If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo TerminarIMMAGA
      End If
    End If
    End If
    rsFacturacionProductosCPT.MoveNext
  Loop
  'Receta automatica
  'debb-18/05/2016 (inicio)
  If oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
        'Paciente con Boleta
        oDoRecetaCabecera.IdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
        If oRecetaCabecera.SeleccionarPorIdComprobantePagoRecetaDetalleItem(oDoRecetaCabecera) = True Then
           lnIdReceta = oDoRecetaCabecera.idReceta
        End If
  End If
  If lnIdReceta > 0 Then
       If oDoLabMovimientoLaboratorio.IdComprobantePago = 0 Then
          mo_ReglasImagenes.RecetaAutomaticaCantidadDespachadaCERO lnIdReceta, oConexion
       End If
       Set oRsTmp1 = mo_reglasComunes.RecetaDetalleSeleccioarPorIdReceta(lnIdReceta, oConexion)
       Set oRecetaDetalle.Conexion = oConexion
       If oRsTmp1.RecordCount > 0 Then
            oDoRecetaDetalle.idReceta = lnIdReceta
            oDoRecetaDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
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
               oDoRecetaDetalleItem.DocumentoDespacho = Trim(Str(oDOLabMovimiento.IdMovimiento))
               oDoRecetaDetalleItem.idItem = rsFacturacionProductosCPT.Fields!idProducto
               oDoRecetaDetalleItem.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
               
               If oDoLabMovimientoLaboratorio.IdComprobantePago = 0 Then
                    If oRecetaDetalleItem.Insertar(oDoRecetaDetalleItem) = False Then
                        bProcesoOK = False: ms_MensajeError = oRecetaDetalleItem.MensajeError: GoTo TerminarIMMAGA
                    End If
               Else
                    oDoRecetaDetalleItem.IdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
                    oDoRecetaDetalleItem.IdEstado = sghRecetaEstados.sighRecetaDespachada
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
  If LaboratorioActualizaInterconexionSistemaLaboratorio(sghModificar, oDoPaciente, rsFacturacionProductosCPT, _
                                                         oDoLabMovimientoLaboratorio, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                         lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion) = False Then
     bProcesoOK = False
  End If
  '
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
  Set oRsTmp1 = Nothing
  Set oRecetaCabecera = Nothing
  Set oDoRecetaCabecera = Nothing
  Set oRecetaDetalle = Nothing
  Set oDoRecetaDetalle = Nothing
  Set mo_ReglasImagenes = Nothing
  Set mo_reglasComunes = Nothing
  Set oRecetaDetalleItem = Nothing
  Set oDoRecetaDetalleItem = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioAnular(oDOLabMovimiento As DOLabMovimiento, oDoLabMovimientoLaboratorio As DoLabMovimientoLaboratorio, _
                                 oDoPaciente As DOPaciente, oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, _
                                 idCuentaAtencion As Long, rsFacturacionProductosCPT As Recordset, mo_lnIdTablaLISTBARITEMS As Long, _
                                 mo_lcNombrePc As String, lnIdReceta As Long, lcBoleta As String, _
                                                        lcMedicoDNI As String, lcMedico As String, lcCAma As String, _
                                    lcDx As String, lcDxCodigo As String, lcUPS As String, lnIdOrdenPago As Long) As Boolean
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
  Dim oFacturacionServicioFin As New FacturacionServicioFin
  Dim oDoSiCitas As New DoSiCitas
  Dim oSiCitas As New SiCitas
  Dim oRecetaCabecera As New RecetaCabecera, oDoRecetaCabecera As New DoRecetaCabecera
  Dim bProcesoOK As Boolean
  Dim lnIdComprobantePago As Long
  Dim lcSql As String
  Dim oRsTmp1 As New Recordset
  Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
  Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexionShape
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
  Set oSiCitas.Conexion = oConexion
  'Set oFacturacionServicioFin.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0   'Anulado
  If Not oLabMovimiento.modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarIAMAGA
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "E", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  'Anula Cabecera Productos
  If idCuentaAtencion > 0 Then
    oDoFactOrdenServ.IdOrden = oDoLabMovimientoLaboratorio.IdOrden
    If Not oFactOrdenesServicio.SeleccionarPorId(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
    oDoFactOrdenServ.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
    oDoFactOrdenServ.IdEstadoFacturacion = 9  'anulado
    If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then bProcesoOK = False: ms_MensajeError = oFactOrdenesServicio.MensajeError: GoTo TerminarIAMAGA
    Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden, oConexion)
    If oRsTmp1.RecordCount > 0 Then
      If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
        oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
        If Not oFactServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
        oDOFactOrdenServicioPagos.IdEstadoFacturacion = 9 'anulado
        oDOFactOrdenServicioPagos.IdUsuarioAuditoria = oDoLabMovimientoLaboratorio.IdUsuarioAuditoria
        If Not oFactServicioPagos.modificar(oDOFactOrdenServicioPagos) Then bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo TerminarIAMAGA
      End If
    End If
    If Not oFacturacionServicioFinan.ActualizaIdEstadoFacturacion(oDoFactOrdenServ.IdOrden, 9) Then
       bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo TerminarIAMAGA
    End If
  Else
    'Actualiza Paquetes (si fuera el caso)
    If oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
         lnIdComprobantePago = oDoLabMovimientoLaboratorio.IdComprobantePago
         rsFacturacionProductosCPT.MoveFirst
         Do While Not rsFacturacionProductosCPT.EOF
         
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "ActualizaFacturacionPaquetePorProductoYComprobante"
              Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, rsFacturacionProductosCPT.Fields!idProducto): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
              .Execute
            End With
            Set oParameter = Nothing
            Set oCommand = Nothing
            
            rsFacturacionProductosCPT.MoveNext
         Loop
    End If
  End If
  'Receta automatica debb-03/04/2012

  If lnIdReceta > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalle lnIdReceta, oDoLabMovimientoLaboratorio.idCuentaAtencion, oDoLabMovimientoLaboratorio.IdComprobantePago, oConexion
  ElseIf oDoLabMovimientoLaboratorio.IdComprobantePago > 0 Then
        mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalleSoloPagantes oDoLabMovimientoLaboratorio.IdComprobantePago, oConexion
  End If
  If LaboratorioActualizaInterconexionSistemaLaboratorio(sghEliminar, oDoPaciente, rsFacturacionProductosCPT, _
                                                         oDoLabMovimientoLaboratorio, oDoFactOrdenServ, lcBoleta, lcMedicoDNI, _
                                                         lcMedico, lcCAma, lcDx, lcDxCodigo, lcUPS, oConexion) = False Then
     bProcesoOK = False
  End If
  '
  If mo_AdminAdmision.EPSeliminaOrdenPago(oConexion, lnIdOrdenPago, oDOLabMovimiento.IdUsuarioAuditoria) = False Then
  End If
  '
   
    oDoSiCitas.IdMovimiento = oDOLabMovimiento.IdMovimiento
    oDoSiCitas.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
    If oSiCitas.SiCitasSeleccionarPorIdMovimiento(oDoSiCitas) = True Then
       oDoSiCitas.IdMovimiento = 0
       oDoSiCitas.IdEstado = sghSiCitasEstados.sghSiCitaActiva
       If oSiCitas.modificar(oDoSiCitas) = True Then
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
  Set oDoSiCitas = Nothing
  Set oSiCitas = Nothing
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
  Set oFacturacionServicioFin = Nothing
  Set oRsTmp1 = Nothing
  Set mo_ReglasImagenes = Nothing
  Set oRecetaCabecera = Nothing
  Set oDoRecetaCabecera = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSeleccionarPorId(IdMovimiento As Long) As DOLabMovimiento
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimiento As New LabMovimiento
  Dim oDOLabMovimiento As New DOLabMovimiento
  Set LabMovimientoSeleccionarPorId = Nothing
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoLaboratorioSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoDetalleSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoDetalleSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoCPTSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoCPTSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoLaboratorioSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long, Optional oConexion1 As Connection) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorIdComprobantePago = Nothing
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
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorIdComprobantePago"
    Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
    Set oRecordset = .Execute
    If oConexion1 Is Nothing Then
       Set oRecordset.ActiveConnection = Nothing
    End If
  End With
  Set LabMovimientoLaboratorioSeleccionarPorIdComprobantePago = oRecordset
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoIngresosSeleccionarPorId(IdMovimiento As Long) As DOLabMovimientoIngresos
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimientoIngresos As New LabMovimientoIngresos
  Dim oDoLabMovimientoIngresos As New DOLabMovimientoIngresos
  Set LabMovimientoIngresosSeleccionarPorId = Nothing
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSalidasSeleccionarPorId(IdMovimiento As Long) As DOLabMovimientoSalidas
  Dim oConexion As New ADODB.Connection
  Dim oLabMovimientoSalidas As New LabMovimientoSalidas
  Dim oDoLabMovimientoSalidas As New DOLabMovimientoSalidas
  Set LabMovimientoSalidasSeleccionarPorId = Nothing
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSeleccionarPorFechasPuntoCargaIngresos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSeleccionarPorFechasPuntoCargaIngresos = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSeleccionarPorFechasPuntoCargaSalidas(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSeleccionarPorFechasPuntoCargaSalidas = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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
  
  oConexion.Open sighentidades.CadenaConexion
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
    oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
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
  
  oConexion.Open sighentidades.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoIngresos.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimientoIngresos.modificar(oDoLabMovimientoIngresos) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoIngresos.MensajeError: GoTo TerminarLabMod
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
   
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  If Not oLabMovimientoDetalle.Eliminar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabMod
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
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
  
  oConexion.Open sighentidades.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoIngresos.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0 'Anulado
  If Not oLabMovimiento.modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabAnu
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
  
  oConexion.Open sighentidades.CadenaConexion
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
    oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
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
  
  oConexion.Open sighentidades.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoSalidas.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  If Not oLabMovimientoSalidas.modificar(oDoLabMovimientoSalidas) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoSalidas.MensajeError: GoTo TerminarLabSalMod
  '
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabMovimiento.IdUsuarioAuditoria, "M", oDOLabMovimiento.IdMovimiento, "LabMovimiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")                'ListBarItems.idListItem
  
  oDoLabMovimientoDetalle.IdMovimiento = oDOLabMovimiento.IdMovimiento
  oDoLabMovimientoDetalle.IdUsuarioAuditoria = oDOLabMovimiento.IdUsuarioAuditoria
  If Not oLabMovimientoDetalle.Eliminar(oDoLabMovimientoDetalle) Then bProcesoOK = False: ms_MensajeError = oLabMovimientoDetalle.MensajeError: GoTo TerminarLabSalMod
  rsFacturacionProductos.MoveFirst
  Do While Not rsFacturacionProductos.EOF
    oDoLabMovimientoDetalle.cantidad = rsFacturacionProductos.Fields!cantidad
    oDoLabMovimientoDetalle.idProducto = rsFacturacionProductos.Fields!idProducto
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
  
  oConexion.Open sighentidades.CadenaConexion
  oConexion.BeginTrans
  
  bProcesoOK = True
  ms_MensajeError = ""
  Set oLabMovimiento.Conexion = oConexion
  Set oLabMovimientoSalidas.Conexion = oConexion
  Set oLabMovimientoDetalle.Conexion = oConexion
  
  oDOLabMovimiento.IdLabEstado = 0   'Anulado
  If Not oLabMovimiento.modificar(oDOLabMovimiento) Then bProcesoOK = False: ms_MensajeError = oLabMovimiento.MensajeError: GoTo TerminarLabSalAnul
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

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Function LabMotivoSalidasSeleccionarTodos() As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMotivoSalidasSeleccionarTodos"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMotivoSalidasSeleccionarTodos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure REVISADO
Function LabRecojeExamenSeleccionarTodos() As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabRecojeExamenSeleccionarTodos = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabRecojeExamenSeleccionarTodos"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabRecojeExamenSeleccionarTodos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
  
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoCPTSeleccionarPorFechas(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
    
  Set LabMovimientoCPTSeleccionarPorFechas = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LaboratorioPruebasPorEmpleado(ldFechaIni As Date, ldFechaFin As Date, idEmpleado As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
    
  Set LaboratorioPruebasPorEmpleado = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabFactCatalogoServiciosHospFiltraPorPuntoCarga(lcFiltro As String, lnIdPuntoCarga As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabFactCatalogoServiciosHospFiltraPorPuntoCarga = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnIdProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoDetalleSeleccionarPorFechasYpuntoCarga = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorFechasPuntoCarga = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoIngresosSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoIngresosSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSalidasSeleccionarPorIdMovimiento(lnIdMovimiento As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSalidasSeleccionarPorIdMovimiento = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Function LabRecojeExamenSeleccionaPorID(lnIdRecojeExamen As Long) As String
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  LabRecojeExamenSeleccionaPorID = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabRecojeExamenSeleccionaPorID"
        Set oParameter = .CreateParameter("@IdRecojeExamen", adInteger, adParamInput, 0, lnIdRecojeExamen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
  End With
  oConexion.Close
    
  If oRecordset.RecordCount > 0 Then
    LabRecojeExamenSeleccionaPorID = oRecordset.Fields!RecojeExamen
  Else
    LabRecojeExamenSeleccionaPorID = ""
  End If
    
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing

End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LabMovimientoLaboratorioSeleccionarPorFechasPToInsumos(lnIdPuntoCarga As Long, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorFechasPToInsumos = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Public Function LabRecuperaResultados_Res(CAnalisis As String, Orden As Double) As String

  LabRecuperaResultados_Res = ""
  If Val(CAnalisis) = 0 Then Exit Function
  
  Dim oRecordset As New ADODB.Recordset
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
    .CommandText = "LabRecuperaResultados_Res"
    Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
      If oRecordset!ResultadoAnalisis <> "" Then LabRecuperaResultados_Res = oRecordset!ResultadoAnalisis
  End If
  oRecordset.Close
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Public Function LabRecuperaResultados_ReaP(CAnalisis As String, Orden As Double, ByRef ldFechaResultado As Date) As Long

  LabRecuperaResultados_ReaP = 0
  If Val(CAnalisis) = 0 Then Exit Function
  
  Dim oRecordset As New ADODB.Recordset
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
    .CommandText = "LabRecuperaResultados_ReaP"
    Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
    
  If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
      If oRecordset!ResultadoAnalisis <> "" Then LabRecuperaResultados_ReaP = oRecordset!realizaAnalisis
      ldFechaResultado = IIf(IsNull(oRecordset.Fields!fecha), 0, oRecordset.Fields!fecha)
  End If
  
  oRecordset.Close
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
    
End Function

Public Sub LabImprimeCabeceraResultados(CodPac As String, NomPac As String, HisPac As String, Fec_ As Date, NomMed As String)
  Dim Mensaje, TempY
  If Fec_ = 0 Then Fec_ = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
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
  Printer.Print Tab(25); HCigualDNI_DevuelveHistoriaConCerosIzquierda(HisPac, False); Tab(85); Mensaje
  
   
  
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
  Mensaje = "Atentamente, " & InicialesDelDigitador
  Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Mensaje) / 2
  Printer.Print Mensaje
  Printer.EndDoc
End Sub

Function InicialesDelDigitador() As String
    Dim oConexion As New Connection
    Dim oReglasCaja As New ReglasCaja
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    InicialesDelDigitador = "(Digitador: " & oReglasCaja.SeleccionaDatosCajeroConexion(sighentidades.Usuario, sghIniciales, oConexion) & ")"
    oConexion.Close
    Set oConexion = Nothing
    Set oReglasCaja = Nothing
End Function

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
    Res05 = SeparaCadena(Resu)
    Res06 = SeparaCadena(Resu)
    Res07 = SeparaCadena(Resu)
    Res08 = SeparaCadena(Resu)
    Res09 = SeparaCadena(Resu)
    Res10 = SeparaCadena(Resu)
        
    TempY = Printer.CurrentY
    Printer.CurrentY = TempY
    Printer.FontBold = True
    Printer.Print Tab(7); "RESULTADO:"; Tab(41); "VALOR DE REFERENCIA:"; Tab(96); "MÉTODO:"
    

    
    Printer.FontBold = False
    Printer.CurrentY = TempY
    Printer.Print Tab(23); Res01 & " " & Res02; Tab(73); Res03 & " " & Res02; Tab(118); Res04
    If Res05 <> "" Then
        TempY = Printer.CurrentY
        Printer.CurrentY = TempY
        Printer.Print Tab(7); "BILIRUBINA DIRECTA:"; Res05
    End If
    If Res09 <> "" Then
        TempY = Printer.CurrentY
        Printer.CurrentY = TempY
        Printer.Print Tab(7); "BILIRUBINA INDIRECTA:"; Res09
    End If
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

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Public Function LabRecuperaResultados_Obs(CAnalisis As String, Orden As Double) As String

  LabRecuperaResultados_Obs = ""
  If Val(CAnalisis) = 0 Then Exit Function
  
  Dim oRecordset As New ADODB.Recordset
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
    .CommandText = "LabRecuperaResultados_Obs"
    Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
      If oRecordset!ObservacionResultado <> "" Then LabRecuperaResultados_Obs = oRecordset!ObservacionResultado
  End If
  
  oRecordset.Close
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  
End Function

'debb-2/3/2015
Public Function LabIngresaResultados(CAnalisis As String, Orden As Double, Resultado As String, _
                    Observacion As String, Usuario As Long, RealizaPrueba As Long, _
                    Detalleorden As ADODB.Recordset, NroOrden As Long, lcGrupoSanguineo As String, _
                    lcFactorRh As String, lnIdPaciente As Long, ldFechaResultado As Date, mo_lcNombrePc As String, _
                    mo_lnIdTablaLISTBARITEMS As Long, lcPaciente As String, lcExamen As String, _
                    Optional lnHemoglobina As Double, Optional lnParasitosis As Integer)
  Dim Busca As String
  Dim oConexion As New ADODB.Connection
  Dim oTabla As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oReglasComunes  As New ReglasComunes
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  If Val(CAnalisis) <> 0 Then
  

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "LabConsultaIngresaResultados"
      Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
      Set oTabla = .Execute
      Set oTabla.ActiveConnection = Nothing
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
        
    If Not (oTabla.EOF = True And oTabla.BOF = True) Then
      oTabla.Close
      
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "LabActualizaResultados"
          Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, ldFechaResultado): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ResultadoAnalisis", adVarChar, adParamInput, 5000, Resultado): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ObservacionResultado", adVarChar, adParamInput, 500, Observacion): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, Usuario): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@RealizaAnalisis", adInteger, adParamInput, 50, RealizaPrueba): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
          .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing
        
      Call mo_ReglasSeguridad.AuditoriaAgregarV(Usuario, "M", NroOrden, Left(lcExamen, 50), oConexion, 99999, mo_lcNombrePc, "(Resultado Laboratorio): " & lcPaciente)
      MsgBox "Resultado de Prueba actualizado correctamente.", vbInformation, "SIGH "
    Else
      oTabla.Close
      
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "LabIngresaResultados"
          Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, ldFechaResultado): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ResultadoAnalisis", adVarChar, adParamInput, 5000, Resultado): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ObservacionResultado", adVarChar, adParamInput, 500, Observacion): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, Usuario): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@RealizaAnalisis", adInteger, adParamInput, 50, RealizaPrueba): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
          .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing
      
      Call mo_ReglasSeguridad.AuditoriaAgregarV(Usuario, "A", NroOrden, Left(lcExamen, 50), oConexion, 99999, mo_lcNombrePc, "(Resultado Laboratorio): " & lcPaciente)
      MsgBox "Resultado de Prueba grabado correctamente.", vbInformation, "SIGH "
    End If
    
    'Actualiza GrupoSanguineo y FactorRH en la tabla Pacientes - debb-20/02/2012
    If lnIdPaciente > 0 And (lcGrupoSanguineo <> "" Or lcFactorRh <> "") Then
        ActualizaGrupoSanguineoYfactorRH lnIdPaciente, lcGrupoSanguineo, lcFactorRh, oConexion
'        With oCommand
'          .CommandType = adCmdStoredProc
'          Set .ActiveConnection = oConexion
'          .CommandTimeout = 150
'          .CommandText = "PacientesActualizarGrupoSanYfactorRH"
'          Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
'          Set oParameter = .CreateParameter("@lcGrupoSanguineo", adVarChar, adParamInput, 50, lcGrupoSanguineo): .Parameters.Append oParameter
'          Set oParameter = .CreateParameter("@lcFactorRh", adVarChar, adParamInput, 50, lcFactorRh): .Parameters.Append oParameter
'          .Execute
'        End With
'        Set oParameter = Nothing
'        Set oCommand = Nothing
    
    End If
    

    '
    If LabAveriguaSiSeHanRegistradoTodosLosResultados(Detalleorden, NroOrden) = True Then
      LabActualizaEstadoFacturacion Orden
      LabActualizaEstadoLaboratorio CDbl(NroOrden)
    End If
    '*****************Hemoglobina/parasitosis**********************
    If lnIdPaciente > 0 Then
        Dim oRsTmp1 As New Recordset
        Dim oReglasFacturacion  As New ReglasFacturacion
        Set oRsTmp1 = LabMovimientoLaboratorioSeleccionarPorIdMovimiento(NroOrden)
        If oRsTmp1.RecordCount > 0 Then
            If oRsTmp1!idCuentaAtencion > 0 Then
                If lnHemoglobina > 0 And lnParasitosis >= 0 Then
                     oReglasComunes.ActualizaTablaPacientesMovimientos oConexion, ldFechaResultado, oRsTmp1!idCuentaAtencion, _
                                    sghLaboratorioPC, False, , , , , , , , , , , lnHemoglobina, lnParasitosis
                ElseIf lnHemoglobina > 0 Then
                     oReglasComunes.ActualizaTablaPacientesMovimientos oConexion, ldFechaResultado, oRsTmp1!idCuentaAtencion, _
                                    sghLaboratorioPC, False, , , , , , , , , , , lnHemoglobina
                ElseIf lnParasitosis >= 0 Then
                     oReglasComunes.ActualizaTablaPacientesMovimientos oConexion, ldFechaResultado, oRsTmp1!idCuentaAtencion, _
                                    sghLaboratorioPC, False, , , , , , , , , , , , lnParasitosis
                End If
            End If
        End If
        Set oRsTmp1 = Nothing
        Set oReglasFacturacion = Nothing
    End If
    '
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set oCommand = Nothing
    Set oReglasComunes = Nothing
    '
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

Public Function LabImprimeResultados(CAnalisis As Double, Orden As Double, Resultado As String, _
                                     Observacion As String)
'  Dim Busca As String
'  Dim oConexSQL As New ADODB.Connection
'  Dim oTabla As New ADODB.Recordset
'  oConexSQL.Open sighentidades.CadenaConexion

  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  If Val(CAnalisis) <> 0 Then

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "LabImprimeResultados"
      Set oParameter = .CreateParameter("@IdAnalisis", adInteger, adParamInput, 0, CAnalisis): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ResultadoAnalisis", adVarChar, adParamInput, 50, Resultado): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@ObservacionResultado", adVarChar, adParamInput, 50, Observacion): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, Usuario): .Parameters.Append oParameter
      Set oRecordset = .Execute
    End With
    oConexion.Close
    Set oCommand = Nothing
    End If
    Set oConexion = Nothing
    Set oRecordset = Nothing
    Exit Function
    
ManejadorDeError:
      ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
      Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Function EmpleadosDeLab(lcFiltro As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosDeLab"
        Set oParameter = .CreateParameter("@IdFiltro", adInteger, adParamInput, 0, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
  End With
  Set EmpleadosDeLab = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
        
End Function

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Function TodosEmpleadosDeLab() As ADODB.Recordset

On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oRsEmpleadosLab As New Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Dim lbEsNuevo As Boolean
  With oRsEmpleadosLab
          .Fields.Append "IdEmpleado", adInteger, 4, adFldIsNullable
          .Fields.Append "ApNom", adVarChar, 100, adFldIsNullable
          .Fields.Append "IdCargo", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
  
  
  Set TodosEmpleadosDeLab = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "TodosEmpleadosDeLab"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  If oRecordset.RecordCount > 0 Then
     oRecordset.MoveFirst
     Do While Not oRecordset.EOF
        lbEsNuevo = True
        If oRsEmpleadosLab.RecordCount > 0 Then
           oRsEmpleadosLab.MoveFirst
           oRsEmpleadosLab.Find "idEmpleado=" & oRecordset!idEmpleado
           If Not oRsEmpleadosLab.EOF Then
              lbEsNuevo = False
           End If
        End If
        If lbEsNuevo = True Then
            oRsEmpleadosLab.AddNew
            oRsEmpleadosLab.Fields!idEmpleado = oRecordset!idEmpleado
            oRsEmpleadosLab.Fields!apNom = oRecordset!apNom
            oRsEmpleadosLab.Fields!idCargo = oRecordset!idCargo
            oRsEmpleadosLab.Update
        End If
        oRecordset.MoveNext
     Loop
  End If
  Set TodosEmpleadosDeLab = oRsEmpleadosLab
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 18/09/2013 - se cambio a store procedure
Public Function LabEmpleado(idEmpleado As Long) As String
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  LabEmpleado = ""
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabEmpleado"
    Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, idEmpleado): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With

  If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
    If oRecordset!ApellidoPaterno <> "" Then LabEmpleado = oRecordset!ApellidoPaterno
    If oRecordset!ApellidoMaterno <> "" Then LabEmpleado = LabEmpleado & " " & oRecordset!ApellidoMaterno
    If oRecordset!Nombres <> "" Then LabEmpleado = LabEmpleado & " " & oRecordset!Nombres
  End If
  
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function ServiciosDondeEstuvoElPacientePorId(idPaciente As Long, IdServicio As Long, idAtencion As Long) As ADODB.Recordset
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ServiciosDondeEstuvoElPacientePorId"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, idPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idServicio", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
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


'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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
  'debb-05/04/2011
  Set SeleccionaFarmaciaPorCuentaYServicio = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaFarmaciaPorCuenta = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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
  'debb-05/04/2011
  Set SeleccionaInsumosPorCuentaYServicio = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaInsumosPorCuenta = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function LaboratorioYaRegistroPrueba(idCuentaAtencion As Long, fechaDespacho As Date, idProductoCPT As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar si ya registro la prueba en las ultimas x horas 'x= parametros hora
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LaboratorioYaRegistroPrueba"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idCuentaAtencion): .Parameters.Append oParameter
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

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
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

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarPorFechas(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lnOrden As sghTipoOrden) As ADODB.Recordset
    
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabMovimientoLaboratorioSeleccionarPorFechas = Nothing
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasDinamicoOrden"
    Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@TipoOrden", adInteger, adParamInput, 0, IIf(lnOrden = sghPorIdProductoMasFecha, 1, 0)): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarPorFechas = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure REVISADO
Function LabSeleccionaMedicos() As ADODB.Recordset
On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabSeleccionaMedicos = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabSeleccionaMedicos"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabSeleccionaMedicos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabOrdenConResultados(Detalle As ADODB.Recordset, NroOrden As Long) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String, lnIdAnalisis As Long
  Dim oConexion As New ADODB.Connection
  Dim oParameter As ADODB.Parameter
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  LabOrdenConResultados = False
  'oCommand.CommandType = adCmdText
  oCommand.CommandType = adCmdStoredProc
  Set oCommand.ActiveConnection = oConexion
  Detalle.MoveFirst
  
  Do While Not Detalle.EOF

    lnIdAnalisis = Val(Detalle!Codigo)
    With oCommand
      .CommandTimeout = 150
      .CommandText = "LabOrdenConResultados"
      Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 20, Left(Detalle!Codigo, 20)): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, NroOrden): .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
  
    If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
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

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabAveriguaCodigoPrueba(idPrueba As String) As String
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  LabAveriguaCodigoPrueba = ""
  'oCommand.CommandType = adCmdText
  oCommand.CommandType = adCmdStoredProc
  Set oCommand.ActiveConnection = oConexion
 
  With oCommand
      .CommandTimeout = 150
      .CommandText = "LabAveriguaCodigoPrueba"
      Set oParameter = .CreateParameter("@codigoCPT", adVarChar, adParamInput, 10, Trim(idPrueba)): .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
  End With

  If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then
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

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function LabAveriguaSiSeHanRegistradoTodosLosResultados(Detalle As ADODB.Recordset, NroOrden As Long) As Boolean
  
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim TotReg As Integer, TotRes As Integer
  TotReg = 0: TotRes = 0
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  
  oCommand.CommandType = adCmdStoredProc
  Set oCommand.ActiveConnection = oConexion
  Detalle.MoveFirst
  Do While Not Detalle.EOF
    TotReg = TotReg + 1
    If oRecordset.State = 1 Then oRecordset.Close
    With oCommand
        .CommandType = adCmdStoredProc
        .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabAveriguaSiSeHanRegistradoTodosLosResultados"
        Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 20, Detalle!Codigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, NroOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordSet.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
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
  MsgBox Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Public Function LabActualizaEstadoFacturacion(Orden As Double)
  Dim oConexion As New ADODB.Connection
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
 With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabActualizaEstadoFacturacion"
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
    .Execute
  End With
  
    Set oConexion = Nothing
    Set oRecordset = Nothing
    Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Public Function LabActualizaEstadoLaboratorio(Orden As Double)
'  Dim Busca As String
  Dim oConexion As New ADODB.Connection
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  

    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandTimeout = 150
       .CommandText = "LabActualizaEstadoLaboratorio"
       Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
        .Execute
     End With
  
    Set oConexion = Nothing
    Set oRecordset = Nothing
    Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Public Function PruebaTieneResultado(CAnalisis As String, Orden As Double) As Boolean
  'Dim Busca As String
  Dim oConexion As New ADODB.Connection
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  PruebaTieneResultado = False
  If Val(CAnalisis) <> 0 Then

    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandTimeout = 150
       .CommandText = "PruebaTieneResultado"
       Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
       Set oRecordset = .Execute
       Set oRecordset.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
        
    If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then PruebaTieneResultado = True
    oRecordset.Close
    
    If PruebaTieneResultado = False Then

        Set oCommand = Nothing
        With oCommand
           .CommandType = adCmdStoredProc
           Set .ActiveConnection = oConexion
           .CommandTimeout = 150
           .CommandText = "PruebaTieneResultadoItems"
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, Trim(CAnalisis)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, Orden): .Parameters.Append oParameter
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
       Set oCommand = Nothing
       Set oParameter = Nothing
       
       If Not (oRecordset.EOF = True And oRecordset.BOF = True) Then PruebaTieneResultado = True
       oRecordset.Close
       
    End If
    oConexion.Close
  End If
    Set oConexion = Nothing
    Set oRecordset = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function SacarGruposLaboratorio() As ADODB.Recordset
  
On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set SacarGruposLaboratorio = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SacarGruposLaboratorio"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SacarGruposLaboratorio = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 19/09/2013 - se cambio a store procedure
Function SacarPruebasPorGrupo(Grupo As String) As ADODB.Recordset
  
On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set SacarPruebasPorGrupo = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SacarPruebasPorGrupo"
    Set oParameter = .CreateParameter("@Grupo", adVarChar, adParamInput, 50, Grupo): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SacarPruebasPorGrupo = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
  
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function CuentaPruebas(idProducto As String, FI As String, FF As String, lnIdTipoServicio As Long) As Long
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oRecordset1 As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  'Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  CuentaPruebas = 0
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CuentaPruebas"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CuentaPruebas1"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
    Set oRecordset1 = .Execute
    Set oRecordset1.ActiveConnection = Nothing
  End With
  CuentaPruebas = 0
  If oRecordset.RecordCount > 0 Then
     If Not IsNull(oRecordset!suma) Then
        CuentaPruebas = CuentaPruebas + oRecordset!suma
     End If
  End If
  If oRecordset1.RecordCount > 0 Then
     If Not IsNull(oRecordset1!suma) Then
        CuentaPruebas = CuentaPruebas + oRecordset1!suma
     End If
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Set oRecordset1 = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
'Function CuentaPruebasConResultados(idProducto As String, codigoCPT As String, FI As String, FF As String) As Long
'  On Error GoTo ManejadorDeError
'  Dim oRecordset As New ADODB.Recordset
'  Dim oCommand As New ADODB.Command
'  Dim oParameter As ADODB.Parameter
'  Dim oConexion As New ADODB.Connection
'  oConexion.CursorLocation = adUseClient
'  oConexion.CommandTimeout = 300
'  oConexion.Open sighentidades.CadenaConexion
'  CuentaPruebasConResultados = 0
'
'  With oCommand
'    .CommandType = adCmdStoredProc
'    Set .ActiveConnection = oConexion
'    .CommandTimeout = 150
'    .CommandText = "CuentaPruebasConResultados"
'    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
'    Set oParameter = .CreateParameter("@IdAnalisis", adVarChar, adParamInput, 50, codigoCPT): .Parameters.Append oParameter
'    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
'    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
'    Set oRecordset = .Execute
'    Set oRecordset.ActiveConnection = Nothing
'  End With
'  CuentaPruebasConResultados = 0
'  If oRecordset.RecordCount > 0 Then
'     CuentaPruebasConResultados = IIf(IsNull(oRecordset!suma), 0, oRecordset!suma)
'  End If
'  oConexion.Close
'  Set oConexion = Nothing
'  Set oCommand = Nothing
'  Set oRecordset = Nothing
'  Exit Function
'
'ManejadorDeError:
'  ms_MensajeError = Err.Description
'End Function
Function CuentaPruebasConResultados(lnIdProducto As String, lcCodigoCpt As String, FI As String, FF As String) As Long
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oRecordset1 As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  CuentaPruebasConResultados = 0

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioXfechas"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adDBTimeStamp, adParamInput, 0, CDate(FI)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adDBTimeStamp, adParamInput, 0, CDate(FF)): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  CuentaPruebasConResultados = 0
  If oRecordset.RecordCount > 0 Then
     oRecordset.MoveFirst
     Do While Not oRecordset.EOF
        Set oRecordset1 = LabResultadosPorItemsSeleccionarXfiltro("idOrden=" & oRecordset!IdOrden & " and idProductoCpt=" & lnIdProducto)
        If oRecordset1.RecordCount > 0 Then
           CuentaPruebasConResultados = CuentaPruebasConResultados + oRecordset!cantidad
        Else
            oRecordset1.Close
            Set oParameter = Nothing
            Set oCommand = Nothing
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "LabResultadoSeleccionarPorIdporCpt"
              Set oParameter = .CreateParameter("@codigoCpt", adVarChar, adParamInput, 20, lcCodigoCpt): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRecordset!IdOrden): .Parameters.Append oParameter
              Set oRecordset1 = .Execute
              Set oRecordset1.ActiveConnection = Nothing
            End With
            If oRecordset1.RecordCount > 0 Then
               CuentaPruebasConResultados = CuentaPruebasConResultados + oRecordset!cantidad
            End If
        End If
        oRecordset1.Close
        oRecordset.MoveNext
     Loop
  End If
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Set oRecordset1 = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function




'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function SacarPruebasTodas() As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set SacarPruebasTodas = Nothing
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SacarPruebasTodas"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With

  Set SacarPruebasTodas = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function AveriguaConsumosDeExternos(FI As String, FF As String, idProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  Set AveriguaConsumosDeExternos = Nothing
  

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AveriguaConsumosDeExternos"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With

  Set AveriguaConsumosDeExternos = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function AveriguaConsumosDeConsultoriosExternos(FI As String, FF As String, idProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
'  Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  
  Set AveriguaConsumosDeConsultoriosExternos = Nothing
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AveriguaConsumosDeConsultoriosExternos"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With

  Set AveriguaConsumosDeConsultoriosExternos = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 23/09/2013 - se cambio a store procedure
Function AveriguaConsumosDeEmergencia(FI As String, FF As String, idProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  'Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  
  Set AveriguaConsumosDeEmergencia = Nothing

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AveriguaConsumosDeEmergencia"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With

  Set AveriguaConsumosDeEmergencia = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

Function AveriguaConsumosDeHospitalizacion(FI As String, FF As String, idProducto As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  'Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open sighentidades.CadenaConexion
  
  Set AveriguaConsumosDeHospitalizacion = Nothing
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "AveriguaConsumosDeHospitalizacion"
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  
  Set AveriguaConsumosDeHospitalizacion = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function

Function SeleccionaDiagnosticoDeAtencion(idCuentaAtencion As Long) As String
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  'Dim sSQL1 As String
  Dim oConexion As New ADODB.Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  SeleccionaDiagnosticoDeAtencion = ""
  

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaDiagnosticoDeAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
  End With
  
  If oRecordset.RecordCount > 0 Then SeleccionaDiagnosticoDeAtencion = "Diagnóstico: " & oRecordset!CodigoCIE10 & " --> " & oRecordset!Descripcion
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
    
ManejadorDeError:
  ms_MensajeError = Err.Description
End Function


'debb-05/04/2011
'disminuye DEVOLUCIONES de Pacientes con Cuenta
Function DisminuyeDevolucionesDePacientes(oRecordset As Recordset, oConexion As Connection) As ADODB.Recordset
  Dim oRecordset1 As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  Dim lnCantidadDevuelta As Long, oRsTmp1 As New Recordset, lcSql As String
  If oRecordset.RecordCount = 0 Then
     Set DisminuyeDevolucionesDePacientes = oRecordset
  Else
    With oRecordset1
         .Fields.Append "tipo", adVarChar, 20, adFldIsNullable
         .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
         .Fields.Append "nombre", adVarChar, 300, adFldIsNullable
         .Fields.Append "cantidad", adInteger
         .Fields.Append "precio", adDouble
         .Fields.Append "total", adDouble
         .Fields.Append "idTipoFinanciamiento", adInteger
         .Fields.Append "idProducto", adInteger
         .Fields.Append "idFuenteFinanciamiento", adInteger
         .Fields.Append "idTipoFinanAtenciones", adInteger
         .Fields.Append "idServicioPaciente", adInteger
         .LockType = adLockOptimistic
         .Open
    End With
    oRecordset.MoveFirst
    Do While Not oRecordset.EOF

       With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "DisminuyeDevolucionesDePacientes"
          Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 50, oRecordset.Fields!movNumero): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 50, oRecordset.Fields!MovTipo): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRecordset.Fields!idProducto): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
          Set oRsTmp1.ActiveConnection = Nothing
       End With
       Set oCommand = Nothing
       Set oParameter = Nothing
              
       lnCantidadDevuelta = 0
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
             lnCantidadDevuelta = lnCantidadDevuelta + oRsTmp1.Fields!cantidadAdevolver
             oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       If (oRecordset.Fields!cantidad - lnCantidadDevuelta) <> 0 Then
            oRecordset1.AddNew
            oRecordset1.Fields!tipo = oRecordset.Fields!tipo
            oRecordset1.Fields!Codigo = oRecordset.Fields!Codigo
            oRecordset1.Fields!nombre = oRecordset.Fields!nombre
            oRecordset1.Fields!cantidad = oRecordset.Fields!cantidad - lnCantidadDevuelta
            oRecordset1.Fields!Precio = oRecordset.Fields!Precio
            oRecordset1.Fields!total = oRecordset.Fields!total - (lnCantidadDevuelta * oRecordset.Fields!Precio)
            oRecordset1.Fields!IdTipoFinanciamiento = oRecordset.Fields!IdTipoFinanciamiento
            oRecordset1.Fields!idProducto = oRecordset.Fields!idProducto
            oRecordset1.Fields!idFuenteFinanciamiento = oRecordset.Fields!idFuenteFinanciamiento
            oRecordset1.Fields!idTipoFinanAtenciones = oRecordset.Fields!idTipoFinanAtenciones
            oRecordset1.Fields!IdServicioPaciente = oRecordset.Fields!IdServicioPaciente
            oRecordset1.Update
       End If
       oRecordset.MoveNext
    Loop
    Set DisminuyeDevolucionesDePacientes = oRecordset1
  End If
  Set oRsTmp1 = Nothing
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaDevolucionesPorCuentaYServicio(idAtencion As Long, IdServicio As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Devoluciones de Medicamentos e Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaDevolucionesPorCuentaYServicio"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaDevolucionesPorCuentaYServicio = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaLaboratorioPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Laboratorio de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaLaboratorioPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaLaboratorioPorCuentaYServicioYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaImagenologiaPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Imageneologia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaImagenologiaPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaImagenologiaPorCuentaYServicioYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaInsumosPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaInsumosPorCuentaYServicioYFecha = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaFarmaciaPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Farmacia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaFarmaciaPorCuentaYServicioYFecha = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaProcedimientosPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaProcedimientosPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaProcedimientosPorCuentaYServicioYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaDevolucionesPorCuentaYServicioYFecha(idAtencion As Long, IdServicio As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaDevolucionesPorCuentaYServicioYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ServicioPaciente", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaDevolucionesPorCuentaYServicioYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaDevolucionesPorCuenta(idAtencion As Long) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Devoluciones de Medicamentos e Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaDevolucionesPorCuenta"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaDevolucionesPorCuenta = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaLaboratorioPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Laboratorio de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaLaboratorioPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaLaboratorioPorCuentaYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaImagenologiaPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Imageneologia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaImagenologiaPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaImagenologiaPorCuentaYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaInsumosPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaInsumosPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaInsumosPorCuentaYFecha = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaFarmaciaPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Farmacia de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaFarmaciaPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  'debb-05/04/2011
  Set SeleccionaFarmaciaPorCuentaYFecha = DisminuyeDevolucionesDePacientes(oRecordset, oConexion)
  '
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaProcedimientosPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de procedimientos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaProcedimientosPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaProcedimientosPorCuentaYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO FRANKLIN CACHAY 19/09/2013 - se agrego oConexion.CommandTimeout = 300
Function SeleccionaDevolucionesPorCuentaYFecha(idAtencion As Long, FI As Date, FF As Date) As ADODB.Recordset
  'Adams Bonilla Magallanes
  'Procedimiento para averiguar Consumos de Insumos de un Paciente por Cuenta y Servicio
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "SeleccionaDevolucionesPorCuentaYFecha"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FInicio", adDate, adParamInput, 0, FI): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FFinal", adDate, adParamInput, 0, FF): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set SeleccionaDevolucionesPorCuentaYFecha = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function ReporteSeleccionaTodoContenido(lcNombreTabla As String) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection

        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexionShape


        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "ReporteSeleccionaTodoContenido"
          Set oParameter = .CreateParameter("@NombreTabla", adVarChar, adParamInput, 500, lcNombreTabla): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With

        Set ReporteSeleccionaTodoContenido = oRecordset
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing '
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Sub ReporteEliminaTodoContenido(lcNombreTabla As String)

        Dim oRecordset As New ADODB.Recordset
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
          .CommandText = "ReporteEliminaTodoContenido"
          Set oParameter = .CreateParameter("@NombreTabla", adVarChar, adParamInput, 50, lcNombreTabla): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        
        oConexion.Close
        Set oRecordset = Nothing
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'Falta cambiar a store procedure
Function ReporteMuestraResultado() As Recordset
    Dim LcTexto1 As String, mRs_Cab As New Recordset
    LcTexto1 = " SHAPE {SELECT NroCuenta, Paciente FROM reporte_cabecera ORDER BY paciente} AS padre" & _
               " APPEND ({SELECT idUsuario, NroCuenta, ConsumoDescripcion, ConsumoImporte FROM reporte_detalle} AS hijo" & _
               " RELATE NroCuenta TO NroCuenta) AS hijo"
    mRs_Cab.Open LcTexto1, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
    Set ReporteMuestraResultado = mRs_Cab
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function CptHistoricosPorPaciente(lnIdPaciente As Long, lnIdMovimiento As Long) As Recordset

      On Error GoTo ManejadorDeError
      Dim oRecordset As New ADODB.Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
      Dim ms_MensajeError  As String
      Set CptHistoricosPorPaciente = Nothing
      ms_MensajeError = ""

      oConexion.CursorLocation = adUseClient
      oConexion.CommandTimeout = 300
      oConexion.Open sighentidades.CadenaConexion
    
      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CptHistoricosPorPacienteLab"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
      End With
      Set CptHistoricosPorPaciente = oRecordset
      oConexion.Close
      Set oRecordset = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarPorFechasRecalculo(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lbSoloParticulares As Boolean) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    ms_MensajeError = ""
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasRecalculo"
        Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set LabMovimientoLaboratorioSeleccionarPorFechasRecalculo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
 End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarPorFechasSoloExternosConBoletas(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, oConexion As Connection) As ADODB.Recordset
  
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Set LabMovimientoLaboratorioSeleccionarPorFechasSoloExternosConBoletas = Nothing
    ms_MensajeError = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasSoloExternosConBoletas"
        Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set LabMovimientoLaboratorioSeleccionarPorFechasSoloExternosConBoletas = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
     
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
 
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarPorFechasSoloConSeguro(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, oConexion As Connection) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Set LabMovimientoLaboratorioSeleccionarPorFechasSoloConSeguro = Nothing
    ms_MensajeError = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasSoloConSeguro"
        Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set LabMovimientoLaboratorioSeleccionarPorFechasSoloConSeguro = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
     
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarPorFechasSoloConCuentasPagantes(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, oConexion As Connection) As ADODB.Recordset
  
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Set LabMovimientoLaboratorioSeleccionarPorFechasSoloConCuentasPagantes = Nothing
    ms_MensajeError = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabMovimientoLaboratorioSeleccionarPorFechasSoloConCuentasPagantes"
        Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set LabMovimientoLaboratorioSeleccionarPorFechasSoloConCuentasPagantes = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
     
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LaboratorioDevuelveAlertaDeCantidades() As Recordset
  
  
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LaboratorioDevuelveAlertaDeCantidades = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LaboratorioDevuelveAlertaDeCantidades"
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LaboratorioDevuelveAlertaDeCantidades = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabItemsCptSeleccionarXfiltro(lcFiltro As String) As Recordset

  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabItemsCptSeleccionarXfiltro = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabItemsCptSeleccionarXfiltro"
    Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabItemsCptSeleccionarXfiltro = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabResultadosPorItemsSeleccionarXfiltro(lcFiltro As String) As Recordset

  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabResultadosPorItemsSeleccionarXfiltro = Nothing
  ms_MensajeError = ""

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadosPorItemsSeleccionarXfiltro"
    Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabResultadosPorItemsSeleccionarXfiltro = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabResultadosPorItemsActualizar(lnIdProductoCpt As Long, lnIdOrden As Long, oRsResultados As Recordset, _
                        lnIdRealizaAnalisis As Long, lnIdUsuario As Long, lcFecha As String, _
                        lnIdPaciente As Long, Optional oConexion1 As Connection) As Boolean
  On Error GoTo ErrRes
  LabResultadosPorItemsActualizar = False
  
  Dim lcGrupoSanguineo As String, lcFactorRh As String
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  lcGrupoSanguineo = "": lcFactorRh = ""
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
    .CommandText = "LabResultadosEliminarXidProductoIdOrden"
    Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProductoCpt): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    .Execute
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
  If oRsResultados.RecordCount > 0 Then
     oRsResultados.MoveFirst
     Do While Not oRsResultados.EOF
        With oCommand
          If UCase(oRsResultados!Item) = "GRUPO SANGUINEO" Then
             If Not IsNull(oRsResultados!valorTexto) Then
                lcGrupoSanguineo = Trim(oRsResultados!valorTexto)
             ElseIf Not IsNull(oRsResultados!ValorCombo) Then
                lcGrupoSanguineo = Trim(oRsResultados!ValorCombo)
             End If
          End If
          If UCase(oRsResultados!Item) = "FACTOR RH" Then
             If Not IsNull(oRsResultados!valorTexto) Then
                lcFactorRh = Trim(oRsResultados!valorTexto)
             ElseIf Not IsNull(oRsResultados!ValorCombo) Then
                lcFactorRh = Trim(oRsResultados!ValorCombo)
             End If
          End If
          .CommandType = adCmdStoredProc
          If oConexion1 Is Nothing Then
             Set .ActiveConnection = oConexion
          Else
             Set .ActiveConnection = oConexion1
          End If
          .CommandTimeout = 150
          .CommandText = "LabResultadosPorItemsActualizar"
          Set oParameter = .CreateParameter("@IdProductoCpt", adInteger, adParamInput, 0, lnIdProductoCpt): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ordenXresultado", adInteger, adParamInput, 0, oRsResultados.Fields!ordenXresultado): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorNumero", adCurrency, adParamInput, 0, oRsResultados.Fields!valorNumero): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 500, oRsResultados.Fields!valorTexto): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorCombo", adVarChar, adParamInput, 100, oRsResultados.Fields!ValorCombo): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorCheck", adVarChar, adParamInput, 1, oRsResultados.Fields!ValorCheck): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lnIdRealizaAnalisis", adInteger, adParamInput, 0, lnIdRealizaAnalisis): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lnIdUsuario", adInteger, adParamInput, 0, lnIdUsuario): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lcFecha", adDBTimeStamp, adParamInput, 0, CDate(lcFecha)): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, oRsResultados!IdGrupo): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, oRsResultados!idItem): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorReferencial", adVarChar, adParamInput, 100, oRsResultados!ValorReferencial): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@ValorMetodo", adVarChar, adParamInput, 50, oRsResultados!Metodo): .Parameters.Append oParameter
          .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        oRsResultados.MoveNext
     Loop
  End If
  
  If lnIdPaciente > 0 And (lcGrupoSanguineo <> "" Or lcFactorRh <> "") Then
        If oConexion1 Is Nothing Then
           ActualizaGrupoSanguineoYfactorRH lnIdPaciente, lcGrupoSanguineo, lcFactorRh, oConexion
        Else
           ActualizaGrupoSanguineoYfactorRH lnIdPaciente, lcGrupoSanguineo, lcFactorRh, oConexion1
        End If
  End If
  
 
  LabResultadosPorItemsActualizar = True
ErrRes:
  If LabResultadosPorItemsActualizar = False Then
     MsgBox Err.Number & " : " & Err.Description
     
  End If
  If oConexion1 Is Nothing Then
     oConexion.Close
  End If
  Set oRecordset = Nothing
  Set oConexion = Nothing
End Function

Sub ActualizaGrupoSanguineoYfactorRH(lnIdPaciente As Long, lcGrupoSanguineo As String, _
                                     lcFactorRh As String, oConexion As Connection)
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "PacientesActualizarGrupoSanYfactorRH"
          Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lcGrupoSanguineo", adVarChar, adParamInput, 50, lcGrupoSanguineo): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lcFactorRh", adVarChar, adParamInput, 50, lcFactorRh): .Parameters.Append oParameter
          .Execute
        End With
        Set oParameter = Nothing
        Set oCommand = Nothing

End Sub

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioSeleccionarXidOrden(lnIdOrden As Long) As Recordset

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Set LabMovimientoLaboratorioSeleccionarXidOrden = Nothing

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarXidOrden"
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioSeleccionarXidOrden = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoSeleccionarXidentificador(lnIdMovimiento As Long, Optional oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Set LabMovimientoSeleccionarXidentificador = Nothing

  If oConexion Is Nothing Then
    oConexion = New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  End If

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSeleccionarXidentificador"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabMovimientoSeleccionarXidentificador = oRecordset
  Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoIngresosXidMovimiento(lnIdMovimiento As Long, Optional oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter

  If oConexion Is Nothing Then
    oConexion = New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  End If

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoIngresosXidMovimiento"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabMovimientoIngresosXidMovimiento = oRecordset
  Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoSalidasXidMovimiento(lnIdMovimiento As Long, Optional oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter

  If oConexion Is Nothing Then
    oConexion = New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  End If

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoSalidasXidMovimiento"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabMovimientoSalidasXidMovimiento = oRecordset
  Set oCommand = Nothing

End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function ConstanciasSeleccionarXidentificador(lnIdConstancia As Long, Optional oConexion1 As Connection)

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  Set ConstanciasSeleccionarXidentificador = Nothing

  If oConexion1 Is Nothing Then
    Set oConexion = New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  Else
    Set oConexion = oConexion1
  End If

  With oCommand
    .CommandType = adCmdStoredProc
    If oConexion1 Is Nothing Then
       Set .ActiveConnection = oConexion
    Else
       Set .ActiveConnection = oConexion1
    End If
    .CommandTimeout = 150
    .CommandText = "ConstanciasSeleccionarXidentificador"
    Set oParameter = .CreateParameter("@IdConstancia", adInteger, adParamInput, 0, lnIdConstancia): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set ConstanciasSeleccionarXidentificador = oRecordset
  Set oCommand = Nothing
  
End Function

'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabMovimientoLaboratorioXidOrden(lnIdOrden As Long, oConexion As Connection) As Recordset

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  
  Set LabMovimientoLaboratorioXidOrden = Nothing

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioXidOrden"
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioXidOrden = oRecordset
  Set oCommand = Nothing
  
End Function


'MODIFICADO POR FRANKLIN CACHAY 20/09/2013 - se cambio a store procedure
Function LabResultadoParaSip200(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Set LabResultadoParaSip200 = Nothing

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadoParaSip200"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabResultadoParaSip200 = oRecordset
  Set oCommand = Nothing
  
End Function

'debb-18/08/2015
Function DevuelveDatosParaImpresionResultadoLaboratorio(lnIdOrden As Long) As String
    Dim rs As New Recordset
    Dim rs1 As New Recordset
    Dim rs2 As New Recordset
    Dim rs3 As New Recordset
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim oConexion As New Connection
    Dim oCommand As New ADODB.Command
    Dim oConexionExterna As New Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    DevuelveDatosParaImpresionResultadoLaboratorio = ""
    'Set rs = mo_ReglasFacturacion.FactOrdenServicioSeleccionarXidOrden(lnIdOrden)
    Set rs = mo_ReglasAdmision.FactOrdenServicioSeleccionarPorIdOrden(lnIdOrden, oConexion)
    If rs.RecordCount > 0 Then
        If Not IsNull(rs!IdServicioPaciente) And (rs!idTipoServicio = 1 Or rs!idTipoServicio = 2 Or rs!idTipoServicio = 3) Then
           DevuelveDatosParaImpresionResultadoLaboratorio = " (S.Act: " & mo_ReglasFacturacion.BuscaServicioActualDelPaciente(rs!IdServicioPaciente) & ")"
        ElseIf rs!IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal Then
           Set rs1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(lnIdOrden, oConexion)
           If rs1.RecordCount > 0 Then
              If Not IsNull(rs1!IdComprobantePago) Then     'debb-16/01/2016
                    Set rs2 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(rs1!IdComprobantePago, oConexion)
                    If rs2.RecordCount > 0 Then
                       If Not IsNull(rs2!observaciones) Then
                          DevuelveDatosParaImpresionResultadoLaboratorio = " (Part. " & Trim(rs2!observaciones) & ")"
                       End If
                    End If
                    rs2.Close
              End If
           End If
           rs1.Close
        ElseIf rs!IdTipoFinanciamiento = sghTipoFinanciamiento.sghSis Then
             Set rs3 = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarPorIdCuenta(rs!idCuentaAtencion)
             If rs3.RecordCount > 0 Then
                DevuelveDatosParaImpresionResultadoLaboratorio = " (PacExt-SIS)"
                If Not IsNull(rs3!idSiaSis) And Not IsNull(rs3!sisCodigo) Then
                     oConexionExterna.CommandTimeout = 300
                     oConexionExterna.CursorLocation = adUseClient
                     oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
                     With oCommand
                         .CommandType = adCmdStoredProc
                         Set .ActiveConnection = oConexionExterna
                         .CommandTimeout = 150
                         .CommandText = "SisFiliacionesSeleccionarPorIdSiaSis"
                         Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, rs3!idSiaSis): .Parameters.Append oParameter
                         Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 2, rs3!sisCodigo):   .Parameters.Append oParameter
                         Set rs1 = .Execute
                    End With
                    Set oParameter = Nothing
                    Set oCommand = Nothing
                    If rs1.RecordCount > 0 Then
                       If Not IsNull(rs1!CodigoEstablAdscripcion) Then
                             With oCommand
                                 .CommandType = adCmdStoredProc
                                 Set .ActiveConnection = oConexion
                                 .CommandTimeout = 150
                                 .CommandText = "EstablecimientosSeleccionarTodoCamposXCodigo"
                                 Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 6, Right(rs1!CodigoEstablAdscripcion, 5)):  .Parameters.Append oParameter
                                 Set rs2 = .Execute
                            End With
                            Set oParameter = Nothing
                            Set oCommand = Nothing
                            If rs2.RecordCount > 0 Then
                               DevuelveDatosParaImpresionResultadoLaboratorio = " (PacExt-SIS adscrito: " & Trim(rs2!nombre) & ")"
                            End If
                            rs2.Close
                       End If
                    End If
                    rs1.Close
                    oConexionExterna.Close
               End If
            End If
            rs3.Close
        Else
            Set rs1 = mo_reglasComunes.FuentesFinanciamientoSegunFiltro("idFuenteFinanciamiento=" & rs!idFuenteFinanciamiento)
            If rs1.RecordCount > 0 Then
               DevuelveDatosParaImpresionResultadoLaboratorio = " (PacExt-Fte.Finan: " & Trim(rs1!Descripcion) & ")"
            End If
            rs1.Close
        End If
    End If
    oConexion.Close
    rs.Close
    
    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasCaja = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oConexionExterna = Nothing
End Function

'debb-17/08/2015
Function labResultadoXfechas(ldFechaEgresoIni As Date, ldFechaEgresoFin As Date, lbSoloGestantes As Boolean) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oRecordset1 As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Dim lnFor As Integer
  ms_MensajeError = ""
  With oRecordset
        .Fields.Append "fecha", adDate
        .Fields.Append "Paciente", adVarChar, 60, adFldIsNullable
        .Fields.Append "cpt", adVarChar, 20, adFldIsNullable
        .Fields.Append "cptDescripcion", adVarChar, 130, adFldIsNullable
        .Fields.Append "Servicio", adVarChar, 50, adFldIsNullable
        .Fields.Append "tServicio", adVarChar, 50, adFldIsNullable
        .Fields.Append "nroSerie", adVarChar, 4, adFldIsNullable
        .Fields.Append "nroDocumento", adVarChar, 20, adFldIsNullable
        .Fields.Append "NroMovimiento", adInteger
        .Fields.Append "Responsable", adVarChar, 150, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
        .Sort = "fecha"
  End With
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  For lnFor = 1 To 2
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          If lnFor = 1 Then
             .CommandText = "labResultadoXfechas"
          Else
             .CommandText = "labResultadoItemsXfechas"
          End If
          Set oParameter = .CreateParameter("@FechaEgresoIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@FechaEgresoFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
          Set oRecordset1 = .Execute
          
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        If lbSoloGestantes = True Then
           oRecordset1.Filter = "Eo_EG>0"
        End If
        If oRecordset1.RecordCount > 0 Then
           oRecordset1.MoveFirst
           Do While Not oRecordset1.EOF
              oRecordset.AddNew
              oRecordset.Fields!fecha = oRecordset1.Fields!fecha
              oRecordset.Fields!Paciente = Left(oRecordset1.Fields!Paciente, 60)
              oRecordset.Fields!cpt = oRecordset1.Fields!cpt
              oRecordset.Fields!cptDescripcion = Left(oRecordset1.Fields!cptDescripcion, 130)
              oRecordset.Fields!Servicio = Left(oRecordset1.Fields!Servicio, 50)
              oRecordset.Fields!tServicio = oRecordset1.Fields!tServicio
              oRecordset.Fields!NroSerie = oRecordset1.Fields!NroSerie
              oRecordset.Fields!NroDocumento = oRecordset1.Fields!NroDocumento
              oRecordset.Fields!nroMovimiento = IIf(IsNull(oRecordset1.Fields!IdMovimiento), 0, oRecordset1.Fields!IdMovimiento)
              oRecordset!responsable = IIf(IsNull(oRecordset1!responsable), "", oRecordset1!responsable)
              oRecordset.Update
              oRecordset1.MoveNext
           Loop
        End If
        oRecordset1.Close
  Next
  Set labResultadoXfechas = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oRecordset1 = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
  Resume
  
End Function


'debb-31/03/2016
Function LabPruebasSeleccionarPorCpt(lcCodigoCpt As String, oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Set LabPruebasSeleccionarPorCpt = Nothing
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabPruebasSeleccionarPorCpt"
    Set oParameter = .CreateParameter("@codigoCPT", adVarChar, adParamInput, 20, lcCodigoCpt): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabPruebasSeleccionarPorCpt = oRecordset
  Set oCommand = Nothing
End Function


'debb-09/09/2016
Function LabMovimientoLaboratorioSeleccionarPorIdPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorIdPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabMovimientoLaboratorioSeleccionarPorIdPaciente = oRecordset
  Set oCommand = Nothing
  
End Function

'debb-09/09/2016
Function LabMovimientoCPTXcuenta(lnIdCuenta As Long, oConexion As Connection) As Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoCPTXcuenta"
    Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set LabMovimientoCPTXcuenta = oRecordset
  Set oCommand = Nothing
  
End Function

'debb-08/03/2017
Function ServiciosDondeEstuvoElPacienteCE(idAtencion As Long) As ADODB.Recordset
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
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "ServiciosDondeEstuvoElPacienteCE"
    Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set ServiciosDondeEstuvoElPacienteCE = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function



Function LabResultadoPorItemsSeleccionarPorPaciente(lnIdPaciente As Long, lnIdItem As Long, oConexion As Connection, lnIdOrden As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  ms_MensajeError = ""
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadoPorItemsSeleccionarPorPaciente"
    Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set oCommand = Nothing
  Set oParameter = Nothing
  If oRecordset.RecordCount = 0 Then
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "LabResultadoPorItemsXidOrdenXitem"
          Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, lnIdItem): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
  End If
  Set LabResultadoPorItemsSeleccionarPorPaciente = oRecordset
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Function LabResultadosPorItemsSeleccionarPORfiltro(lcFiltro As String) As Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError  As String
  Set LabResultadosPorItemsSeleccionarPORfiltro = Nothing
  ms_MensajeError = ""
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadosPorItemsSeleccionarPORfiltro"
    Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabResultadosPorItemsSeleccionarPORfiltro = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Sub ResultadosAutomaticosActualizaHaciaGalenhos(lnIdOrden As Long)
    Dim lcMensajeLicencia As String
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    
    
    On Error GoTo ErrResAutHacGal
    'Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oRsLabItems As New Recordset
    Dim oRsLabItemsGrupos As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsResultados As New Recordset
    Dim oConexion As New Connection
    Dim oConexionExterna As New Connection
    Dim lcSql As String, lnItems As Long, lnGrupoItems As Long, lnId1 As Long, lcCpt1 As String, lnIdProductoCpt As Long
    Dim lnIdEmpleadoResultado As Long, lcFechaResultado As String, lnIdItem As Long, lnIdItemGrupo As Long, lnNumero As Long
    Dim lbTodoOK As Boolean, lbEncontroDato As Boolean, lnIdPaciente99 As Long
    
    oConexion.CommandTimeout = 300: oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexionExterna.CommandTimeout = 300: oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    lcSql = "select * from InterLabDetalle where idEstado=" & sghInterLabEstados.sghIntTieneResultadoLab & _
            IIf(lnIdOrden > 0, " and identificador=" & lnIdOrden, "") & _
            " order by identificador,cptCodigo"
    oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       '
'       If  False Then   'licencia
'           GoTo ErrResAutHacGal
'       End If
       '
       lbEncontroDato = True
       oConexion.BeginTrans
       oConexionExterna.BeginTrans
       
       lnItems = 1
       lcSql = "select * from LabItems order by idItem desc"
       oRsLabItems.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsLabItems.RecordCount > 0 Then
          lnItems = oRsLabItems!idItem + 1
       End If
       
       lnGrupoItems = 1
       lcSql = "select * from LabItemsGrupos order by idItemGrupo desc"
       oRsLabItemsGrupos.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       If oRsLabItemsGrupos.RecordCount > 0 Then
          lnGrupoItems = oRsLabItemsGrupos!idItemGrupo + 1
       End If
       
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If oRsResultados.State = 1 Then oRsResultados.Close
          With oRsResultados
                  .Fields.Append "SoloNumero", adBoolean
                  .Fields.Append "SoloTexto", adBoolean
                  .Fields.Append "SoloCombo", adBoolean
                  .Fields.Append "SoloCheck", adBoolean
                  .Fields.Append "ordenXresultado", adInteger, 4, adFldIsNullable
                  .Fields.Append "Grupo", adVarChar, 100, adFldIsNullable
                  .Fields.Append "Item", adVarChar, 100, adFldIsNullable
                  .Fields.Append "idItem", adInteger
                  .Fields.Append "ValorNumero", adDouble
                  .Fields.Append "ValorTexto", adVarChar, 500, adFldIsNullable
                  .Fields.Append "ValorCombo", adVarChar, 100, adFldIsNullable
                  .Fields.Append "ValorCheck", adVarChar, 1, adFldIsNullable
                  .Fields.Append "ValorReferencial", adVarChar, 100, adFldIsNullable
                  .Fields.Append "Metodo", adVarChar, 50, adFldIsNullable
                  .Fields.Append "idGrupo", adInteger
                  .LockType = adLockOptimistic
                  .Open
          End With
          
          lnId1 = oRsTmp1!identificador
          lcCpt1 = oRsTmp1!cptCodigo
          
          lnIdProductoCpt = 0
          Set oRsTmp2 = mo_reglasComunes.CatalogoServiciosSeleccionarPorCodigo(lcCpt1, oConexion)
          If oRsTmp2.RecordCount > 0 Then
             lnIdProductoCpt = oRsTmp2!idProducto
          End If
          oRsTmp2.Close
          
          lnIdPaciente99 = 0
          Set oRsTmp2 = mo_AdminAdmision.FactOrdenServicioSeleccionarPorIdOrden(lnIdOrden, oConexion)
          If oRsTmp2.RecordCount > 0 Then
             If Not IsNull(oRsTmp2!idPaciente) Then
                lnIdPaciente99 = oRsTmp2!idPaciente
             End If
          End If
          oRsTmp2.Close
          
          lnIdEmpleadoResultado = 0
          Set oRsTmp2 = mo_reglasComunes.EmpleadosSeleccionarPorDNI(oRsTmp1!DNI)
          If oRsTmp2.RecordCount > 0 Then
             lnIdEmpleadoResultado = oRsTmp2!idEmpleado
          Else
             MsgBox "No existe DNI del EMPLEADO que realiza RESULTADO", vbInformation, ""
             Exit Sub
          End If
          oRsTmp2.Close
          
          lcFechaResultado = Format(oRsTmp1!fechaHoraResultado, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
          lnNumero = 1
          Do While Not oRsTmp1.EOF And lnId1 = oRsTmp1!identificador And lcCpt1 = oRsTmp1!cptCodigo
                oRsLabItems.Filter = "item='" & oRsTmp1!Item & "'"
                If oRsLabItems.RecordCount = 0 Then
                   oRsLabItems.AddNew
                   oRsLabItems!idItem = lnItems
                   oRsLabItems!Item = oRsTmp1!Item
                   oRsLabItems.Update
                   lnIdItem = lnItems
                   lnItems = lnItems + 1
                Else
                   lnIdItem = oRsLabItems!idItem
                End If
                
                oRsLabItemsGrupos.Filter = "grupo='" & oRsTmp1!Grupo & "'"
                If oRsLabItemsGrupos.RecordCount = 0 Then
                   oRsLabItemsGrupos.AddNew
                   oRsLabItemsGrupos!idItemGrupo = lnGrupoItems
                   oRsLabItemsGrupos!Grupo = oRsTmp1!Grupo
                   oRsLabItemsGrupos.Update
                   lnIdItemGrupo = lnGrupoItems
                   lnGrupoItems = lnGrupoItems + 1
                Else
                   lnIdItemGrupo = oRsLabItemsGrupos!idItemGrupo
                End If
                
                oRsResultados.AddNew
                oRsResultados!SoloNumero = False
                oRsResultados!SoloTexto = True
                oRsResultados!SoloCombo = False
                oRsResultados!SoloCheck = False
                oRsResultados!ordenXresultado = lnNumero
                oRsResultados!Grupo = "x"
                oRsResultados!Item = "x"
                oRsResultados!idItem = lnIdItem
                'oRsResultados!ValorNumero = 0
                oRsResultados!valorTexto = oRsTmp1!itemResultado
               ' oRsResultados!ValorCombo = ""
               ' oRsResultados!ValorCheck = ""
                oRsResultados!ValorReferencial = oRsTmp1!itemReferencia
                oRsResultados!Metodo = oRsTmp1!itemMetodo
                oRsResultados!IdGrupo = lnIdItemGrupo
                oRsResultados.Update
                
                lnNumero = lnNumero + 1
                
                InterLabDetalleActualizar oConexionExterna, lnId1, lcCpt1, sghInterLabEstados.sghIntTieneResultadoLeidoEnGalenhos

                
                oRsTmp1.MoveNext
                If oRsTmp1.EOF Then
                   Exit Do
                End If
          Loop
          If LabResultadosPorItemsActualizar(lnIdProductoCpt, lnId1, oRsResultados, _
                                 lnIdEmpleadoResultado, sighentidades.Usuario, lcFechaResultado, _
                                 lnIdPaciente99, oConexion) = True Then
                                 
          End If
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
        oRsLabItems.Close
        oRsLabItemsGrupos.Close
    Else
'        oConexion.RollbackTrans
'        oConexionExterna.RollbackTrans
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsResultados = Nothing
    Set oConexion = Nothing
    Set oConexionExterna = Nothing
    Set oRsLabItemsGrupos = Nothing
    Set oRsLabItems = Nothing
    Set mo_reglasComunes = Nothing
    Exit Sub
    Resume
End Sub



Sub Imprimir_LabResultadosItems(rsTmp As Recordset, lcEdadEnAtencion As String, lcHistoria As String, _
                                                                   lcServicioActualPaciente As String)
        Dim lbHalloUnCPT As Boolean
        Dim oRsResultados As New Recordset
        Dim oRsResultadosCPT As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
        Dim lcSql As String, lcTexto As String
        Dim lnIdProducto As Long, lnIdItem As Long, ldFechaResultado As Date, lnCantidadCPT As Integer
        Dim lcPaciente As String, ml_idTipoSexo As Long, ml_FechaNacimiento As Date, lcMedico As String
        lbHalloUnCPT = False
        If rsTmp.RecordCount > 0 Then
            With oRsResultados
                  .Fields.Append "idProducto", adInteger
                  .Fields.Append "SoloNumero", adBoolean
                  .Fields.Append "SoloTexto", adBoolean
                  .Fields.Append "SoloCombo", adBoolean
                  .Fields.Append "SoloCheck", adBoolean
                  .Fields.Append "ordenXresultado", adInteger, 4, adFldIsNullable
                  .Fields.Append "Grupo", adVarChar, 100, adFldIsNullable
                  .Fields.Append "Item", adVarChar, 100, adFldIsNullable
                  .Fields.Append "idItem", adInteger
                  .Fields.Append "ValorNumero", adDouble
                  .Fields.Append "ValorTexto", adVarChar, 500, adFldIsNullable
                  .Fields.Append "ValorCombo", adVarChar, 100, adFldIsNullable
                  .Fields.Append "ValorCheck", adVarChar, 1, adFldIsNullable
                  .Fields.Append "ValorReferencial", adVarChar, 100, adFldIsNullable
                  .Fields.Append "Metodo", adVarChar, 100, adFldIsNullable
                  .LockType = adLockOptimistic
                  .Open
           End With

           rsTmp.MoveFirst
           Do While Not rsTmp.EOF
              Set oRsTmp1 = LabResultadosPorItemsSeleccionarPORfiltro("dbo.LabResultadoPorItems.idOrden=" & rsTmp("idOrden") & _
                                                 " and dbo.LabResultadoPorItems.idProductoCpt=" & rsTmp!idProducto)
              If oRsTmp1.RecordCount > 0 Then
                   If IsNull(oRsTmp1!realizaAnalisis) Then
                      MsgBox "No grabó bien PERSONA DE LABORATORIO que realiza el RESULTADO", vbInformation, ""
                      Exit Sub
                   ElseIf oRsTmp1!realizaAnalisis <= 0 Then
                      MsgBox "No grabó bien PERSONA DE LABORATORIO que realiza el RESULTADO", vbInformation, ""
                      Exit Sub
                   End If
              
                   oRsTmp1.MoveFirst
                   If lbHalloUnCPT = False Then
                        'Medico que ordena
                        Set oRsTmp2 = LabMovimientoLaboratorioSeleccionarXidOrden(rsTmp("idOrden"))
                        lcPaciente = "": ml_idTipoSexo = 0: ml_FechaNacimiento = 0: lcMedico = ""
                        If oRsTmp2.RecordCount > 0 Then
                           ml_idTipoSexo = IIf(IsNull(oRsTmp2.Fields!idTipoSexo), 1, oRsTmp2.Fields!idTipoSexo)
                           ml_FechaNacimiento = IIf(IsNull(oRsTmp2.Fields!FechaNacimiento), Date, oRsTmp2.Fields!FechaNacimiento)
                           If lcEdadEnAtencion = "" Then
                              lcEdadEnAtencion = "F.Nacim: " & ml_FechaNacimiento
                           End If
                           lcMedico = oRsTmp2.Fields!OrdenaPrueba
                           lcPaciente = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(lcHistoria, False) & ")" & oRsTmp2.Fields!Paciente
                        End If
                    End If
                    lbHalloUnCPT = True
                    'Llena Temporal
                    oRsResultados.AddNew
                    oRsResultados.Fields!idProducto = rsTmp!idProducto
                    oRsResultados.Fields!valorTexto = "ANALISIS: " & rsTmp!Codigo & " - " & Trim(rsTmp!nombre)
                    oRsResultados.Fields!Metodo = IIf(rsTmp!resultadoAutomatico = True, " <<ra>>", " ")
                    oRsResultados.Update
                    oRsResultados.AddNew
                    oRsResultados.Fields!idProducto = rsTmp!idProducto
                    oRsResultados.Fields!valorTexto = "Personal q realizó la prueba: " & mo_reglasComunes.EmpleadosDevuelveNombre(oRsTmp1!realizaAnalisis)
                    oRsResultados.Update
                    oRsResultados.AddNew
                    oRsResultados.Fields!idProducto = rsTmp!idProducto
                    oRsResultados.Fields!valorTexto = "F.resultado: " & oRsTmp1!fecha
                    oRsResultados.Update
                    Do While Not oRsTmp1.EOF
                        oRsResultados.AddNew
                        oRsResultados.Fields!idProducto = rsTmp!idProducto
                        oRsResultados.Fields!ordenXresultado = oRsTmp1!ordenXresultado
                        oRsResultados.Fields!Grupo = oRsTmp1!nombreGrupo
                        oRsResultados.Fields!Item = oRsTmp1!Item
                        oRsResultados.Fields!idItem = oRsTmp1!idItem
                        oRsResultados.Fields!ValorReferencial = oRsTmp1!ValorReferencial
                        oRsResultados.Fields!Metodo = oRsTmp1!Metodo
                        If Not IsNull(oRsTmp1.Fields!valorNumero) Then
                           oRsResultados.Fields!valorNumero = oRsTmp1.Fields!valorNumero
                           oRsResultados.Fields!SoloNumero = True
                        End If
                        If Not IsNull(oRsTmp1.Fields!valorTexto) Then
                           oRsResultados.Fields!valorTexto = oRsTmp1.Fields!valorTexto
                           oRsResultados.Fields!SoloTexto = True
                        End If
                        If Not IsNull(oRsTmp1.Fields!ValorCombo) Then
                           oRsResultados.Fields!ValorCombo = oRsTmp1.Fields!ValorCombo
                           oRsResultados.Fields!SoloCombo = True
                        End If
                        If Not IsNull(oRsTmp1.Fields!ValorCheck) Then
                           oRsResultados.Fields!ValorCheck = oRsTmp1.Fields!ValorCheck
                           oRsResultados.Fields!SoloCheck = True
                        End If
                        oRsResultados.Update
                        oRsTmp1.MoveNext
                    Loop
              End If
              rsTmp.MoveNext
           Loop

           
        End If
        If lbHalloUnCPT = False Then
           Exit Sub
        End If
        '
        Dim iFila As Long, iColumna As Integer
        Dim lbEsOpenOffice As Boolean
        
        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
        On Error GoTo ManejadorError
        
        If lbEsOpenOffice = True Then
            Dim ServiceManager As Object
            Dim Desktop As Object
            Dim Document As Object
            Dim Feuille As Object
            Dim Plage As Object
            Dim args()
            Dim Chemin As String
            Dim Fichier As String
            Dim lcArchivoExcel As String
            Dim PrintArea(0)
            Dim Style As Object
            Dim Border As Object
            'encabezado
            Dim PageStyles As Object
            Dim Sheet As Object
            Dim StyleFamilies As Object
            Dim DefPage As Object
            Dim Htext As Object
            Dim Hcontent As Object
            Dim lnHwnd As Long
            Dim ret As Long
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
        End If
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\LabResultadoXitem.ods"
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_reglasComunes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LabResultadoXitem.xls")
            
            oWorkBookPlantilla.Worksheets("Hoja1").COPY Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_reglasComunes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(3, 0).setFormula(lcPaciente & _
                                       " (Sexo: " & IIf(ml_idTipoSexo = 1, "Masculino", "Femenino") & ") ( " & _
                                       lcEdadEnAtencion & ")  " & lcServicioActualPaciente)
            Call Feuille.getcellbyposition(3, 1).setFormula(lcMedico)
            Call Feuille.getcellbyposition(5, 1).setFormula("")
            Call Feuille.getcellbyposition(2, 2).setFormula("")
        Else
            oWorkSheet.Cells(1, 4).value = lcPaciente & _
                                           " (Sexo: " & IIf(ml_idTipoSexo = 1, "Masculino", "Femenino") & ") (" & _
                                           lcEdadEnAtencion & ")  " & lcServicioActualPaciente
            oWorkSheet.Cells(2, 4).value = lcMedico
            oWorkSheet.Cells(2, 6).value = ""
            oWorkSheet.Cells(3, 3).value = " "
        End If
        iFila = 9
        lnCantidadCPT = 0
        If oRsResultados.RecordCount > 0 Then
           oRsResultados.MoveFirst
           Do While Not oRsResultados.EOF
              lnCantidadCPT = lnCantidadCPT + 1
              lnIdProducto = oRsResultados.Fields!idProducto
              Do While Not oRsResultados.EOF And lnIdProducto = oRsResultados.Fields!idProducto
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(3, 0).setFormula(IIf(IsNull(oRsResultados.Fields!Grupo), "", oRsResultados.Fields!Grupo))
                        Call Feuille.getcellbyposition(3, 0).setFormula(IIf(IsNull(oRsResultados.Fields!Item), "", oRsResultados.Fields!Item))
                    Else
                        oWorkSheet.Cells(iFila, 2).value = oRsResultados.Fields!Grupo
                        oWorkSheet.Cells(iFila, 3).value = oRsResultados.Fields!Item
                    End If
                    
                    lcTexto = ""
                    If oRsResultados.Fields!valorNumero > 0 Then
                      lcTexto = lcTexto & Trim(Str(IIf(IsNull(oRsResultados.Fields!valorNumero), "", oRsResultados.Fields!valorNumero))) & "| "
                    End If
                    If Len(Trim(oRsResultados.Fields!valorTexto)) > 0 Then
                      lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!valorTexto), "", oRsResultados.Fields!valorTexto)) & "| "
                    End If
                    If Len(Trim(oRsResultados.Fields!ValorCombo)) > 0 Then
                      lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!ValorCombo), "", oRsResultados.Fields!ValorCombo)) & "| "
                    End If
                    If Not IsNull(oRsResultados.Fields!ValorCheck) Then
                      lcTexto = lcTexto & IIf(oRsResultados.Fields!ValorCheck = True, "x", "")
                    End If
                    If lcTexto <> "" Then
                        lcTexto = Trim(lcTexto)
                        If Right(lcTexto, 1) = "|" Then
                            lcTexto = Left(lcTexto, Len(lcTexto) - 1)
                        End If
                    
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcTexto)
                        Else
                            oWorkSheet.Cells(iFila, 4).value = lcTexto
                        End If
                    End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!ValorReferencial), "", oRsResultados.Fields!ValorReferencial))
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!Metodo), "", oRsResultados.Fields!Metodo))
                    Else
                        oWorkSheet.Cells(iFila, 8).value = oRsResultados.Fields!ValorReferencial
                        oWorkSheet.Cells(iFila, 9).value = oRsResultados.Fields!Metodo
                    End If
                    iFila = iFila + 1
                    oRsResultados.MoveNext
                    If oRsResultados.EOF Then
                       Exit Do
                    End If
              Loop
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("*************************************")
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula("*************************************")
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula("*************************************")
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula("*************************************************************************************")
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula("*************************************")
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula("*************************************")
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula("**************************************************************************")
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula("*************************************")
                Else
                    
                    oWorkSheet.Cells(iFila, 2).value = "*************************************"
                    oWorkSheet.Cells(iFila, 3).value = "*************************************"
                    oWorkSheet.Cells(iFila, 4).value = "*************************************"
                    oWorkSheet.Cells(iFila, 5).value = "*************************************************************************************"
                    oWorkSheet.Cells(iFila, 6).value = "*************************************"
                    oWorkSheet.Cells(iFila, 7).value = "*************************************"
                    oWorkSheet.Cells(iFila, 8).value = "**************************************************************************"
                    oWorkSheet.Cells(iFila, 9).value = "*************************************"
                    
                End If
              iFila = iFila + 2
           Loop
        End If
        If lbEsOpenOffice = True Then
        Else
            oWorkSheet.Cells(iFila, 2).value = "Digitador: " & InicialesDelDigitador
            oWorkSheet.Range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila + 2, 100)).Select
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 9
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El reporte se generó con exito: " & lcArchivoExcel, vbInformation, "Reporte"
        Else
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               
               oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
               
            End If
            oExcel.Visible = True
            If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) <> 7637 Then                               'huaral
                If lnCantidadCPT = 1 Then
                    oWorkSheet.PageSetup.Zoom = 73
                    oWorkSheet.PageSetup.Orientation = xlLandscape
                    oWorkSheet.PrintPreview
                    oExcel.DisplayAlerts = False
                    oExcel.Quit
                Else
                    oWorkSheet.PrintPreview
                End If
            Else
               ' oWorkSheet.PrintPreview
            End If
            
            
            


        End If
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
        Set oRsResultados = Nothing
        Set oRsResultadosCPT = Nothing
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set oExcel = Nothing
        Set mo_reglasComunes = Nothing
        Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub


Sub LlenaItemsConResultadosParaImpresion(ByRef rsTmp As Recordset, lnIdOrden As Long, ByRef lcEdadEnAtencion As String, _
                                         ByRef lcHistoria As String, ByRef lcServicioActualPaciente As String, lcCpt As String)
    Dim rs As New Recordset
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim ml_puntoCarga As Long, lnIdMovimiento As Long
    If rsTmp.State = 1 Then rsTmp.Close
    With rsTmp
      .Fields.Append "Imprime", adBoolean
      .Fields.Append "NroOrden", adDouble
      .Fields.Append "idOrden", adDouble
      .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
      .Fields.Append "Nombre", adVarChar, 50, adFldIsNullable
      .Fields.Append "idProducto", adDouble
      .Fields.Append "Cantidad", adInteger
      .Fields.Append "Precio", adDouble
      .Fields.Append "Total", adDouble
      .Fields.Append "idPuntoCarga", adDouble
      .Fields.Append "Resultado", adVarChar, 2, adFldIsNullable
      .Fields.Append "ResultadoAutomatico", adBoolean
      .Fields.Append "ObsReceta", adVarChar, 300, adFldIsNullable
      .LockType = adLockOptimistic
      .Open
    End With
    lcEdadEnAtencion = "": lcServicioActualPaciente = "": lcHistoria = "": ml_puntoCarga = 0: lnIdMovimiento = 0
    Set rs = LabMovimientoLaboratorioPORidOrden(lnIdOrden)
    If rs.RecordCount = 0 Then
       Exit Sub
    Else
       lcServicioActualPaciente = rs!nombre
       lcHistoria = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rs!NroHistoriaClinica)), False)
       
       Dim oEdad As Edad
       oEdad = sighentidades.CalcularEdad(rs!FechaNacimiento, rs!fecha)
       lcEdadEnAtencion = oEdad.Edad & " " & oEdad.NombreEdad
       ml_puntoCarga = rs!IdPuntoCarga
       lnIdMovimiento = rs!IdMovimiento
    End If
    rs.Close
    
    Set rs = mo_ReglasFacturacion.LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga(lnIdOrden, ml_puntoCarga, lnIdMovimiento)  'Set rs = mo_ReglasFacturacion.FacturacionServicioDespachoFiltraPorIdOrden(ml_IdRegistroSeleccionado)
    rs.Filter = "codigo='" & lcCpt & "'"
    If rs.RecordCount > 0 Then
  
      Dim Tot As Double
      Dim TotP As Integer
      Dim TotRes As Integer
      Dim T As Integer
      
      Tot = 0: TotP = 0: TotRes = 0: T = 0
      If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
              If PruebaTieneResultado(rs!Codigo, rs!IdOrden) = True Then
                  T = T + 1
                  rsTmp.AddNew
                  rsTmp!NroOrden = T
                  rsTmp!IdOrden = rs!IdOrden
                  rsTmp!idProducto = rs!idProducto
                  rsTmp!cantidad = rs!cantidad
                  rsTmp!Precio = rs!Precio
                  rsTmp!total = rs!total
                  rsTmp!Codigo = rs!Codigo
                  rsTmp!nombre = Left(rs!nombre, 50)
                  rsTmp!IdPuntoCarga = rs!IdPuntoCarga
                  rsTmp!Resultado = "SI"
                  rsTmp!Imprime = True
                  If Not IsNull(rs!resultadoAutomatico) Then
                     If rs!resultadoAutomatico = 1 Then
                        rsTmp!resultadoAutomatico = True
                     End If
                  End If
                  rsTmp.Update
              End If
              rs.MoveNext
            Loop
       End If
   End If
   Set rs = Nothing
   Set mo_ReglasFacturacion = Nothing
End Sub


Function LabMovimientoLaboratorioPORidOrden(lnIdOrden As Long) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabMovimientoLaboratorioPORidOrden"
    Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set LabMovimientoLaboratorioPORidOrden = oRecordset
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function




Function InterLabDetalleActualizar(oConexionExterna As Connection, lnIdOrden As Long, lcCodigoCpt As String, lnIdEstado As Long) As ADODB.Recordset
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
    .CommandText = "InterLabDetalleActualizar"
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

Function UsaNuevaVentanaResultadosLaboratorio(lnIdOrden As Long, lcCodigoCpt As String) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Dim ms_MensajeError As String
  Dim lnIdServicioDelPaciente As Long
  
  ms_MensajeError = ""
  
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "LabResultadoSeleccionarPorIdporCpt"
    Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, lcCodigoCpt): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  UsaNuevaVentanaResultadosLaboratorio = True
  If oRecordset.RecordCount > 0 Then
     UsaNuevaVentanaResultadosLaboratorio = False
  End If
  oRecordset.Close
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function


Sub DevuelveHemoglobinaDeLaAtencion(ByRef ldFechaHemoglobina As Date, ByRef lcHemoglobina As String, _
                                    lnIdPaciente As Long, lnIdAtencion As Long, ldFechaIngreso As Date, oConexion As Connection)
    Dim oRsTmp987 As New Recordset
    ldFechaHemoglobina = 0
    lcHemoglobina = ""
    Set oRsTmp987 = LabResultadoPorItemsSeleccionarPorPaciente(lnIdPaciente, 1, oConexion, _
                                                               sighentidades.DevuelveNroOrdenParaHemoglobina(lnIdAtencion))
    If oRsTmp987.RecordCount > 0 Then
       oRsTmp987.MoveFirst
       oRsTmp987.Find "fecha<='" & ldFechaIngreso & "'"
       If Not oRsTmp987.EOF Then
             ldFechaHemoglobina = oRsTmp987!fecha
             If Not IsNull(oRsTmp987!valorTexto) Then
                lcHemoglobina = oRsTmp987!valorTexto
             ElseIf Not IsNull(oRsTmp987!valorNumero) Then
                lcHemoglobina = Trim(Str(oRsTmp987!valorNumero))
             End If
       End If
    End If
    oRsTmp987.Close
    Set oRsTmp987 = Nothing
End Sub
