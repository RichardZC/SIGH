VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para CAJA
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
'WCG-05 Ini

Option Explicit
Dim ms_MensajeError As String

Const ID_PAGADO = 6
Const ID_PENDIENTE_PAGO = 2
Dim oRsDetalleProductosConLotes As New Recordset
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


Function TiposComprobanteSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaTiposComprobante
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposComprobanteSeleccionarTodos = oTabla.SeleccionarTodos()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function

'Function BuscarCoincidenciasEnCatalogos2(iTipoCatalogo As Integer, sNombre As String) As ADODB.Recordset
'Dim oConexion As New ADODB.Connection
'Dim oTablaServicios As New CatalogoServicios
'Dim oTablaBienes As New CatalogoBienesInsumos
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    If iTipoCatalogo = sighentidades.sghDetalleComprobanteServicios Then
'        Set oTablaServicios.Conexion = oConexion
'        Set BuscarCoincidenciasEnCatalogos2 = oTablaServicios.BuscarCoincidenciaNombre(iTipoCatalogo, sNombre)
'        ms_MensajeError = oTablaServicios.MensajeError
'    Else
'        Set oTablaBienes.Conexion = oConexion
'        Set BuscarCoincidenciasEnCatalogos2 = oTablaBienes.BuscarCoincidenciaNombre2(iTipoCatalogo, sNombre)
'        ms_MensajeError = oTablaBienes.MensajeError
'    End If
'    Set oTablaServicios = Nothing
'    Set oTablaBienes = Nothing
'
'End Function


'Function CajaTiposMoneda() As ADODB.Recordset
'Dim oConexion As New ADODB.Connection
'Dim oTabla As New CajaTiposMoneda
'
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    Set oTabla.Conexion = oConexion
'    Set CajaTiposMoneda = oTabla.SeleccionarTodos()
'    oConexion.Close
'
'    ms_MensajeError = oTabla.MensajeError
'
'    Set oConexion = Nothing
'    Set oTabla = Nothing
'End Function

Function ImpuestoIGV() As Double
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposImpuesto
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    ImpuestoIGV = oTabla.ObtenerIGV
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ObtenerSiguienteNumeroDocumento(oCajaNroDoc As DOCajaNroDocumento) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCaja

    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    ObtenerSiguienteNumeroDocumento = oTabla.ObtenerSiguienteNumeroDocumento(oCajaNroDoc)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function ObtenerSiguienteNumeroDocumentoYgrabarlo(ByVal oCajaNroDoc As DOCajaNroDocumento) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New CajaCaja
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    ObtenerSiguienteNumeroDocumentoYgrabarlo = oTabla.ObtenerSiguienteNumeroDocumentoYgrabarlo(oCajaNroDoc)
    oConexion.Close
    Set oTabla.Conexion = Nothing
    ms_MensajeError = oTabla.MensajeError
End Function


Function ComprobantePagoAgregar(oDOCajaComprobantesPago As DOCajaComprobantesPago, cItemsAFacturar As Collection, cItemsDinero As Collection, oDOCaja As DOCajaCaja) As Boolean
End Function
Function ComprobantePagoModificar(oDOCajaComprobantesPago As DOCajaComprobantesPago, cItemsAFacturar As Collection, cItemsDinero As Collection) As Boolean
End Function

Function ComprobantePagoEliminar(oDOCajaComprobantesPago As DOCajaComprobantesPago) As Boolean
End Function

Function ComprobantePagoSeleccionarPorId(IdComprobantePago As Long, oConexion As Connection) As DOCajaComprobantesPago
Dim oCajaComprobantesPago As New CajaComprobantesPago
Dim oDoComprobantesPago As New DOCajaComprobantesPago
    Set ComprobantePagoSeleccionarPorId = Nothing
    Set oCajaComprobantesPago.Conexion = oConexion
    oDoComprobantesPago.IdComprobantePago = IdComprobantePago
    If oCajaComprobantesPago.SeleccionarPorId(oDoComprobantesPago) Then
        Set ComprobantePagoSeleccionarPorId = oDoComprobantesPago
    End If
    ms_MensajeError = oCajaComprobantesPago.MensajeError
    Set oCajaComprobantesPago = Nothing
End Function

Function ComprobantePagoSeleccionarPorCuentaAtencion(idCuentaAtencion As Long) As DOCajaComprobantesPago
Dim oConexion As New ADODB.Connection
Dim oCajaComprobantesPago As New CajaComprobantesPago
Dim oDoComprobantesPago As New DOCajaComprobantesPago
    
    Set ComprobantePagoSeleccionarPorCuentaAtencion = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajaComprobantesPago.Conexion = oConexion
    oDoComprobantesPago.idCuentaAtencion = idCuentaAtencion
    If oCajaComprobantesPago.SeleccionarPorCuentaAtencion(oDoComprobantesPago) Then
        Set ComprobantePagoSeleccionarPorCuentaAtencion = oDoComprobantesPago
    Else
        Set ComprobantePagoSeleccionarPorCuentaAtencion = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oCajaComprobantesPago.MensajeError
    Set oConexion = Nothing
    Set oCajaComprobantesPago = Nothing
End Function


Function ObtenerCuentasAtencionPorHistoriaClinica1(NroHistoriaClinica As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion


    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oCuentasAtencion.Conexion = oConexion

    Set ObtenerCuentasAtencionPorHistoriaClinica1 = oCuentasAtencion.ObtenerCuentasAtencionPorHistoriaClinica(NroHistoriaClinica)
    oConexion.Close
    ms_MensajeError = oCuentasAtencion.MensajeError
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing

End Function
Function ObtenerUltimaCuentaAtencionPorIdPaciente(lidPaciente As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    Set ObtenerUltimaCuentaAtencionPorIdPaciente = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oCuentasAtencion.Conexion = oConexion
    
    Set ObtenerUltimaCuentaAtencionPorIdPaciente = oCuentasAtencion.ObtenerUltimaCuentaAtencionPorIdPaciente(lidPaciente)
    oConexion.Close
    ms_MensajeError = oCuentasAtencion.MensajeError
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing

End Function
Function ObtenerCuentasAtencionPorHistoriaClinicaV2(NroHistoriaClinica As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    Set oCuentasAtencion.Conexion = oConexion
    
    Set ObtenerCuentasAtencionPorHistoriaClinicaV2 = oCuentasAtencion.ObtenerCuentasAtencionPorHistoriaClinicaV2(NroHistoriaClinica)
    oConexion.Close
    ms_MensajeError = oCuentasAtencion.MensajeError
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing

End Function
'Function ObtenerDetallesFacturablesPorPaciente(idPaciente As Long, TipoFiltro As String) As Recordset
'Dim oConexion As New ADODB.Connection
'Dim oCuentasAtencion As New CuentasAtencion
'
'
'    Set ObtenerDetallesFacturablesPorPaciente = Nothing
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    Set oCuentasAtencion.Conexion = oConexion
'
'    Set ObtenerDetallesFacturablesPorPaciente = oCuentasAtencion.ObtenerDetallesFacturablesPorPaciente(idPaciente, ID_PENDIENTE_PAGO, TipoFiltro)
'    oConexion.Close
'    ms_MensajeError = oCuentasAtencion.MensajeError
'    Set oConexion = Nothing
'    Set oCuentasAtencion = Nothing
'
'End Function
'Function ObtenerDetallesFacturablesPorCuenta(IdCuenta As Long, TipoFiltro As String) As Recordset
'Dim oConexion As New ADODB.Connection
'Dim oCuentasAtencion As New CuentasAtencion
'
'
'    Set ObtenerDetallesFacturablesPorCuenta = Nothing
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    Set oCuentasAtencion.Conexion = oConexion
'
'    Set ObtenerDetallesFacturablesPorCuenta = oCuentasAtencion.ObtenerDetallesFacturablesPorCuenta(IdCuenta, ID_PENDIENTE_PAGO, TipoFiltro)
'    oConexion.Close
'    ms_MensajeError = oCuentasAtencion.MensajeError
'    Set oConexion = Nothing
'    Set oCuentasAtencion = Nothing
'
'End Function
Function ObtenerPacientePorId(oPaciente As DOPaciente) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    ObtenerPacientePorId = False
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    ObtenerPacientePorId = oTabla.SeleccionarPorId(oPaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
'-------------------------------------------------------------------------------------
'                                       Cajas
'-------------------------------------------------------------------------------------

Function RealizarFiltroCajas(oCaja As DOCajaCaja) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCaja
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RealizarFiltroCajas = oTabla.RealizarFiltro(oCaja)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function CajaAgregar(oDOCaja As DOCajaCaja, cNroDocumentos As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNcaja As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCaja As New CajaCaja
Dim oNroDocumento As New CajaNroDocumento
Dim ItemDocumento As DOCajaNroDocumento
    
    CajaAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oCaja.Conexion = oConexion
    Set oNroDocumento.Conexion = oConexion
    If oCaja.Insertar(oDOCaja) Then
        For Each ItemDocumento In cNroDocumentos
            ItemDocumento.idCaja = oDOCaja.idCaja
            If Not oNroDocumento.Insertar(ItemDocumento) Then
                ms_MensajeError = oNroDocumento.MensajeError
                Exit Function
            End If
        Next
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCaja.IdUsuarioAuditoria, "A", oDOCaja.idCaja, "CajaCaja", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNcaja)        'ListBarItems.idListItem
        CajaAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCaja.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCaja = Nothing
    Set oNroDocumento = Nothing
End Function

Function CajaModificar(oDOCaja As DOCajaCaja, cNroDocumentos As Collection, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNcaja As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCaja As New CajaCaja
Dim oNroDocumento As New CajaNroDocumento
Dim ItemDocumento As DOCajaNroDocumento
    
    CajaModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oCaja.Conexion = oConexion
    Set oNroDocumento.Conexion = oConexion
    If oCaja.modificar(oDOCaja) Then
        If Not oNroDocumento.EliminarPorCaja(oDOCaja.idCaja) Then
            ms_MensajeError = oNroDocumento.MensajeError
            Exit Function
        End If
        For Each ItemDocumento In cNroDocumentos
            ItemDocumento.idCaja = oDOCaja.idCaja
            If Not oNroDocumento.Insertar(ItemDocumento) Then
                ms_MensajeError = oNroDocumento.MensajeError
                Exit Function
            End If
        Next
    
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCaja.IdUsuarioAuditoria, "M", oDOCaja.idCaja, "CajaCaja", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNcaja)        'ListBarItems.idListItem
        CajaModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCaja.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCaja = Nothing
    Set oNroDocumento = Nothing
    
End Function

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Function CajaEliminar(oDOCaja As DOCajaCaja, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                     lcNcaja As String) As Boolean
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim oCaja As New CajaCaja
'Dim oRsTmp As New Recordset
    CajaEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oCaja.Conexion = oConexion
  
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "CajaNroDocumentoEliminarXiIdCaja"
      Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, oDOCaja.idCaja): .Parameters.Append oParameter
      .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
'    Set oRsTmp = Nothing
    If oCaja.Eliminar(oDOCaja) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCaja.IdUsuarioAuditoria, "E", oDOCaja.idCaja, "CajaCaja", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNcaja)        'ListBarItems.idListItem
        CajaEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oCaja.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
    Set oCaja = Nothing
End Function

Function CajaSeleccionarPorId(idCaja As Long) As DOCajaCaja
Dim oConexion As New ADODB.Connection
Dim oCaja As New CajaCaja
Dim oDOCaja As New DOCajaCaja
    
    Set CajaSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCaja.Conexion = oConexion


    oDOCaja.idCaja = idCaja
    If oCaja.SeleccionarPorId(oDOCaja) Then
        Set CajaSeleccionarPorId = oDOCaja
    End If
    oConexion.Close
    ms_MensajeError = oCaja.MensajeError
    Set oConexion = Nothing
    Set oCaja = Nothing
End Function
Function CajaSeleccionarTodosParaLista() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCaja
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CajaSeleccionarTodosParaLista = oTabla.SeleccionarTodosParaLista()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function CajaDevuelveNombreImpresoraPorId(lnIdCaja As Long, lbEsLaImpresoraDefault As Boolean, Optional oConexion1 As Connection) As String
    Dim oRsTmp1 As New ADODB.Recordset
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
      .CommandText = "CajaDevuelveNombreImpresoraPorId"
      Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
      Set oRsTmp1 = .Execute
      If oConexion1 Is Nothing Then
         Set oRsTmp1.ActiveConnection = Nothing
      End If
    End With
    Set oCommand = Nothing
      
    CajaDevuelveNombreImpresoraPorId = ""
    If oRsTmp1.RecordCount > 0 Then
        If lbEsLaImpresoraDefault = True Then
           CajaDevuelveNombreImpresoraPorId = IIf(IsNull(oRsTmp1.Fields!ImpresoraDefault), "", oRsTmp1.Fields!ImpresoraDefault)
        Else
           CajaDevuelveNombreImpresoraPorId = IIf(IsNull(oRsTmp1.Fields!Impresora2), "", oRsTmp1.Fields!Impresora2)
        End If
    End If
    If oConexion1 Is Nothing Then
        oRsTmp1.Close
        oConexion.Close
        Set oRsTmp1 = Nothing
    End If
    Set oConexion = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Cajeros
'-------------------------------------------------------------------------------------

Function RealizarFiltroCajeros(oEmpleado As DOEmpleado) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCajero
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set RealizarFiltroCajeros = oTabla.RealizarFiltro(oEmpleado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function CajeroAgregar(oDOCajero As DOCajaCajero) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajero As New CajaCajero

    CajeroAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajero.Conexion = oConexion
       
    If oCajero.Insertar(oDOCajero) Then
        CajeroAgregar = True
    Else
        ms_MensajeError = oCajero.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajero = Nothing
End Function

Function CajeroModificar(oDOCajero As DOCajaCajero) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajero As New CajaCajero
    
    CajeroModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajero.Conexion = oConexion
       
    If oCajero.modificar(oDOCajero) Then
        CajeroModificar = True
    Else
        ms_MensajeError = oCajero.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajero = Nothing
    
End Function

Function CajeroEliminar(oDOCajero As DOCajaCajero) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajero As New CajaCajero
    
    CajeroEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajero.Conexion = oConexion
        
    If oCajero.Eliminar(oDOCajero) Then
        CajeroEliminar = True
    Else
        ms_MensajeError = oCajero.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajero = Nothing
    
End Function

Function CajeroSeleccionarPorId(IdCajero As Long) As DOCajaCajero
Dim oConexion As New ADODB.Connection
Dim oCajero As New CajaCajero
Dim oDOCajero As New DOCajaCajero
    
    Set CajeroSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajero.Conexion = oConexion
    oDOCajero.IdCajero = IdCajero
    If oCajero.SeleccionarPorId(oDOCajero) Then
        Set CajeroSeleccionarPorId = oDOCajero
    End If
    oConexion.Close
    ms_MensajeError = oCajero.MensajeError
    Set oConexion = Nothing
    Set oCajero = Nothing
End Function
Function CajerosSeleccionarTodosParaLista() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCajero
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CajerosSeleccionarTodosParaLista = oTabla.SeleccionarTodosParaLista()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function CajerosSeleccionarSegunUsuario(idEmpleado As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaCajero
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CajerosSeleccionarSegunUsuario = oTabla.SeleccionarSegunUsuario(idEmpleado)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
'-------------------------------------------------------------------------------------
'                                       Turnos
'-------------------------------------------------------------------------------------

Function TurnosSeleccionarTodosParaLista() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaTurno
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TurnosSeleccionarTodosParaLista = oTabla.SeleccionarTodosParaLista()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'-------------------------------------------------------------------------------------
'                                       Nro Documentos x Caja
'-------------------------------------------------------------------------------------

Function NroDocumentoSeleccionarPorIdCaja(ByVal lIdCaja As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaNroDocumento
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set NroDocumentoSeleccionarPorIdCaja = oTabla.SeleccionarPorIdCaja(lIdCaja)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function NroDocumentoSeleccionarPorIdCajaYTipoComprobante(ByVal lIdCaja As Long, idTipoComprobante As Long) As DOCajaNroDocumento
Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaNroDocumento
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set NroDocumentoSeleccionarPorIdCajaYTipoComprobante = oTabla.SeleccionarPorIdCajaYTipoComprobante(lIdCaja, idTipoComprobante)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'--------------------------------------------------------------------------------------
'                                Detalles de facturacion
'--------------------------------------------------------------------------------------

Function CatalogoServicioPorCuentaAtencionSQl2000(idCuentaAtencion As Long, IdTipoFinanciamiento As Long, IdPuntosDeCarga As Long, lEstadoFacturacion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoServicioPorCuentaAtencionSQl2000 = oTabla.SeleccionarPorCuentaAtencionNoPagados(idCuentaAtencion, IdTipoFinanciamiento, IdPuntosDeCarga, lEstadoFacturacion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function CatalogoBienesInsumosPorCuentaAtencion(idCuentaAtencion As Long, IdTipoFinanciamiento As Long, IdPuntosDeCarga As Long, lEstadoFacturacion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CatalogoBienesInsumosPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencionNoPagados(idCuentaAtencion, IdTipoFinanciamiento, IdPuntosDeCarga, lEstadoFacturacion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


'-----------------------------gestion de caja ----------------------------------------
Function CajaGestionFiltrarSql2000(doCajaGestion As doCajaGestion, oDOCajaComprobantesPago As DOCajaComprobantesPago, oDoPaciente As DOPaciente, lcFechaComprobante As String) As ADODB.Recordset

Dim oConexion As New ADODB.Connection
Dim oTabla As New CajaGestion
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    
    
    Set CajaGestionFiltrarSql2000 = oTabla.FiltrarCajaGestion(doCajaGestion, oDOCajaComprobantesPago, oDoPaciente, lcFechaComprobante)
    
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError


    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function CajaGestionAgregar(oDoCajaGestion As doCajaGestion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaGestion As New CajaGestion

    CajaGestionAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaGestion.Conexion = oConexion
       
    If oCajaGestion.Insertar(oDoCajaGestion) Then
        CajaGestionAgregar = True
    Else
        ms_MensajeError = oCajaGestion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaGestion = Nothing
End Function

Function CajaGestionModificar(oDoCajaGestion As doCajaGestion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaGestion As New CajaGestion
Dim total As Double
    
    CajaGestionModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaGestion.Conexion = oConexion
    
    total = oCajaGestion.TotalPorGestion(oDoCajaGestion.IdGestionCaja)
    oDoCajaGestion.TotalCobrado = total
    If oCajaGestion.modificar(oDoCajaGestion) Then
        CajaGestionModificar = True
    Else
        ms_MensajeError = oCajaGestion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaGestion = Nothing
    
End Function

Function CajaGestionEliminar(oDoCajaGestion As doCajaGestion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaGestion As New CajaGestion
    
    CajaGestionEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaGestion.Conexion = oConexion
        
    If oCajaGestion.Eliminar(oDoCajaGestion) Then
        CajaGestionEliminar = True
    Else
        ms_MensajeError = oCajaGestion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaGestion = Nothing
    
End Function

Function CajaGestionSeleccionarPorId(IdCajaGestion As Long) As doCajaGestion
Dim oConexion As New ADODB.Connection
Dim oCajaGestion As New CajaGestion
Dim oDoCajaGestion As New doCajaGestion
    
    Set CajaGestionSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajaGestion.Conexion = oConexion
    oDoCajaGestion.IdCajero = IdCajaGestion
    If oCajaGestion.SeleccionarPorId(oDoCajaGestion) Then
        Set CajaGestionSeleccionarPorId = oDoCajaGestion
    End If
    oConexion.Close
    ms_MensajeError = oCajaGestion.MensajeError
    Set oConexion = Nothing
    Set oCajaGestion = Nothing
End Function

'Function CajaComprobantePagoCuentaAtencionAgregarS(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, rsFacturacionProductosServ As Recordset, lnIdCuentaAtencionAcerrar As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oDOFacturacionServicios As New DOFacturacionServicios
'Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
'Dim oDOSeguros As New DOFacturacionSeguros
'Dim oFacturacionServicios As New FacturacionServicios
'Dim oFacturacionSeguros As New FacturacionSeguros
'Dim oFactOrdenesServicio As New FactOrdenesServicio
'Dim oCajaComprobantePago As New CajaComprobantesPago
'Dim oCajaNroDocumento As New CajaNroDocumento
'Dim oDOCuentaAtencion As New DOCuentaAtencion
'Dim oCuentasAtencion As New CuentasAtencion
'Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
'Dim bProcesoOK As Boolean
'Dim sResultado As String
'Dim sIdOrdenesServProcesadas  As String
'Dim sIdOrdenesServProcesadasMemo As String
'Dim sIdOrdenesBienProcesadas  As String
'Dim sIdOrdenesBienProcesadasMemo As String
'Dim sSql As String: Dim lnDctosEnCajaComprobantesPago  As Double
'Dim rsActualizar As New ADODB.Recordset
'Dim oFactOrdenServicio As New FactOrdenServicio
'Dim oDoFactOrdenServ As New DoFactOrdenServ
'Dim oFactOrdenServicioPagos As New FactOrdenServicioPagos
'Dim oDoFactOrdenServPagos As New DoFactOrdenServPagos
'Dim lnIdOrdenPago As Long
'
'    CajaComprobantePagoCuentaAtencionAgregarS = False
'
'    oConexion.Open SIGHEntidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    oConexion.BeginTrans
'    Set oFacturacionSeguros.Conexion = oConexion
'    Set oCajaComprobantePago.Conexion = oConexion
'    Set oCajaNroDocumento.Conexion = oConexion
'    Set oCuentasAtencion.Conexion = oConexion
'    Set oFactOrdenServicioPagos.Conexion = oConexion
'    Set oFactOrdenServicio.Conexion = oConexion
'
'    ms_MensajeError = ""
'
'    'AGREGAR EL COMPROBANTE DE PAGO
'    oDOCajaComprobantePago.idEstadoComprobante = 4  'PAGADO
'    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
'        bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
'    End If
'    '
'    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
'
'    '*********************************************************************************
'    '                                  PROCESAR SERVICIOS
'    '*********************************************************************************
'    If rsFacturacionProductosServ.RecordCount > 0 Then
'        rsFacturacionProductosServ.MoveFirst
'        Do While Not rsFacturacionProductosServ.EOF
'           lnIdOrdenPago = rsFacturacionProductosServ.Fields!IdOrden
'           oDoFactOrdenServPagos.idOrdenPago = lnIdOrdenPago
'           If Not oFactOrdenServicioPagos.SeleccionarPorId(oDoFactOrdenServPagos) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenServicioPagos.MensajeError: GoTo ErrorManager
'           End If
'           oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'pagado
'           oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
'           oDoFactOrdenServPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'           If Not oFactOrdenServicioPagos.Modificar(oDoFactOrdenServPagos) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenServicioPagos.MensajeError: GoTo ErrorManager
'           End If
'           oDoFactOrdenServ.IdOrden = oDoFactOrdenServPagos.IdOrden
'           If Not oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenServicio.MensajeError: GoTo ErrorManager
'           End If
'           oDoFactOrdenServ.IdEstadoFacturacion = 4  'pagado
'           oDoFactOrdenServ.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'           If Not oFactOrdenServicio.Modificar(oDoFactOrdenServ) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenServicio.MensajeError: GoTo ErrorManager
'           End If
'           Do While Not rsFacturacionProductosServ.EOF And lnIdOrdenPago = rsFacturacionProductosServ.Fields!IdOrden
'              rsFacturacionProductosServ.MoveNext
'              If rsFacturacionProductosServ.EOF Then
'                 Exit Do
'              End If
'           Loop
'        Loop
'    End If
'
'
'    '********************************************************************************
'    '                               COMPROBANTE DE PAGO
'    '********************************************************************************
'    'ACTUALIZAR NRO DE DOCUMENTO
'    Dim oDoCajaNroDocumento As New DOCajaNroDocumento
'    oDoCajaNroDocumento.IdCaja = oDOGestionCaja.IdCaja
'    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
'    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
'    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
'    If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
'        bProcesoOK = False: ms_MensajeError = oCajaNroDocumento.MensajeError: GoTo ErrorManager
'    End If
'
'    'Cierra Cuenta si se CANCELA BOLETAS DE FARMACIA Y SERVICIOS
'    If lnIdCuentaAtencionAcerrar > 0 Then
'        oDOCuentaAtencion.IdCuentaAtencion = lnIdCuentaAtencionAcerrar
'        If Not oCuentasAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
'           bProcesoOK = False: ms_MensajeError = oCuentasAtencion.MensajeError: GoTo ErrorManager
'        End If
'        oDOCuentaAtencion.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'        If Not mo_ReglasFacturacion.CuentasAtencionPagada(oDOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") = True Then
'           bProcesoOK = False: ms_MensajeError = mo_ReglasFacturacion.MensajeError: GoTo ErrorManager
'        End If
'    End If
'
'
'    oConexion.CommitTrans
'    oConexion.Close
'    ms_MensajeError = ""
'    CajaComprobantePagoCuentaAtencionAgregarS = True
'ErrorManager:
'    If CajaComprobantePagoCuentaAtencionAgregarS = False Then
'        oConexion.RollbackTrans
'        oConexion.Close
'    End If
'    Set oConexion = Nothing
'    Set oDOFacturacionServicios = Nothing
'    Set oDoFacturacionBienes = Nothing
'    Set oDOSeguros = Nothing
'    Set oFacturacionServicios = Nothing
'    Set oFacturacionSeguros = Nothing
'    Set oFactOrdenesServicio = Nothing
'    Set oCajaComprobantePago = Nothing
'    Set oCajaNroDocumento = Nothing
'    Set oDOCuentaAtencion = Nothing
'    Set oCuentasAtencion = Nothing
'    Set mo_ReglasFacturacion = Nothing
'End Function

'Function CajaComprobantePagoCuentaAtencionAgregarB(oDOCajaComprobantePago As DOCajaComprobantesPago, rsFacturacionProductosBien As Recordset, lnIdCuentaAtencionAcerrar As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oFactOrdenesBienes As New FactOrdenesBienes
'Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
'Dim oCajaComprobantePago As New CajaComprobantesPago
'Dim oCajaNroDocumento As New CajaNroDocumento
'Dim oDoCajaNroDocumento As New DOCajaNroDocumento
'Dim oDOCuentaAtencion As New DOCuentaAtencion
'Dim oCuentasAtencion As New CuentasAtencion
'Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
'Dim bProcesoOK As Boolean
'Dim sResultado As String
'Dim lnIdOrden As Long
'
'    CajaComprobantePagoCuentaAtencionAgregarB = False
'
'  oConexion.CursorLocation = adUseClient
'  oConexion.CommandTimeout = 300
'  oConexion.Open SIGHEntidades.CadenaConexion
'
'    ms_MensajeError = ""
'    oConexion.BeginTrans
'    Set oFactOrdenesBienes.Conexion = oConexion
'    Set oCajaComprobantePago.Conexion = oConexion
'    Set oCajaNroDocumento.Conexion = oConexion
'    Set oCuentasAtencion.Conexion = oConexion
'
'    oDOCajaComprobantePago.idEstadoComprobante = 4  'PAGADO
'    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
'        bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
'    End If
'    '
'    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
'
'    '*********************************************************************************
'    '                             PROCESAR BIENES E INSUMOS
'    '*********************************************************************************
'    If rsFacturacionProductosBien.RecordCount > 0 Then
'        rsFacturacionProductosBien.MoveFirst
'        Do While Not rsFacturacionProductosBien.EOF
'           lnIdOrden = rsFacturacionProductosBien.Fields!IdOrden
'           Do While Not rsFacturacionProductosBien.EOF And lnIdOrden = rsFacturacionProductosBien.Fields!IdOrden
'              rsFacturacionProductosBien.MoveNext
'              If rsFacturacionProductosBien.EOF Then
'                 Exit Do
'              End If
'           Loop
'           oDoFactOrdenesBienes.IdOrden = lnIdOrden
'           If Not oFactOrdenesBienes.SeleccionarPorId(oDoFactOrdenesBienes) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrorManager
'           End If
'           oDoFactOrdenesBienes.IdEstadoFacturacion = 4   'pagado
'           oDoFactOrdenesBienes.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
'           oDoFactOrdenesBienes.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'           If Not oFactOrdenesBienes.Modificar(oDoFactOrdenesBienes) Then
'              bProcesoOK = False: ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrorManager
'           End If
'        Loop
'    End If
'    '********************************************************************************
'    '                               COMPROBANTE DE PAGO
'    '********************************************************************************
'    'ACTUALIZAR NRO DE DOCUMENTO
'    oDoCajaNroDocumento.IdCaja = oDOCajaComprobantePago.IdCaja
'    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
'    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
'    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
'    oDoCajaNroDocumento.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'    If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
'        bProcesoOK = False: ms_MensajeError = oCajaNroDocumento.MensajeError: GoTo ErrorManager
'    End If
'
'    'Cierra Cuenta si se CANCELA BOLETAS DE FARMACIA Y SERVICIOS
'    If lnIdCuentaAtencionAcerrar > 0 Then
'        oDOCuentaAtencion.IdCuentaAtencion = lnIdCuentaAtencionAcerrar
'        If Not oCuentasAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
'           bProcesoOK = False: ms_MensajeError = oCuentasAtencion.MensajeError: GoTo ErrorManager
'        End If
'        oDOCuentaAtencion.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
'        If Not mo_ReglasFacturacion.CuentasAtencionPagada(oDOCuentaAtencion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") = True Then
'           bProcesoOK = False: ms_MensajeError = mo_ReglasFacturacion.MensajeError: GoTo ErrorManager
'        End If
'    End If
'    '
'    oConexion.CommitTrans
'    oConexion.Close
'    ms_MensajeError = ""
'    CajaComprobantePagoCuentaAtencionAgregarB = True
'ErrorManager:
'    If CajaComprobantePagoCuentaAtencionAgregarB = False Then
'        oConexion.RollbackTrans
'        oConexion.Close
'    End If
'    Set oConexion = Nothing
'    Set oFactOrdenesBienes = Nothing
'    Set oDoFactOrdenesBienes = Nothing
'    Set oCajaComprobantePago = Nothing
'    Set oCajaNroDocumento = Nothing
'    Set oDoCajaNroDocumento = Nothing
'    Set oDOCuentaAtencion = Nothing
'    Set oCuentasAtencion = Nothing
'    Set mo_ReglasFacturacion = Nothing
'
'End Function

'Function CajaComprobantePagoServicioAgregar(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, oDoFactOrdenServPagos As DoFactOrdenServPagos, rsFacturacionProductos As Recordset, ml_IdUsuarioAuditoria As Long, oDoAtencion As DOAtencion, lnIdPuntoCarga As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdReceta As Long) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oDOFacturacionServicios As New DOFacturacionServicios
'Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
'Dim oDOSeguros As New DOFacturacionSeguros
'Dim oFacturacionServicios As New FacturacionServicios
'Dim oFacturacionBienes As New FacturacionBienesInsumos
'Dim oFacturacionSeguros As New FacturacionSeguros
'Dim oFactOrdenesServicio As New FactOrdenesServicio
'Dim oCajaComprobantePago As New CajaComprobantesPago
'Dim oCajaNroDocumento As New CajaNroDocumento
'Dim oFactOrdenServPagos As New FactOrdenServicioPagos
'Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
'Dim oFacturacionServicioPagos As New FacturacionServicioPagos
'Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
'Dim oDoFacturacionServicioDespacho As New DoFacturacionServicioDespacho
'Dim oDoFactOrdenServ As New DoFactOrdenServ
'Dim oFactOrdenServ As New FactOrdenServicio
'Dim oCitas As New Citas
'Dim bProcesoOK As Boolean
'Dim sResultado As String
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'
'    CajaComprobantePagoServicioAgregar = False
'
'    oConexion.CursorLocation = adUseClient
'    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
'
'    oConexion.BeginTrans
'    Set oFacturacionServicios.Conexion = oConexion
'    Set oFacturacionBienes.Conexion = oConexion
'    Set oFacturacionSeguros.Conexion = oConexion
'    Set oFactOrdenesServicio.Conexion = oConexion
'    Set oCajaComprobantePago.Conexion = oConexion
'    Set oCajaNroDocumento.Conexion = oConexion
'    Set oCitas.Conexion = oConexion
'    Set oFactOrdenServPagos.Conexion = oConexion
'    Set oFacturacionServicioPagos.Conexion = oConexion
'    Set oFacturacionServicioDespacho.Conexion = oConexion
'    Set oFactOrdenServ.Conexion = oConexion
'
'    oDOCajaComprobantePago.idEstadoComprobante = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
'    oDOCajaComprobantePago.IdTipoOrden = 1  'Servicio
'    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
'        bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
'    End If
'    '
'    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
'
'    'Crear orden de servicio
'    oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'Pagado
'    oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
'    If oDoFactOrdenServPagos.idOrdenPago = 0 Then
'        'Se ha creado desde CAJA
'        oDoFactOrdenServ.FechaCreacion = oDoFactOrdenServPagos.FechaCreacion
'        oDoFactOrdenServ.fechaDespacho = oDoFactOrdenServPagos.FechaCreacion
'        oDoFactOrdenServ.IdCuentaAtencion = oDOCajaComprobantePago.IdCuentaAtencion
'        oDoFactOrdenServ.IdEstadoFacturacion = 4 'PAGADO
'        oDoFactOrdenServ.idFuenteFinanciamiento = 1 'Contado
'        oDoFactOrdenServ.idPaciente = oDOCajaComprobantePago.idPaciente
'        oDoFactOrdenServ.idPuntoCarga = lnIdPuntoCarga
'        oDoFactOrdenServ.IdTipoFinanciamiento = oDOCajaComprobantePago.IdTipoFinanciamiento
'        oDoFactOrdenServ.idUsuario = oDoFactOrdenServPagos.IdUsuarioAuditoria
'        oDoFactOrdenServ.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
'        oDoFactOrdenServ.IdUsuarioDespacho = oDoFactOrdenServPagos.IdUsuarioAuditoria
'        If oDoFactOrdenServ.IdCuentaAtencion > 0 Then
'           oDoFactOrdenServ.IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFactOrdenServ.IdCuentaAtencion, Format(oDoFactOrdenServ.FechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY), Format(Now, "hh:mm"))
'        Else
'           oDoFactOrdenServ.IdServicioPaciente = 0
'        End If
'        If Not oFactOrdenServ.Insertar(oDoFactOrdenServ) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'        oDoFactOrdenServPagos.IdOrden = oDoFactOrdenServ.IdOrden
'        If Not oFactOrdenServPagos.Insertar(oDoFactOrdenServPagos) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'        rsFacturacionProductos.MoveFirst
'        Do While Not rsFacturacionProductos.EOF
'           oDoFacturacionServicioDespacho.cantidad = rsFacturacionProductos!cantidad
'           oDoFacturacionServicioDespacho.IdOrden = oDoFactOrdenServ.IdOrden
'           oDoFacturacionServicioDespacho.idProducto = rsFacturacionProductos!idProducto
'           oDoFacturacionServicioDespacho.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
'           oDoFacturacionServicioDespacho.Precio = rsFacturacionProductos!PrecioUnitario
'           oDoFacturacionServicioDespacho.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
'           If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDespacho) Then
'                bProcesoOK = False: GoTo ErrorManager
'           End If
'           oDoFacturacionServicioPagos.cantidad = rsFacturacionProductos!cantidad
'           oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
'           oDoFacturacionServicioPagos.idProducto = rsFacturacionProductos!idProducto
'           oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
'           oDoFacturacionServicioPagos.Precio = rsFacturacionProductos!PrecioUnitario
'           oDoFacturacionServicioPagos.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
'           If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
'                bProcesoOK = False: GoTo ErrorManager
'           End If
'           rsFacturacionProductos.MoveNext
'        Loop
'    Else
'        'Se ha creado en FACTURACION, Es una PREVENTA (SERVICIOS)
'        If Not oFactOrdenServPagos.Modificar(oDoFactOrdenServPagos) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'        oDoFactOrdenServ.IdOrden = oDoFactOrdenServPagos.IdOrden
'        If Not oFactOrdenServ.SeleccionarPorId(oDoFactOrdenServ) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'        oDoFactOrdenServ.IdEstadoFacturacion = 4 'Pagado
'        If Not oFactOrdenServ.Modificar(oDoFactOrdenServ) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'    End If
'
'    'Si la orden es de Admsión de Consulta Externa
'    If lnIdPuntoCarga = 6 Then
'        'Generar una solicitud automatica, actualizar el estado de la cita a pagado
'        If Not oCitas.ActualizarCitaPagadaDEBB(oDoAtencion.idAtencion, ml_IdUsuarioAuditoria) Then
'            bProcesoOK = False: GoTo ErrorManager
'        End If
'    End If
'
'    Dim oDoCajaNroDocumento As New DOCajaNroDocumento
'    oDoCajaNroDocumento.IdCaja = oDOCajaComprobantePago.IdCaja
'    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
'    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
'    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
'    If Not oCajaNroDocumento.Modificar(oDoCajaNroDocumento) Then
'        bProcesoOK = False: GoTo ErrorManager
'    End If
'    'Receta automatica debb-03/04/2012
'    If lnIdReceta > 0 Then
'        With oCommand
'          .CommandType = adCmdStoredProc
'          Set .ActiveConnection = oConexion
'          .CommandTimeout = 150
'          .CommandText = "ActualizaRecetaCabecerarXEstadoYComprobantePago"
'          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oDOCajaComprobantePago.IdComprobantePago): .Parameters.Append oParameter
'          Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, lnIdReceta): .Parameters.Append oParameter
'          .Execute
'        End With
'        Set oCommand = Nothing
'       'Set oRsTmp = Nothing
'    End If    '
'    oConexion.CommitTrans
'    oConexion.Close
'    Set oConexion = Nothing
'    ms_MensajeError = ""
'    CajaComprobantePagoServicioAgregar = True
'ErrorManager:
'
'    If CajaComprobantePagoServicioAgregar = False Then
'
'        ms_MensajeError = Err.Description
'        oConexion.RollbackTrans
'        oConexion.Close
'        Set oConexion = Nothing
'    End If
'    Set oDOFacturacionServicios = Nothing
'    Set oDoFacturacionBienes = Nothing
'    Set oDOSeguros = Nothing
'    Set oFacturacionServicios = Nothing
'    Set oFacturacionBienes = Nothing
'    Set oFacturacionSeguros = Nothing
'    Set oFactOrdenesServicio = Nothing
'    Set oCajaComprobantePago = Nothing
'    Set oCajaNroDocumento = Nothing
'    Set oFactOrdenServPagos = Nothing
'    Set oDoFacturacionServicioPagos = Nothing
'    Set oFacturacionServicioPagos = Nothing
'    Set oFacturacionServicioDespacho = Nothing
'    Set oDoFacturacionServicioDespacho = Nothing
'    Set oDoFactOrdenServ = Nothing
'    Set oFactOrdenServ = Nothing
'    Set oCitas = Nothing
'End Function

Function CajaComprobantePagoBienInsumoAgregar(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, oDOFactOrdenBien As DOFactOrdenBienInsumo, rsFacturacionProductosBien As Recordset, mo_ProductosBienEliminados As Collection, ml_IdUsuarioAuditoria As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
Dim oFacturacionBienes As New FacturacionBienesInsumos
Dim oFacturacionSeguros As New FacturacionSeguros
Dim oFactOrdenesBienes As New FactOrdenesBienesInsumo
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim oCajaNroDocumento As New CajaNroDocumento
Dim bProcesoOK As Boolean
Dim sResultado As String

    CajaComprobantePagoBienInsumoAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oFacturacionBienes.Conexion = oConexion
    Set oFacturacionSeguros.Conexion = oConexion
    Set oFactOrdenesBienes.Conexion = oConexion
    Set oCajaComprobantePago.Conexion = oConexion
    Set oCajaNroDocumento.Conexion = oConexion
    
    oDOCajaComprobantePago.IdTipoOrden = 2  'Bienes e Insumos, CAJA DE FARMACIA
    oDOCajaComprobantePago.idEstadoComprobante = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
    
    'Crear orden de servicio
    oDOFactOrdenBien.IdEstadoOrden = 4  'Pagado
    oDOFactOrdenBien.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
    
    'Se supone que el comprobante ya se agrego al momento de agregar los servicios
    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
    
    oDOFactOrdenBien.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
    
    'Si es una orden nueva
    If oDOFactOrdenBien.IdOrden = 0 Then
        If Not oFactOrdenesBienes.Insertar(oDOFactOrdenBien) Then
            bProcesoOK = False: GoTo ErrorManager
        End If
    Else
        If Not oFactOrdenesBienes.modificar(oDOFactOrdenBien) Then
            bProcesoOK = False: GoTo ErrorManager
        End If
    End If
    
    'Cargar los servicios agregados y modificados
    
    rsFacturacionProductosBien.MoveFirst
    Do While Not rsFacturacionProductosBien.EOF
    
        oDoFacturacionBienes.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
        oDoFacturacionBienes.IdFacturacionBienes = IIf(IsNull(rsFacturacionProductosBien!IdFacturacionProducto), 0, rsFacturacionProductosBien!IdFacturacionProducto)
        oDoFacturacionBienes.IdOrden = oDOFactOrdenBien.IdOrden
        oDoFacturacionBienes.idProducto = IIf(IsNull(rsFacturacionProductosBien!idProducto), 0, rsFacturacionProductosBien!idProducto)
        oDoFacturacionBienes.PrecioUnitario = IIf(IsNull(rsFacturacionProductosBien!PrecioUnitario), 0, rsFacturacionProductosBien!PrecioUnitario)
        oDoFacturacionBienes.cantidad = IIf(IsNull(rsFacturacionProductosBien!cantidad), 0, rsFacturacionProductosBien!cantidad)
        'oDoFacturacionBienes.IdEstadoFacturacion = 4 'Pagado  'IIf(IsNull(rsFacturacionProductosBien!IdEstadoFacturacion), 0, rsFacturacionProductosBien!IdEstadoFacturacion)
        oDoFacturacionBienes.IdEstadoFacturacion = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
        oDoFacturacionBienes.idFuenteFinanciamiento = IIf(IsNull(rsFacturacionProductosBien!idFuenteFinanciamiento), 0, rsFacturacionProductosBien!idFuenteFinanciamiento)
        oDoFacturacionBienes.IdTipoFinanciamiento = IIf(IsNull(rsFacturacionProductosBien!IdTipoFinanciamiento), 0, rsFacturacionProductosBien!IdTipoFinanciamiento)
        oDoFacturacionBienes.IdPuntoCarga = IIf(IsNull(rsFacturacionProductosBien!IdPuntoCarga), 0, rsFacturacionProductosBien!IdPuntoCarga)
        oDoFacturacionBienes.IdUsuarioAutorizaSeguro = IIf(IsNull(rsFacturacionProductosBien!IdUsuarioAutorizaSeguro), 0, rsFacturacionProductosBien!IdUsuarioAutorizaSeguro)
        oDoFacturacionBienes.IdUsuarioAutorizaPendiente = IIf(IsNull(rsFacturacionProductosBien!IdUsuarioAutorizaPendiente), 0, rsFacturacionProductosBien!IdUsuarioAutorizaPendiente)
        oDoFacturacionBienes.IdUsuarioAutorizaDevolucion = IIf(IsNull(rsFacturacionProductosBien!IdUsuarioAutorizaDevolucion), 0, rsFacturacionProductosBien!IdUsuarioAutorizaDevolucion)
        oDoFacturacionBienes.IdCajero = oDOGestionCaja.IdCajero 'IIf(IsNull(rsFacturacionProductosBien!IdCajero), 0, rsFacturacionProductosBien!IdCajero)
        oDoFacturacionBienes.FechaAutorizaSeguro = IIf(IsNull(rsFacturacionProductosBien!FechaAutorizaSeguro), 0, rsFacturacionProductosBien!FechaAutorizaSeguro)
        oDoFacturacionBienes.FechaAutorizaPendiente = IIf(IsNull(rsFacturacionProductosBien!FechaAutorizaPendiente), 0, rsFacturacionProductosBien!FechaAutorizaPendiente)
        oDoFacturacionBienes.FechaAutorizaDevolucion = IIf(IsNull(rsFacturacionProductosBien!FechaAutorizaDevolucion), 0, rsFacturacionProductosBien!FechaAutorizaDevolucion)
        oDoFacturacionBienes.FechaCajero = Now 'IIf(IsNull(rsFacturacionProductosBien!FechaCajero), 0, rsFacturacionProductosBien!FechaCajero)
        oDoFacturacionBienes.idAtencion = IIf(IsNull(rsFacturacionProductosBien!idAtencion), 0, rsFacturacionProductosBien!idAtencion)
        oDoFacturacionBienes.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
            
        Select Case rsFacturacionProductosBien!EstadoLocal
        Case "A"
            If Not oFacturacionBienes.InsertarBienInsumoSP(oDoFacturacionBienes, sResultado) Then
                If sResultado <> "" Then
                    MsgBox sResultado, vbInformation, "Facturacion Servicios"
                End If
                bProcesoOK = False: GoTo ErrorManager
            End If
        Case Else   'Sea M o L igual se modificar al nro de comprobante y el estado
            If Not oFacturacionBienes.ModificarBienInsumoSP(oDoFacturacionBienes, sResultado) Then
                MsgBox sResultado, vbInformation, "Facturacion Servicios"
                bProcesoOK = False: GoTo ErrorManager
            End If
        End Select
        
        rsFacturacionProductosBien.MoveNext
    Loop
        
    'elimina servicios
    Dim I As Integer
    Dim lIdFacturacionBien As Long
    For I = 1 To mo_ProductosBienEliminados.Count
        lIdFacturacionBien = Val(mo_ProductosBienEliminados.Item(I))
        oDoFacturacionBienes.IdFacturacionBienes = lIdFacturacionBien
        oDoFacturacionBienes.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        If Not oFacturacionBienes.Eliminar(oDoFacturacionBienes) Then
            bProcesoOK = False: GoTo ErrorManager
        End If
    Next
            
    Dim oDoCajaNroDocumento As New DOCajaNroDocumento
    oDoCajaNroDocumento.idCaja = oDOGestionCaja.idCaja
    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
    Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOGestionCaja.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
    If Not oCajaNroDocumento.modificar(oDoCajaNroDocumento) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
    
    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = ""
    CajaComprobantePagoBienInsumoAgregar = True
ErrorManager:
    If CajaComprobantePagoBienInsumoAgregar = False Then
        oConexion.RollbackTrans
        oConexion.Close
        Set oConexion = Nothing
    End If
    Set oDoFacturacionBienes = Nothing
    Set oFacturacionBienes = Nothing
    Set oFacturacionSeguros = Nothing
    Set oFactOrdenesBienes = Nothing
    Set oCajaComprobantePago = Nothing
    Set oCajaNroDocumento = Nothing
End Function



Function CajaComprobantePagoAnulacionOrdenServicio(lIdComprobantePago As Long, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim bProcesoOK As Boolean

    CajaComprobantePagoAnulacionOrdenServicio = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oCajaComprobantePago.Conexion = oConexion
    
    If Not oCajaComprobantePago.AnularComprobanteYOrdenServicio(lIdComprobantePago, ml_IdUsuarioAuditoria) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
                
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagoAnulacionOrdenServicio = True
    Set oConexion = Nothing
    Set oCajaComprobantePago = Nothing

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    CajaComprobantePagoAnulacionOrdenServicio = False
End Function
Function CajaComprobantePagoAnulacionOrdenBienInsumo(lIdComprobantePago As Long, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim bProcesoOK As Boolean

    CajaComprobantePagoAnulacionOrdenBienInsumo = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oCajaComprobantePago.Conexion = oConexion
    
    If Not oCajaComprobantePago.AnularComprobanteYOrdenBienInsumo(lIdComprobantePago, ml_IdUsuarioAuditoria) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
                
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagoAnulacionOrdenBienInsumo = True

    Set oConexion = Nothing
    Set oCajaComprobantePago = Nothing
    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    CajaComprobantePagoAnulacionOrdenBienInsumo = False
End Function

Function CajaComprobantePagodevolucionOrdenServicio(lIdComprobantePago As Long, oDOCajaComprobantePagoDeDevolucion As DOCajaComprobantesPago, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim bProcesoOK As Boolean

    CajaComprobantePagodevolucionOrdenServicio = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oCajaComprobantePago.Conexion = oConexion
    
    oDOCajaComprobantePagoDeDevolucion.idEstadoComprobante = 4  'PAGADO
    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePagoDeDevolucion) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
    
    If Not oCajaComprobantePago.DevolverComprobanteYOrdenServicio(lIdComprobantePago, oDOCajaComprobantePagoDeDevolucion.IdComprobantePago, ml_IdUsuarioAuditoria) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
                
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagodevolucionOrdenServicio = True
    Set oConexion = Nothing
    Set oCajaComprobantePago = Nothing

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    CajaComprobantePagodevolucionOrdenServicio = False
End Function

Function CajaComprobantePagoDevolucionOrdenBienInsumo(lIdComprobantePago As Long, oDOCajaComprobantePagoDeDevolucion As DOCajaComprobantesPago, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim bProcesoOK As Boolean

    CajaComprobantePagoDevolucionOrdenBienInsumo = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oCajaComprobantePago.Conexion = oConexion
    
    oDOCajaComprobantePagoDeDevolucion.idEstadoComprobante = 4  'PAGADO
    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePagoDeDevolucion) Then
        bProcesoOK = False: GoTo ErrorManager
    End If

    If Not oCajaComprobantePago.DevolverComprobanteYOrdenBienInsumo(lIdComprobantePago, oDOCajaComprobantePagoDeDevolucion.IdComprobantePago, ml_IdUsuarioAuditoria) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
                
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagoDevolucionOrdenBienInsumo = True
    Set oConexion = Nothing
    Set oCajaComprobantePago = Nothing

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    CajaComprobantePagoDevolucionOrdenBienInsumo = False
End Function

Function ObtenerNroDocumento(nroDocumentoActual As String) As String
    Dim nroDocLong As Long
    Dim nroNuevoDocLong As Long
    Dim nroNuevoDoc As String
    Dim longitud  As Integer
    Dim longitudNuevo  As Integer
    Dim I As Integer
    
    longitud = Len(Trim(nroDocumentoActual))
    nroDocLong = Val(nroDocumentoActual)
    nroNuevoDocLong = nroDocLong + 1
    nroNuevoDoc = CStr(nroNuevoDocLong)
    longitudNuevo = Len(nroNuevoDoc)
    
    For I = 1 To longitud - longitudNuevo
        nroNuevoDoc = "0" & nroNuevoDoc
    Next I
    ObtenerNroDocumento = nroNuevoDoc
    
End Function

Function CajeroPerteneceACaja(idEmpleado As Long, idCaja As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oCajero As New CajaCajero
    Dim oDOCajero As New DOCajaCajero
    
    CajeroPerteneceACaja = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajero.Conexion = oConexion
    oDOCajero.idEmpleado = idEmpleado
    oDOCajero.idCaja = idCaja
    
    CajeroPerteneceACaja = oCajero.SeleccionarSegunUsuarioYCaja(oDOCajero)
 
    oConexion.Close
    ms_MensajeError = oCajero.MensajeError
    Set oConexion = Nothing
    Set oCajero = Nothing
    
End Function


Function RetornaCajaAbierta(IdCajero As Long, idCaja As Long, IdTurno As Long) As doCajaGestion
Dim oConexion As New ADODB.Connection
Dim oCajaGestion As New CajaGestion
Dim oRecordset As New ADODB.Recordset
Dim oDoCajaGestion As New doCajaGestion
    
    Set RetornaCajaAbierta = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajaGestion.Conexion = oConexion
    
    Set oRecordset = oCajaGestion.CajaGestionPorCajeroCajaYTurno(idCaja, IdCajero, IdTurno)
    
    If (oRecordset.RecordCount <= 0) Then
        Exit Function
    End If
    
    With oDoCajaGestion
        .TotalCobrado = IIf(IsNull(oRecordset!TotalCobrado), "", oRecordset!TotalCobrado)
        .FechaCierre = IIf(IsNull(oRecordset!FechaCierre), 0, oRecordset!FechaCierre)
        .IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
        .idCaja = IIf(IsNull(oRecordset!idCaja), 0, oRecordset!idCaja)
        .IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
        .EstadoLote = IIf(IsNull(oRecordset!EstadoLote), "", oRecordset!EstadoLote)
        .FechaApertura = IIf(IsNull(oRecordset!FechaApertura), 0, oRecordset!FechaApertura)
        .IdGestionCaja = IIf(IsNull(oRecordset!IdGestionCaja), 0, oRecordset!IdGestionCaja)
    End With
        
    Set RetornaCajaAbierta = oDoCajaGestion
    
    oConexion.Close
    ms_MensajeError = oCajaGestion.MensajeError
    Set oConexion = Nothing
    Set oCajaGestion = Nothing
    Set oRecordset = Nothing
    
End Function


Function ServiciosFiltrarParaCajero(sNombre As String, lIdTipoCatalogo As Long, lIdPuntoCarga As Long) As ADODB.Recordset

Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
        
    Set ServiciosFiltrarParaCajero = oTabla.SeleccionarServiciosLike(sNombre, lIdTipoCatalogo, lIdPuntoCarga)
        
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function BienesFiltrarParaCajero(sNombre As String, lTipoCatalogo As Long, lIdPuntoCarga As Long, _
                                 Optional lbSoloItemsDePaquete As Boolean) As ADODB.Recordset

Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set BienesFiltrarParaCajero = oTabla.SeleccionarBienesLike(sNombre, lTipoCatalogo, lIdPuntoCarga, lbSoloItemsDePaquete)
        
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'***************daniel barrantes**************
'***************Filtrar datos de Boleta, solo Medicamentos de Farmacia
'***************segun Nro Serie y Nro Boleta
'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function ComprobanteFarmaciaSeleccionarPorSerieDocumento(lnNroSerie As String, lnNroDcto As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oConexion As New ADODB.Connection
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 300
oConexion.Open sighentidades.CadenaConexion
         With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "ComprobanteFarmaciaSeleccionarPorSerieDocumento"
          Set oParameter = .CreateParameter("@NroSerie", adInteger, adParamInput, 0, lnNroSerie): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@NroDocumento", adInteger, adParamInput, 0, lnNroDcto): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        
    Set ComprobanteFarmaciaSeleccionarPorSerieDocumento = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Exit Function
End Function


'***************daniel barrantes**************
'***************Filtrar datos de Boleta, solo Servicios
'***************segun Nro Serie y Nro Boleta
'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function ComprobanteServicioSeleccionarPorSerieDocumento(lnNroSerie As String, lnNroDcto As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oConexion As New ADODB.Connection
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "ComprobanteServicioSeleccionarPorSerieDocumento"
      Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 7, lnNroDcto): .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
     
    Set ComprobanteServicioSeleccionarPorSerieDocumento = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function SeleccionaDatosCajero(lnIdCajero As Long, lnDatoDelEmpleado As sghDatoDelEmpleado, Optional oConexion1 As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
'Dim ssql As String
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
      .CommandText = "SeleccionaDatosCajero"
      Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, lnIdCajero): .Parameters.Append oParameter
      Set oRecordset = .Execute
      If oConexion1 Is Nothing Then
         Set oRecordset.ActiveConnection = Nothing
      End If
    End With
  
     If oRecordset.RecordCount > 0 Then
        Select Case lnDatoDelEmpleado
        Case sghIniciales
            SeleccionaDatosCajero = UCase(Left(oRecordset.Fields!ApellidoPaterno, 1) & Left(oRecordset.Fields!ApellidoMaterno, 1) & Left(oRecordset.Fields!Nombres, 1))
        Case sghUsuario
            SeleccionaDatosCajero = oRecordset.Fields!Usuario
        Case sghApellidosYnombres
            SeleccionaDatosCajero = Trim(oRecordset.Fields!ApellidoPaterno) & " " & Trim(oRecordset.Fields!ApellidoMaterno) & " " & Trim(oRecordset.Fields!Nombres)
        End Select
     Else
        SeleccionaDatosCajero = ""
     End If
     If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
     End If
     Set oConexion = Nothing
     Set oCommand = Nothing
     
     Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Resume
   Exit Function
End Function


'***************daniel barrantes**************
'***************Retorna el NOMBRE DEL SERVICIO donde Ingreso el Paciente
'***************segun segun su ID de Cuenta
'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function NombreServicioPorCuentaAtencion(lnIdCuentaAtencion As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As New ADODB.Parameter
'Dim ssql As String
Dim oConexion As New ADODB.Connection
    NombreServicioPorCuentaAtencion = ""
        If lnIdCuentaAtencion > 0 Then

            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
  
            With oCommand
                '.CommandType = adCmdText
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "NombreServicioPorCuentaAtencion"
                Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
                Set oRecordset = .Execute
                Set oParameter = Nothing
                If oRecordset.RecordCount > 0 And Not IsNull(oRecordset.Fields!codServicio) Then
                    NombreServicioPorCuentaAtencion = oRecordset.Fields!codServicio & "-" & oRecordset.Fields!dservicio
                    If oRecordset.Fields!idTipoServicio = 2 Or oRecordset.Fields!idTipoServicio = 3 Or oRecordset.Fields!idTipoServicio = 4 Then
                        oRecordset.Close
                        .CommandText = "NombreServicioPorCuentaAtencionExiste"
                        Set oRecordset = .Execute
                        NombreServicioPorCuentaAtencion = oRecordset.Fields!codServicio & "-" & oRecordset.Fields!dservicio
                    End If
                End If
                Set oRecordset.ActiveConnection = Nothing
                oConexion.Close
            End With
        End If
        Set oConexion = Nothing
      Set oCommand = Nothing
  Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Exit Function
           
End Function


Function SeleccionarOrdenesPendientesDePago(ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarOrdenesPendientesDePago = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionarOrdenesPendientesDePago"
        Set oParameter = .CreateParameter("@FechaInicio", adDate, adParamInput, 10, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDate, adParamInput, 10, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarOrdenesPendientesDePago = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

'sunat
Function CajaComprobantePagoSeleccionarPorFechaOdocumento(lnNroSerie As String, lnNroDcto As String, ldFechaInicio As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoSeleccionarPorFechaOdocumento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFechas"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoSeleccionarPorFechaOdocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'sunat

Function CajaComprobantePagoProductosPorNroSerieNroDocumento(lnNroSerie As String, lnNroDcto As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoProductosPorNroSerieNroDocumento = Nothing
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoProductosPorNroSerieNroDocumento"
        'Facturador Sunat
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        'Facturador Sunat
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoProductosPorNroSerieNroDocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(lnNroSerie As String, lnNroDcto As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function FactOrdenServicioPagosPendientesDePagoPorFechas(ldFechaInicio As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPagosPendientesDePagoPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosPendientesDePagoPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDate, adParamInput, 10, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDate, adParamInput, 10, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPagosPendientesDePagoPorFechas = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago(lidOrdenPago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago"
        Set oParameter = .CreateParameter("@idOrdenPago", adInteger, adParamInput, 0, lidOrdenPago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPagosPendientesDePagoPorIdOrdenPago = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function CajaComprobantePagoServiciosPorNroSerieNroDocumento(lnNroSerie As String, lnNroDcto As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoServiciosPorNroSerieNroDocumento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoServiciosPorNroSerieNroDocumento"
        'Facturador Sunat
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        'Facturador Sunat
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoServiciosPorNroSerieNroDocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantePagoServicioAnulaBoleta(oDOCajaComprobantePago As DOCajaComprobantesPago, _
                  mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lnIdOrden_QueQueda As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicioPagos As New FactOrdenServicioPagos
Dim oDoFactOrdenServPagos As New DoFactOrdenServPagos
Dim oDoFactOrdenServ As New DoFactOrdenServ
Dim oFactOrdenServicio As New FactOrdenServicio
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oCajaComprobantePago As New CajaComprobantesPago
Dim oCitas As New Citas
Dim oDOCita As New doCita
Dim lcCorrelativo As String
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim lcTicket As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lnItem As Long
Dim oRsListaOrdenes As New Recordset
Dim oRsListaOrdenesDetalle As New Recordset
Dim oRsCita As New Recordset
Dim rs As New Recordset
Dim ms_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim lnIdOrdenPago As Long
Dim oRsTmp As New Recordset
Dim lnIdReceta As Long
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    
    CajaComprobantePagoServicioAnulaBoleta = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFactOrdenServicio.Conexion = oConexion
    Set oCajaComprobantePago.Conexion = oConexion
    Set oFactOrdenServicioPagos.Conexion = oConexion
    Set oCitas.Conexion = oConexion
    
    '*********comprobante pago
    oDOCajaComprobantePago.idEstadoComprobante = 9   'anulada
    If Not oCajaComprobantePago.modificar(oDOCajaComprobantePago) Then
        bProcesoOK = False: GoTo ErrorManager1El
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "E", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
    
    Set oRsListaOrdenes = FactOrdenServicioPagosSeleccionarPorIdComprobantePago(oDOCajaComprobantePago.IdComprobantePago)
    If oRsListaOrdenes.RecordCount > 0 Then
        oRsListaOrdenes.MoveFirst
        Do While Not oRsListaOrdenes.EOF

            oDoFactOrdenServPagos.idOrdenPago = oRsListaOrdenes.Fields!idOrdenPago
            If Not oFactOrdenServicioPagos.SeleccionarPorId(oDoFactOrdenServPagos) Then
               bProcesoOK = False: GoTo ErrorManager1El
            End If
            oDoFactOrdenServPagos.IdEstadoFacturacion = 9  'Anulada
            oDoFactOrdenServPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
            oDoFactOrdenServPagos.IdOrden = lnIdOrden_QueQueda
            If Not oFactOrdenServicioPagos.modificar(oDoFactOrdenServPagos) Then
               bProcesoOK = False: GoTo ErrorManager1El
            End If
            oDoFactOrdenServ.IdOrden = oRsListaOrdenes.Fields!IdOrden
            If Not oFactOrdenServicio.SeleccionarPorId(oDoFactOrdenServ) Then
               bProcesoOK = False: GoTo ErrorManager1El
            End If
            If oDoFactOrdenServ.IdPuntoCarga = 99 Then
               'Se registro desde CAJA
               oDoFactOrdenServ.IdEstadoFacturacion = 9   'anulada
            Else
               oDoFactOrdenServ.IdEstadoFacturacion = IIf(oDOCajaComprobantePago.idCuentaAtencion > 0, 1, 9) 'Registrada / Anulada
            End If
            oDoFactOrdenServ.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
            If Not oFactOrdenServicio.modificar(oDoFactOrdenServ) Then
               bProcesoOK = False: GoTo ErrorManager1El
            End If
            
            '***********Solo cuando tiene Nro cuenta, devuelve Cantidades financiadas
            If oDOCajaComprobantePago.idCuentaAtencion > 0 Then
               'agregar Cabecera Pendiente de Pago
               oDoFactOrdenServPagos.FechaCreacion = oRsListaOrdenes.Fields!FechaCreacion
               If oDoFactOrdenServ.IdPuntoCarga = 99 Then
                  'Se registro desde CAJA
                  oDoFactOrdenServPagos.IdEstadoFacturacion = 9  'anulado
               Else
                  oDoFactOrdenServPagos.IdEstadoFacturacion = 1  'registrado
               End If
               oDoFactOrdenServPagos.IdOrden = oRsListaOrdenes.Fields!IdOrden
               oDoFactOrdenServPagos.idUsuario = oDOCajaComprobantePago.IdUsuarioAuditoria
               oDoFactOrdenServPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
               oDoFactOrdenServPagos.IdComprobantePago = 0
               If Not oFactOrdenServicioPagos.Insertar(oDoFactOrdenServPagos) Then
                    bProcesoOK = False: GoTo ErrorManager1El
               End If
               'agregar Detalle Pendiente de Pago
               lnIdOrdenPago = oRsListaOrdenes.Fields!idOrdenPago
               Do While Not oRsListaOrdenes.EOF And lnIdOrdenPago = oRsListaOrdenes.Fields!idOrdenPago
                    oDoFacturacionServicioPagos.cantidad = oRsListaOrdenes.Fields!cantidad
                    oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
                    oDoFacturacionServicioPagos.idProducto = oRsListaOrdenes.Fields!idProducto
                    oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
                    oDoFacturacionServicioPagos.Precio = oRsListaOrdenes.Fields!Precio
                    oDoFacturacionServicioPagos.total = oRsListaOrdenes.Fields!total
                    If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                        bProcesoOK = False: GoTo ErrorManager1El
                    End If
                    oRsListaOrdenes.MoveNext
                    If oRsListaOrdenes.EOF Then
                       Exit Do
                    End If
               Loop
            Else
               lnIdOrdenPago = oRsListaOrdenes.Fields!idOrdenPago
               Do While Not oRsListaOrdenes.EOF And lnIdOrdenPago = oRsListaOrdenes.Fields!idOrdenPago
                  oRsListaOrdenes.MoveNext
                  If oRsListaOrdenes.EOF Then
                     Exit Do
                  End If
               Loop
            End If
        Loop
    End If
    oRsListaOrdenes.Close
    'pone CITA como REGISTRADO
    Set oRsCita = ms_ReglasAdmision.CitasSeleccionarPorNroCuenta(oDOCajaComprobantePago.idCuentaAtencion)
    If oRsCita.RecordCount > 0 Then
       If oRsCita.Fields!idTipoServicio = 1 Then
            oDOCita.IdCita = oRsCita.Fields!IdCita
            If Not oCitas.SeleccionarPorId(oDOCita) Then
                bProcesoOK = False: GoTo ErrorManager1El
            End If
            oDOCita.IdEstadoCita = 1   'registrado
            oDOCita.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
            If Not oCitas.modificar(oDOCita) Then
                bProcesoOK = False: GoTo ErrorManager1El
            End If
            MsgBox "Nueva Nro Orden Pago: " & oDoFactOrdenServPagos.idOrdenPago & Chr(13) & Chr(13) & " o busque pulsando clic en botón PUNTOS"
       End If
    End If

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "ConsultaCajaComprobantePagoServicio"
      Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, oDOCajaComprobantePago.IdComprobantePago): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
    Set oParameter = Nothing
    Set oCommand = Nothing
    'Receta automatica - debb-03/04/2012
    'mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalle 0, 0, oDOCajaComprobantePago.IdComprobantePago, oConexion
    mo_ReglasImagenes.ActualizaRecetaCabeceraYDetalleSoloBoletaAnulada oDOCajaComprobantePago.IdComprobantePago, oConexion
    
    Set oRsTmp = Nothing

    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    CajaComprobantePagoServicioAnulaBoleta = True
ErrorManager1El:
    If CajaComprobantePagoServicioAnulaBoleta = False Then
        oConexion.RollbackTrans
        oConexion.Close
        ms_MensajeError = Err.Description
    End If
    Set oConexion = Nothing
    Set oFactOrdenServicioPagos = Nothing
    Set oDoFactOrdenServPagos = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenServicio = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oCajaComprobantePago = Nothing
    Set oRsListaOrdenes = Nothing
    Set rs = Nothing
    Set oRsCita = Nothing
    Set oRsListaOrdenesDetalle = Nothing
    Set mo_ReglasImagenes = Nothing
End Function
 

Function FactOrdenServicioPagosSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPagosSeleccionarPorIdComprobantePago = Nothing
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
        .CommandText = "FactOrdenServicioPagosSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FactOrdenServicioPagosSeleccionarPorIdComprobantePago = oRecordset
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

Function SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(LnTipoFiltro As sghTipoOrden, lcFiltro As String, lnIdtipoFinanciamiento As Long, lnIdPuntoCarga As Long, lnTipoServicioAfiltrar As sghTipoServicioOfrecidos) As ADODB.Recordset
    Dim oTabla As New CatalogoServicios
    Set SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga = oTabla.SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(LnTipoFiltro, lcFiltro, lnIdtipoFinanciamiento, lnIdPuntoCarga, lnTipoServicioAfiltrar)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function FactCatalogoServiciosSeleccionarPorCodigoOnombre(lcCodigo As String, lcNombre As String, Optional oConexion As Connection) As Recordset
    Dim oTabla As New CatalogoServicios
    Set FactCatalogoServiciosSeleccionarPorCodigoOnombre = oTabla.FactCatalogoServiciosSeleccionarPorCodigoOnombre(lcCodigo, lcNombre, oConexion)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function FactCatalogoServicioDevuelvePrecioUnitario(lnIdProducto As Long, lnIdtipoFinanciamiento As Long) As Double
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
      .CommandText = "FactCatalogoServicioDevuelvePrecioUnitario"
      Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
      Set oRsTmp = .Execute
      Set oRsTmp.ActiveConnection = Nothing
    End With
  
    If oRsTmp.RecordCount > 0 Then
       FactCatalogoServicioDevuelvePrecioUnitario = oRsTmp.Fields!PrecioUnitario
    End If
    oRsTmp.Close
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
End Function

Function SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga(LnTipoFiltro As sghTipoOrden, lcFiltro As String, lnIdPuntoCarga As Long, lnTipoServicioAfiltrar As sghTipoServicioOfrecidos) As ADODB.Recordset
    Dim oTabla As New CatalogoServicios
    Set SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga = oTabla.SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga(LnTipoFiltro, lcFiltro, lnIdPuntoCarga, lnTipoServicioAfiltrar)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

Function CajaComprobantesPagoSeleccionarPorFechas(ldFechaInicio As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoSeleccionarPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantesPagoSeleccionarPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoSeleccionarPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoSeleccionarPorFechasSoloFarmacia(lcFechaInicio As String, lcHoraInicio As String, lcFechaFin As String, lcHoraFin As String, lbTodasFarmacias As Integer, lnIdAlmacen As Long) As Recordset
      Dim oRecordset As New ADODB.Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
      Set CajaComprobantesPagoSeleccionarPorFechasSoloFarmacia = Nothing
    
      oConexion.CursorLocation = adUseClient
      oConexion.CommandTimeout = 300
      oConexion.Open sighentidades.CadenaConexion

    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "CajaComprobantesPagoSeleccionarPorFechasSoloFarmacia"
      Set oParameter = .CreateParameter("@FechaInicio", adVarChar, adParamInput, 50, lcFechaInicio): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@HoraInicio", adVarChar, adParamInput, 50, lcHoraInicio): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@FechaFin", adVarChar, adParamInput, 50, lcFechaFin): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@HoraFin", adVarChar, adParamInput, 50, lcHoraFin): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@TodasFarmacias", adInteger, adParamInput, 0, lbTodasFarmacias): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
    Set CajaComprobantesPagoSeleccionarPorFechasSoloFarmacia = oRecordset
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

Function FactOrdenServicioSeleccionarPuntoCargaPorIdOrden(lnIdOrden As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioSeleccionarPuntoCargaPorIdOrden = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPuntoCargaPorIdOrden"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPuntoCargaPorIdOrden = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function CajaComprobantesSeleccionarPorCuentaAtencion(lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesSeleccionarPorCuentaAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesSeleccionarPorCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesSeleccionarPorCuentaAtencion = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function CajaComprobantesSeleccionarPorId(lnIdComprobantePago As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesSeleccionarPorId = Nothing
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
        .CommandText = "CajaComprobantesPagoSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
        
   End With
   Set CajaComprobantesSeleccionarPorId = oRecordset
   Set oCommand = Nothing
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oConexion = Nothing
   End If
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta(lnIdCuentaAtencion As Long, lnIdProducto As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function CajaComprobantesPagoSeleccionarAdelantosPorIdCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set CajaComprobantesPagoSeleccionarAdelantosPorIdCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoSeleccionarAdelantosPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set CajaComprobantesPagoSeleccionarAdelantosPorIdCuenta = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function RetornaTotalDescuentosPorAdelantosSegunCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As Double
        Dim oRsTmp1 As New Recordset
        Dim lnAdelantos As Double, lnPagosAdelantados As Double
        'PRIMERA CONSULTA SOBRE LOS PAGOS ADELANTADOS para cargar a lnPagosAdelantados
        'Consulta los pagos a cuenta (adelantos) segun producto 4691 que corresponde a adelantos del idParametro 245
        Set oRsTmp1 = CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta(lnIdCuentaAtencion, Val(lcBuscaParametro.SeleccionaFilaParametro(245)), oConexion)
        lnPagosAdelantados = 0
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnPagosAdelantados = lnPagosAdelantados + oRsTmp1.Fields!total
              oRsTmp1.MoveNext
           Loop
        End If
        'SEGUNDA CONSULTA SOBRE LOS PAGOS ADELANTADOS para cargar a lnAdelantos
        Set oRsTmp1 = CajaComprobantesPagoSeleccionarAdelantosPorIdCuenta(lnIdCuentaAtencion, oConexion)
        lnAdelantos = 0
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnAdelantos = lnAdelantos + IIf(IsNull(oRsTmp1.Fields!Adelantos), 0, oRsTmp1.Fields!Adelantos)
              oRsTmp1.MoveNext
           Loop
        End If
        RetornaTotalDescuentosPorAdelantosSegunCuenta = lnPagosAdelantados - lnAdelantos
        Set oRsTmp1 = Nothing
End Function

Function CajaComprobantePagoPorSerieDocumento(lnNroSerie As String, lnNroDcto As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoPorSerieDocumento = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoPorSerieDocumento"
        'Facturador Sunat
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        'Facturador Sunat
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoPorSerieDocumento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajerosSeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
        Set CajerosSeleccionarTodos = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "CajerosSeleccionarTodos"
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        
        Set CajerosSeleccionarTodos = oRecordset
        oConexion.Close
         Set oConexion = Nothing
         Set oCommand = Nothing
        
        ms_MensajeError = ""
        Exit Function
ManejadorDeError:
        ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
        Exit Function
End Function


Function CajaComprobantesPagoSeleccionarFarmaciaPorFechas(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoSeleccionarFarmaciaPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantesPagoSeleccionarFarmaciaPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoSeleccionarFarmaciaPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function





Function CajaGeneraPreventaServicio(ByRef lnPreventaServicio As Long, lnIdServicio As Long, lnIdtipoFinanciamiento As Long, lnTotalPreventa As Double, rsFacturacionProductos As Recordset, ml_IdUsuarioAuditoria As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oDoFactPreventa As New SIGHcomun.DoFactPreventa
Dim oFactPreventa As New FactPreventa
Dim oDoFacturacionPreVenta As New DoFacturacionPreVenta
Dim oFacturacionPreventa As New FacturacionPreventa
Dim oConexion As New ADODB.Connection
Dim bProcesoOK As Boolean
Dim sResultado As String

    CajaGeneraPreventaServicio = False
    '
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    '
    oConexion.BeginTrans
    Set oFactPreventa.Conexion = oConexion
    Set oFacturacionPreventa.Conexion = oConexion
    '
    oDoFactPreventa.FechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
    oDoFactPreventa.idEstadoPreventa = 1
    oDoFactPreventa.IdServicio = lnIdServicio
    oDoFactPreventa.idUsuario = ml_IdUsuarioAuditoria
    oDoFactPreventa.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
    oDoFactPreventa.total = lnTotalPreventa
    oDoFactPreventa.IdTipoFinanciamiento = lnIdtipoFinanciamiento
    If Not oFactPreventa.Insertar(oDoFactPreventa) Then
        bProcesoOK = False: ms_MensajeError = oFactPreventa.MensajeError: GoTo ErrorManager
    End If
    lnPreventaServicio = oDoFactPreventa.idFactPreventa
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_IdUsuarioAuditoria, "A", oDoFactPreventa.idFactPreventa, "FactPreventa", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
    '
    oDoFacturacionPreVenta.IdUsuarioAuditoria = oDoFactPreventa.IdUsuarioAuditoria
    oDoFacturacionPreVenta.idFactPreventa = oDoFactPreventa.idFactPreventa
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
       If rsFacturacionProductos!idProducto > 0 Then
            oDoFacturacionPreVenta.cantidad = rsFacturacionProductos!cantidad
            oDoFacturacionPreVenta.idProducto = rsFacturacionProductos!idProducto
            oDoFacturacionPreVenta.Importe = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
            oDoFacturacionPreVenta.Precio = rsFacturacionProductos!PrecioUnitario
            If Not oFacturacionPreventa.Insertar(oDoFacturacionPreVenta) Then
                 bProcesoOK = False: ms_MensajeError = oFacturacionPreventa.MensajeError: GoTo ErrorManager
            End If
       End If
       rsFacturacionProductos.MoveNext
    Loop
    '
    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = ""
    CajaGeneraPreventaServicio = True

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oDoFactPreventa = Nothing
    Set oFactPreventa = Nothing
    Set oDoFacturacionPreVenta = Nothing
    Set oFacturacionPreventa = Nothing
End Function




Function CajaComprobantePagoServicioPaqueteAgregar(oDOCajaComprobantePago As DOCajaComprobantesPago, oDOGestionCaja As doCajaGestion, oDoFactOrdenServPagos As DoFactOrdenServPagos, rsFacturacionProductos As Recordset, ml_IdUsuarioAuditoria As Long, oDoAtencion As DOAtencion, lnIdPuntoCarga As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
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
Dim oDOFacturacionPaquetes As New DOFacturacionPaquetes
Dim oFacturacionPaquetes As New FacturacionPaquetes
Dim oCitas As New Citas
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim lnPqteGrupo As Long
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia

    CajaComprobantePagoServicioPaqueteAgregar = False
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oFacturacionServicios.Conexion = oConexion
    Set oFacturacionBienes.Conexion = oConexion
    Set oFacturacionSeguros.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oCajaComprobantePago.Conexion = oConexion
    Set oCajaNroDocumento.Conexion = oConexion
    Set oCitas.Conexion = oConexion
    Set oFactOrdenServPagos.Conexion = oConexion
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactOrdenServ.Conexion = oConexion
    Set oFacturacionPaquetes.Conexion = oConexion
    
    oDOCajaComprobantePago.idEstadoComprobante = IIf(oDOCajaComprobantePago.IdTipoPago = 2, 6, 4) 'Devuelto o Pagado
    oDOCajaComprobantePago.IdTipoOrden = 1  'Servicio
    If Not oCajaComprobantePago.Insertar(oDOCajaComprobantePago) Then
        bProcesoOK = False: ms_MensajeError = oCajaComprobantePago.MensajeError: GoTo ErrorManager
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCajaComprobantePago.IdUsuarioAuditoria, "A", oDOCajaComprobantePago.IdComprobantePago, "CajaComprobantesPago", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")        'ListBarItems.idListItem
    '
    rsFacturacionProductos.Sort = "PqteGrupo"
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
        'Crear orden de servicio
        oDoFactOrdenServPagos.IdEstadoFacturacion = 4  'Pagado
        oDoFactOrdenServPagos.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
        oDoFactOrdenServ.FechaCreacion = oDoFactOrdenServPagos.FechaCreacion
        oDoFactOrdenServ.fechaDespacho = oDoFactOrdenServPagos.FechaCreacion
        oDoFactOrdenServ.idCuentaAtencion = 0
        oDoFactOrdenServ.IdEstadoFacturacion = 4 'PAGADO
        oDoFactOrdenServ.idFuenteFinanciamiento = 1 'Contado
        oDoFactOrdenServ.idPaciente = oDOCajaComprobantePago.idPaciente
        oDoFactOrdenServ.IdPuntoCarga = lnIdPuntoCarga
        oDoFactOrdenServ.IdTipoFinanciamiento = oDOCajaComprobantePago.IdTipoFinanciamiento
        oDoFactOrdenServ.idUsuario = oDoFactOrdenServPagos.IdUsuarioAuditoria
        oDoFactOrdenServ.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
        oDoFactOrdenServ.IdUsuarioDespacho = oDoFactOrdenServPagos.IdUsuarioAuditoria
        oDoFactOrdenServ.IdServicioPaciente = 0
        If Not oFactOrdenServ.Insertar(oDoFactOrdenServ) Then
            bProcesoOK = False: ms_MensajeError = oFactOrdenServ.MensajeError: GoTo ErrorManager
        End If
        oDoFactOrdenServPagos.IdOrden = oDoFactOrdenServ.IdOrden
        If Not oFactOrdenServPagos.Insertar(oDoFactOrdenServPagos) Then
            bProcesoOK = False: ms_MensajeError = oFactOrdenServPagos.MensajeError: GoTo ErrorManager
        End If
        lnPqteGrupo = rsFacturacionProductos.Fields!PqteGrupo
        Do While Not rsFacturacionProductos.EOF And lnPqteGrupo = rsFacturacionProductos.Fields!PqteGrupo
           oDoFacturacionServicioDespacho.cantidad = rsFacturacionProductos!cantidad
           oDoFacturacionServicioDespacho.IdOrden = oDoFactOrdenServ.IdOrden
           oDoFacturacionServicioDespacho.idProducto = rsFacturacionProductos!idProducto
           oDoFacturacionServicioDespacho.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
           oDoFacturacionServicioDespacho.Precio = rsFacturacionProductos!PrecioUnitario
           oDoFacturacionServicioDespacho.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
           If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDespacho) Then
                bProcesoOK = False: ms_MensajeError = oFacturacionServicioDespacho.MensajeError: GoTo ErrorManager
           End If
           oDoFacturacionServicioPagos.cantidad = rsFacturacionProductos!cantidad
           oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
           oDoFacturacionServicioPagos.idProducto = rsFacturacionProductos!idProducto
           oDoFacturacionServicioPagos.IdUsuarioAuditoria = oDoFactOrdenServPagos.IdUsuarioAuditoria
           oDoFacturacionServicioPagos.Precio = rsFacturacionProductos!PrecioUnitario
           oDoFacturacionServicioPagos.total = Round(rsFacturacionProductos!cantidad * rsFacturacionProductos!PrecioUnitario, 2)
           If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                bProcesoOK = False: ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrorManager
           End If
           'oDOFacturacionPaquetes.AtencionId
           oDOFacturacionPaquetes.IdComprobantePago = oDOCajaComprobantePago.IdComprobantePago
           oDOFacturacionPaquetes.idEspecialidadServicio = rsFacturacionProductos!PqteIdEspecialidadServicio
           oDOFacturacionPaquetes.IdFactPaquete = rsFacturacionProductos!PqteIdFactPaquete
           oDOFacturacionPaquetes.idOrdenPago = oDoFactOrdenServPagos.idOrdenPago
           oDOFacturacionPaquetes.idProducto = rsFacturacionProductos!idProducto
           oDOFacturacionPaquetes.IdPuntoCarga = rsFacturacionProductos!PqteIdPuntoCarga
           oDOFacturacionPaquetes.IdUsuarioAuditoria = oDOCajaComprobantePago.IdUsuarioAuditoria
           If Not oFacturacionPaquetes.Insertar(oDOFacturacionPaquetes) Then
              bProcesoOK = False: ms_MensajeError = oFacturacionPaquetes.MensajeError: GoTo ErrorManager
           End If
           rsFacturacionProductos.MoveNext
           If rsFacturacionProductos.EOF Then
              Exit Do
           End If
        Loop
    Loop
            
    Dim oDoCajaNroDocumento As New DOCajaNroDocumento
    oDoCajaNroDocumento.idCaja = oDOCajaComprobantePago.idCaja
    oDoCajaNroDocumento.idTipoComprobante = oDOCajaComprobantePago.idTipoComprobante
    Set oDoCajaNroDocumento = oCajaNroDocumento.SeleccionarPorIdCajaYTipoComprobante(oDOCajaComprobantePago.idCaja, _
                                                                   oDOCajaComprobantePago.idTipoComprobante)
    oDoCajaNroDocumento.NroDocumento = Format(Val(oDOCajaComprobantePago.NroDocumento) + 1, "000000")
    oDoCajaNroDocumento.NroSerie = oDOCajaComprobantePago.NroSerie
    If Not oCajaNroDocumento.modificar(oDoCajaNroDocumento) Then
        bProcesoOK = False: ms_MensajeError = oCajaNroDocumento.MensajeError: GoTo ErrorManager
    End If
    oConexion.CommitTrans
    ms_MensajeError = ""
    CajaComprobantePagoServicioPaqueteAgregar = True
ErrorManager:
    If CajaComprobantePagoServicioPaqueteAgregar = False Then
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oDOFacturacionServicios = Nothing
    Set oDoFacturacionBienes = Nothing
    Set oDOSeguros = Nothing
    Set oFacturacionServicios = Nothing
    Set oFacturacionBienes = Nothing
    Set oFacturacionSeguros = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oCajaComprobantePago = Nothing
    Set oCajaNroDocumento = Nothing
    Set oFactOrdenServPagos = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oDoFacturacionServicioDespacho = Nothing
    Set oDoFactOrdenServ = Nothing
    Set oFactOrdenServ = Nothing
    Set oCitas = Nothing
    Set oDOFacturacionPaquetes = Nothing
    Set oFacturacionPaquetes = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function AtencionesSeleccionarDatosCabeceraBoletaHRCxCuenta(lnIdCuentaAtencion As Long) As Recordset
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
        .CommandText = "AtencionesSeleccionarDatosCabeceraBoletaHRCxCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
      End With
      Set AtencionesSeleccionarDatosCabeceraBoletaHRCxCuenta = oRecordset
      oConexion.Close
      Set oRecordset = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Exit Function
      
ManejadorDeError:
      ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
      Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaGestionSeleccionarXFiltroOrdenadoXFApertura(lcFiltro As String) As Recordset

          On Error GoTo ManejadorDeError
          Dim oRecordset As New ADODB.Recordset
          Dim oCommand As New ADODB.Command
          Dim oParameter As ADODB.Parameter
          Dim oConexion As New ADODB.Connection
          Dim ms_MensajeError  As String
          Set CajaGestionSeleccionarXFiltroOrdenadoXFApertura = Nothing
          ms_MensajeError = ""

          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open sighentidades.CadenaConexion

          With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "CajaGestionSeleccionarXFiltroOrdenadoXFApertura"
                Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
                Set oRecordset = .Execute
                Set oRecordset.ActiveConnection = Nothing
          End With
          Set CajaGestionSeleccionarXFiltroOrdenadoXFApertura = oRecordset
          oConexion.Close
          'Set oRecordSet = Nothing
          Set oConexion = Nothing
          Set oCommand = Nothing
          Exit Function
          
ManejadorDeError:
          ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
          Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantePagoSeleccionarXidCajaGestion(lnIdCajaGestion As Long) As Recordset
          On Error GoTo ManejadorDeError
          Dim oRecordset As New ADODB.Recordset
          Dim oCommand As New ADODB.Command
          Dim oParameter As ADODB.Parameter
          Dim oConexion As New ADODB.Connection
          Dim ms_MensajeError  As String
          Set CajaComprobantePagoSeleccionarXidCajaGestion = Nothing
          ms_MensajeError = ""
        
          oConexion.CursorLocation = adUseClient
          oConexion.CommandTimeout = 300
          oConexion.Open sighentidades.CadenaConexion
        
          With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "CajaComprobantePagoSeleccionarXidCajaGestion"
            Set oParameter = .CreateParameter("@IdCajaGestion", adInteger, adParamInput, 0, lnIdCajaGestion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
          End With
          Set CajaComprobantePagoSeleccionarXidCajaGestion = oRecordset
          oConexion.Close
          Set oRecordset = Nothing
          Set oConexion = Nothing
          Set oCommand = Nothing
          Exit Function
          
ManejadorDeError:
          ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
          Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantePagoSeleccionarPaquetes(lnIdComprobantePago As Long) As Recordset

      On Error GoTo ManejadorDeError
      Dim oRecordset As New ADODB.Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
      Dim ms_MensajeError  As String
      Set CajaComprobantePagoSeleccionarPaquetes = Nothing
      ms_MensajeError = ""
    
      oConexion.CursorLocation = adUseClient
      oConexion.CommandTimeout = 300
      oConexion.Open sighentidades.CadenaConexion
    
      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantePagoSeleccionarPaquetes"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
      End With
      Set CajaComprobantePagoSeleccionarPaquetes = oRecordset
      oConexion.Close
      Set oRecordset = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Exit Function
      
ManejadorDeError:
      ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
      Exit Function
         
End Function



'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaActualizaDatosDeFactura(lnSubTotal As Double, lcRuc As String, lcRazonSocial As String, lnIgv As Double, _
                                lnTotal As Double, lcFechaCobranza As String, lnIdComprobantePagoActual As Long, _
                                lcTieneCredito As String)
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
        .CommandText = "CajaActualizaDatosDeFactura"
        Set oParameter = .CreateParameter("@SubTotal", adCurrency, adParamInput, 0, FormatCurrency(lnSubTotal, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Ruc", adChar, adParamInput, 11, lcRuc): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 50, lcRazonSocial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Igv", adCurrency, adParamInput, 0, FormatCurrency(lnIgv, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, FormatCurrency(lnTotal, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, Format(lcFechaCobranza, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdComprobantePagoActual", adInteger, adParamInput, 0, lnIdComprobantePagoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TieneCredito", adVarChar, adParamInput, 1, IIf(lcTieneCredito = "", Null, lcTieneCredito)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing

End Sub

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function FactOrdenServicioPagosSeleccionarXidComprobantePago(lnIdComprobantePagoActual As Long) As Recordset
      On Error GoTo ManejadorDeError
      Dim oRecordset As New ADODB.Recordset
      Dim oCommand As New ADODB.Command
      Dim oParameter As ADODB.Parameter
      Dim oConexion As New ADODB.Connection
      Dim ms_MensajeError  As String
      Set FactOrdenServicioPagosSeleccionarXidComprobantePago = Nothing
      ms_MensajeError = ""
    
      oConexion.CursorLocation = adUseClient
      oConexion.CommandTimeout = 300
      oConexion.Open sighentidades.CadenaConexion
    
      With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosSeleccionarXidComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePagoActual", adInteger, adParamInput, 0, lnIdComprobantePagoActual): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
      End With
      Set FactOrdenServicioPagosSeleccionarXidComprobantePago = oRecordset
      oConexion.Close
      Set oRecordset = Nothing
      Set oConexion = Nothing
      Set oCommand = Nothing
      Exit Function
      
ManejadorDeError:
      ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
      Exit Function
End Function

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaActualizaDatosDeServicioYServicioPago(lnIdComprobantePagoActual As Long, lnIdOrden As Long)
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
    .CommandText = "CajaActualizaDatosDeServicioYServicioPago"
    Set oParameter = .CreateParameter("@IdComprobantePagoActual", adVarChar, adParamInput, 10, lnIdComprobantePagoActual): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    .Execute
  End With
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaReembolsosAnterioresXfactura(lcSerie As String, lcNumero As String, lnIdTipoComprobante As Long) As Recordset
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
    .CommandText = "CajaReembolsosAnterioresXfactura"
    Set oParameter = .CreateParameter("@Serie", adVarChar, adParamInput, 50, lcSerie): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Numero", adVarChar, adParamInput, 50, lcNumero): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, lnIdTipoComprobante): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set CajaReembolsosAnterioresXfactura = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
        
End Function

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoActualizaFarmaciaAddServicio(lnReembolsoPagadoFarmacia As Double, lnIdOrdenPago As Long, _
                                              lnIdOrden As Long, lnReembolsoPagadoServicio As Double)
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
        .CommandText = "CajaReembolsoActualizaFarmaciaAddServicio"
        Set oParameter = .CreateParameter("@ReembolsoPagadoFarmacia", adCurrency, adParamInput, 0, FormatCurrency(lnReembolsoPagadoFarmacia, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ReembolsoPagadoServicio", adCurrency, adParamInput, 0, FormatCurrency(lnReembolsoPagadoServicio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoFarmacia", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(252)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoServicio", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(251)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoActualizaServicioAddFarmacia(lnIdOrdenPago As Long, _
                                              lnIdOrden As Long, lnReembolsoPagadoServicio As Double)
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
        .CommandText = "CajaReembolsoActualizaServicioAddFarmacia"
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ReembolsoPagadoServicio", adCurrency, adParamInput, 0, FormatCurrency(lnReembolsoPagadoServicio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoFarmacia", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(252)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoServicio", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(251)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoActualizaFarmaciaActualizaServicio(lnReembolsoPagadoFarmacia As Double, lnIdOrdenPago As Long, _
                                                   lnIdOrden As Long, lnReembolsoPagadoServicio As Double)
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
        .CommandText = "CajaReembolsoActualizaFarmaciaActualizaServicio"
        Set oParameter = .CreateParameter("@ReembolsoPagadoFarmacia", adCurrency, adParamInput, 0, FormatCurrency(lnReembolsoPagadoFarmacia, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ReembolsoPagadoServicio", adCurrency, adParamInput, 0, FormatCurrency(lnReembolsoPagadoServicio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoFarmacia", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(252)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoServicio", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(251)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoAddFarmacia(lnIdOrdenPago As Long, lnIdOrden As Long, lnTotalFactura As Double)
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
        .CommandText = "CajaReembolsoAddFarmacia"
        Set oParameter = .CreateParameter("@TotalFactura", adCurrency, adParamInput, 0, FormatCurrency(lnTotalFactura, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoFarmacia", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(252)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoActualizaServiciosYfarmacia(lnIdOrdenPago As Long, lnIdOrden As Long, lnTotalFactura As Double)
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
        .CommandText = "CajaReembolsoActualizaServiciosYfarmacia"
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TotalFactura", adCurrency, adParamInput, 0, FormatCurrency(lnTotalFactura, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoEliminaServicio(lnIdOrdenPago As Long, lnIdOrden As Long)

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
    .CommandText = "CajaReembolsoEliminaFarmacia" 'Mismo Store creado en el metodo CajaReembolsoEliminaFarmacia
    Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(251)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    .Execute
  End With
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 26/09/2013 - se cambio a store procedure
Sub CajaReembolsoAddServicio(lnIdOrdenPago As Long, lnIdOrden As Long, lnTotalFactura As Double)
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
        .CommandText = "CajaReembolsoAddServicio"
        Set oParameter = .CreateParameter("@TotalFactura", adCurrency, adParamInput, 0, FormatCurrency(lnTotalFactura, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProductoServicio", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(251)): .Parameters.Append oParameter
        .Execute
      End With
      oConexion.Close
      Set oConexion = Nothing
      Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Sub CajaReembolsoEliminaFarmacia(lnIdOrdenPago As Long, lnIdOrden As Long)
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
    .CommandText = "CajaReembolsoEliminaFarmacia"
    Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(252)): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
    .Execute
  End With
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
End Sub

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaTiposComprobanteSeleccionarTodos(SoloDocumentosSunat As Boolean, soloDocumentosFarmacia As Boolean) As Recordset

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oConexion As New ADODB.Connection
  Set CajaTiposComprobanteSeleccionarTodos = Nothing

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150


    .CommandText = IIf(SoloDocumentosSunat = True, "CajaTiposComprobanteSeleccionarDocSunat", IIf(soloDocumentosFarmacia = True, "CajaTiposComprobanteSeleccionarDocFarmacia", "CajaTiposComprobanteSeleccionarTodos"))
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set CajaTiposComprobanteSeleccionarTodos = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
End Function



'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantesPagosSeleccionarPorIdTipoFinanciamiento(lnIdtipoFinanciamiento As Long) As Recordset

  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New ADODB.Connection
  Set CajaComprobantesPagosSeleccionarPorIdTipoFinanciamiento = Nothing

  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion

  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CajaComprobantesPagosSeleccionarPorIdTipoFinanciamiento"
    Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
  End With
  Set CajaComprobantesPagosSeleccionarPorIdTipoFinanciamiento = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing

End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantePagoProductosPorNroSerieNroDocumentoConexion(lnNroSerie As String, lnNroDcto As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set CajaComprobantePagoProductosPorNroSerieNroDocumentoConexion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantePagoProductosPorNroSerieNroDocumentoConexion"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 10, lnNroDcto): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set CajaComprobantePagoProductosPorNroSerieNroDocumentoConexion = oRecordset
   Set oCommand = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(lnNroSerie As String, lnNroDcto As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 10, lnNroDcto): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion = oRecordset
   Set oCommand = Nothing

    
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function FactOrdenServicioSeleccionarPuntoCargaPorIdOrdenConexion(lnIdOrden As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FactOrdenServicioSeleccionarPuntoCargaPorIdOrdenConexion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPuntoCargaPorIdOrden"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPuntoCargaPorIdOrdenConexion = oRecordset
   Set oCommand = Nothing
   Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function SeleccionaDatosCajeroConexion(lnIdCajero As Long, lnDatoDelEmpleado As sghDatoDelEmpleado, oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter



     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaDatosCajeroConexion"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdCajero): .Parameters.Append oParameter
        Set oRecordset = .Execute
     End With

     If oRecordset.RecordCount > 0 Then
        Select Case lnDatoDelEmpleado
        Case sghIniciales
            SeleccionaDatosCajeroConexion = UCase(Left(oRecordset.Fields!ApellidoPaterno, 1) & Left(oRecordset.Fields!ApellidoMaterno, 1) & Left(oRecordset.Fields!Nombres, 1))
        Case sghUsuario
            SeleccionaDatosCajeroConexion = oRecordset.Fields!Usuario
        Case sghApellidosYnombres
            SeleccionaDatosCajeroConexion = Trim(oRecordset.Fields!ApellidoPaterno) & " " & Trim(oRecordset.Fields!ApellidoMaterno) & " " & Trim(oRecordset.Fields!Nombres)
        End Select
     Else
        SeleccionaDatosCajeroConexion = ""
     End If
     Set oRecordset = Nothing
     Set oCommand = Nothing
     Exit Function
ManejadorDeError:
   MsgBox Err.Description
   Resume
   Exit Function
End Function

Function NombreServicioPorCuentaAtencionConexion(lnIdCuentaAtencion As Long, oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
        NombreServicioPorCuentaAtencionConexion = ""
        If lnIdCuentaAtencion > 0 Then
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "NombreServicioPorCuentaAtencionConexion"
                Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
                Set oRecordset = .Execute
                Set oParameter = Nothing
                If oRecordset.RecordCount > 0 And Not IsNull(oRecordset.Fields!codServicio) Then
                    NombreServicioPorCuentaAtencionConexion = oRecordset.Fields!codServicio & "-" & oRecordset.Fields!dservicio
                    If oRecordset.Fields!idTipoServicio = 2 Or oRecordset.Fields!idTipoServicio = 3 Or oRecordset.Fields!idTipoServicio = 4 Then
                        oRecordset.Close
                        .CommandText = "NombreServicioPorCuentaAtencionConexionExiste"
                        Set oRecordset = .Execute
                        NombreServicioPorCuentaAtencionConexion = oRecordset.Fields!codServicio & "-" & oRecordset.Fields!dservicio
                    End If
                End If
            End With
        End If
        Set oRecordset = Nothing
        Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function AtencionesSeleccionarDatosCabeceraBoletaHRCxCuentaConexion(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "AtencionesSeleccionarDatosCabeceraBoletaHRCxCuentaConexion"
          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set AtencionesSeleccionarDatosCabeceraBoletaHRCxCuentaConexion = oRecordset
        Set oCommand = Nothing
End Function

Function BoletasServicioPorPuntoCarga(ldFechaIni As Date, ldFechaFin As Date, lnIdPuntoCarga As Long) As Recordset
            Dim oRsBoletas As New Recordset
            Dim oRsTmp1 As New Recordset
            Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
            Dim lcBoleta As String, lnTotal As Double, ldFecha As Date, lnIdProducto As Long
            Dim lnIdComprobantePago As Long, lbEsDelPuntoCarga As Boolean
            
            If oRsBoletas.State = 1 Then Set oRsBoletas = Nothing
            With oRsBoletas
                .Fields.Append "IdComprobantePago", adInteger, 4, adFldIsNullable
                .Fields.Append "Boleta", adVarChar, 100, adFldIsNullable
                .Fields.Append "Total", adVarChar, 100, adFldIsNullable
                .Fields.Append "Fecha", adVarChar, 100, adFldIsNullable
                .LockType = adLockOptimistic
                .Open
            End With
            Set oRsTmp1 = mo_ReglasFacturacion.FacturacionServicioPagosPorFechasComprobante(ldFechaIni, ldFechaFin)
            If oRsTmp1.RecordCount > 0 Then
               oRsTmp1.MoveFirst
               Do While Not oRsTmp1.EOF
               
                  lnIdComprobantePago = oRsTmp1.Fields!IdComprobantePago
                  lcBoleta = oRsTmp1.Fields!Boleta
                  lnTotal = oRsTmp1.Fields!total
                  ldFecha = oRsTmp1.Fields!FechaCobranza
                  Do While Not oRsTmp1.EOF And lnIdComprobantePago = oRsTmp1.Fields!IdComprobantePago
                     lnIdProducto = oRsTmp1.Fields!idProducto
                     lbEsDelPuntoCarga = False
                     Do While Not oRsTmp1.EOF And lnIdComprobantePago = oRsTmp1.Fields!IdComprobantePago And lnIdProducto = oRsTmp1.Fields!idProducto
                        If oRsTmp1.Fields!IdPuntoCarga = lnIdPuntoCarga Then
                           lbEsDelPuntoCarga = True
                        End If
                        oRsTmp1.MoveNext
                        If oRsTmp1.EOF Then
                           Exit Do
                        End If
                     Loop
                     If lbEsDelPuntoCarga = False Then
                        If oRsTmp1.EOF Then
                           Exit Do
                        Else
                            Do While Not oRsTmp1.EOF And lnIdComprobantePago = oRsTmp1.Fields!IdComprobantePago
                               oRsTmp1.MoveNext
                               If oRsTmp1.EOF Then
                                  Exit Do
                               End If
                            Loop
                        End If
                     End If
                     If oRsTmp1.EOF Then
                        Exit Do
                     End If
                  Loop
                  If lbEsDelPuntoCarga = True Then
                     oRsBoletas.AddNew
                     oRsBoletas.Fields!IdComprobantePago = lnIdComprobantePago
                     oRsBoletas.Fields!Boleta = lcBoleta
                     oRsBoletas.Fields!fecha = ldFecha
                     oRsBoletas.Fields!total = lnTotal
                     oRsBoletas.Update
                  End If
               Loop
            End If
            oRsTmp1.Close
            Set oRsTmp1 = Nothing
            Set mo_ReglasFacturacion = Nothing
            If oRsBoletas.RecordCount > 0 Then
               oRsBoletas.MoveFirst
            End If
            Set BoletasServicioPorPuntoCarga = oRsBoletas
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset

        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "CajaComprobantesPagoXcuenta"
          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set CajaComprobantesPagoXcuenta = oRecordset
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoXimagenes(lnIdComprobantePago As Long) As Recordset

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
          .CommandText = "CajaComprobantesPagoXimagenes"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaComprobantesPagoXimagenes = oRecordset
        oConexion.Close
        Set oRecordset = Nothing
        Set oCommand = Nothing
        Set oConexion = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoXlaboratorio(lnIdComprobantePago As Long) As Recordset

    
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
          .CommandText = "CajaComprobantesPagoXlaboratorio"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaComprobantesPagoXlaboratorio = oRecordset
        oConexion.Close
        Set oRecordset = Nothing
        Set oCommand = Nothing
        Set oConexion = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function CajaTiposComprobanteFarmacia() As Recordset

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
          .CommandText = "CajaTiposComprobanteFarmacia"
          Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(226)): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, lcBuscaParametro.SeleccionaFilaParametro(225)): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaTiposComprobanteFarmacia = oRecordset
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 25/09/2013 - se cambio a store procedure
Function FactOrdenServicioXidOrden(lnIdOrden As Long, oConexion As Connection) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FactOrdenServicioXidOrden"
          Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set FactOrdenServicioXidOrden = oRecordset
        oRecordset.Close
        Set oRecordset = Nothing
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function CajaComprobantesPagoXnroSerieYDocumento(lcNroSerie As String, lcNroDocumento As String, oConexion As Connection) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "CajaComprobantesPagoXnroSerieYDocumento"
          Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 50, lcNroSerie): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 50, lcNroDocumento): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set CajaComprobantesPagoXnroSerieYDocumento = oRecordset
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function FacturacionServicioFinanciamientosExoneracionesEnBoleta(lnIdComprobantePago As Long, oConexion As Connection) As Recordset

        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FacturacionServicioFinanciamientosExoneracionesEnBoleta"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set FacturacionServicioFinanciamientosExoneracionesEnBoleta = oRecordset
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function FacturacionBienesFinanciamientosExoneracionesEnBoleta(lnIdComprobantePago As Long, oConexion As Connection) As Recordset

        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FacturacionBienesFinanciamientosExoneracionesEnBoleta"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set FacturacionBienesFinanciamientosExoneracionesEnBoleta = oRecordset
        Set oCommand = Nothing
End Function

'MODIFICADO POR FRANKLIN CACHAY 24/09/2013 - se cambio a store procedure
Function CajaCajaSeleccionarPorNombrePC(lcNombrePc As String) As Recordset
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
          .CommandText = "CajaCajaSeleccionarPorNombrePC"
          Set oParameter = .CreateParameter("@LoginPc", adVarChar, adParamInput, 50, lcNombrePc): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaCajaSeleccionarPorNombrePC = oRecordset
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function

Function CajaComprobantesPagoChequeaSiYaPago(lnIdCuentaAtencion As Long, lnIdServicio As Long, oConexion As Connection) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "CajaComprobantesPagoChequeaSiYaPago"
          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@lnIdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaComprobantesPagoChequeaSiYaPago = oRecordset
        Set oCommand = Nothing
End Function




Function CajaGestionActualizaFechaCierre(lnIdGestionCaja As Long, ldFechaCierre As Date) As Boolean

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
                .CommandText = "CajaGestionActualizaFechaCierre"
                Set oParameter = .CreateParameter("@fechaCierre", adDBTimeStamp, adParamInput, 0, ldFechaCierre): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdGestionCaja", adInteger, adParamInput, 0, lnIdGestionCaja): .Parameters.Append oParameter
                 .Execute
          End With
          oConexion.Close
          
          Set oRecordset = Nothing
          Set oConexion = Nothing
          Set oCommand = Nothing
          CajaGestionActualizaFechaCierre = True
          Exit Function
          
ManejadorDeError:
          ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
          Exit Function
End Function


Function CajaGestionActualizaFechaCierreXidCaja(lnIdCaja As Long, ldFechaCierre As Date) As Boolean

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
                .CommandText = "CajaGestionActualizaFechaCierreXidCaja"
                Set oParameter = .CreateParameter("@fechaCierre", adDBTimeStamp, adParamInput, 0, ldFechaCierre): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
                 .Execute
          End With
          oConexion.Close
          
          Set oRecordset = Nothing
          Set oConexion = Nothing
          Set oCommand = Nothing
          CajaGestionActualizaFechaCierreXidCaja = True
          Exit Function
          
ManejadorDeError:
          ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
          Exit Function
End Function


'debb-16/05/2016
Sub ReporteTipoTarifa(ByRef mrs_Tmp As Recordset, lbEsReporte As sghTipoProceso, _
                     lnIdCajero As Long, lnIdTipoTarifaDetalle As Long, _
                     mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lbConProrrateoEnExoneraciones As Boolean, lbConsiderarTodosItems As Boolean)


              Dim lnExoneracionCamaSP As Double, lnExoneracionSOPSP As Double, lnExoneracionRestoSP As Double, lnExoneracionItemSP As Double, lnExoneracionConsultaSP As Double
              Dim lcTexto3 As String, lcTexto2 As String, LcTexto1 As String
              Dim lRecordCount As Long, lnLineas As Long, lnIdTipoServicio As Long, lnIdCentroCosto As Long
              Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double, lnImpTotal As Double
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
              Dim rsReporte As New Recordset
              Dim oRsExoneradoBoleta As New Recordset
              Dim oRsItemsXCuenta As New Recordset
              Dim rsParamet As New Recordset
              Dim rsTmp1 As New Recordset
              Dim rsTmp2 As New Recordset
              Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
              Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
              Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
              Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
              Dim rsTmp999 As New Recordset
              
              Dim lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long
              Dim lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long
              Dim lnImptotalDetall As Double, lnCanTotalDetall As Long
              Dim lnIdProductoDetall As Long, lnCodigoDetall As String, lnDescripcionDetall As String, lnPrecioDetall As Double
              Dim lbHallo As Boolean, lcPie As String
              Dim lcBuscaParametro As New SIGHDatos.Parametros
              Dim lnBoletaExo As Double, lnBoletaTot As Double, lnBoletaAnu As Double
              Dim lnBoletaExo1 As Double, lnBoletaTot1 As Double, lnBoletaAnu1 As Double, lcBoleta As String
              Dim lnIdPartida100 As Long, lnIdProducto100 As Long, lcCodigo100 As String, lcProducto100 As String
              Dim lnPrecioVenta100 As Double, lnCantidadPagar100 As Long, lnTotalPagar100 As Double, lbSigueOK As Boolean
              Dim moConexion As New Connection, lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
              Dim lnIdTipoTarifaFarmacia As Long, lnPrecioBoleta As Double
              Dim lnIdEstadoComprobante As Long
              '
10            On Error GoTo ErrRXCC
              '
20            moConexion.CommandTimeout = 300
30            moConexion.Open sighentidades.CadenaConexion
40            moConexion.CursorLocation = adUseClient
              '
50            lnIdServiciosHospitalarios = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
60            lnIdReembolsosServicios = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
              '
70            lcTexto3 = "..comienza a generar Temporal..."
80            If ml_IdPartidaFiltro = 0 Then
90                With mrs_Tmp
100                  .Fields.Append "idPartida", adInteger
110                  .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
120                  .Fields.Append "Descripcion", adVarChar, 150, adFldIsNullable
130                  .Fields.Append "ImpAnulado", adDouble
140                  .Fields.Append "ImpExonerado", adDouble
150                  .Fields.Append "ImpNormal", adDouble
160                  .Fields.Append "ImpCancelado", adDouble
170                  .Fields.Append "ImpSubTotal", adDouble
180                  .Fields.Append "cantidad", adInteger
190                  .Fields.Append "precio", adDouble
200                  .Fields.Append "total", adDouble
210                  .LockType = adLockOptimistic
220                  .Open
230               End With
240           Else



250               With mrs_Tmp
260                  .Fields.Append "Identificador", adInteger
270                  .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
280                  .Fields.Append "Descripcion", adVarChar, 300, adFldIsNullable
290                  .Fields.Append "ImpAnulado", adDouble
300                  .Fields.Append "ImpExonerado", adDouble
310                  .Fields.Append "ImpNormal", adDouble
320                  .Fields.Append "ImpCancelado", adDouble
330                  .Fields.Append "ImpSubTotal", adDouble
340                  .Fields.Append "cantidad", adInteger
350                  .Fields.Append "precio", adDouble
360                  .Fields.Append "total", adDouble
370                  .Fields.Append "idFuenteFinanciamiento", adInteger
380                  .Fields.Append "FuenteFinanciamiento", adVarChar, 100, adFldIsNullable
                     .Fields.Append "puntoCarga", adVarChar, 50
390                  .LockType = adLockOptimistic
400                  .Open
410               End With
420           End If



              '
430           lcTexto3 = "..llena Temporal..."
440           If ml_IdPartidaFiltro = 0 Then
                  'Llena Temporal con Partidas
450               Set rsTmp = mo_reglasComunes.TiposTarifaSeleccionarTodos(moConexion)
460               If rsTmp.RecordCount = 0 Then
470                  MsgBox "Llene tabla TIPO TARIFA"
480                  Exit Sub
490               End If


500               rsTmp.MoveFirst
510               Do While Not rsTmp.EOF
520                  If rsTmp.Fields!esFarmacia = True Then
530                     lnIdTipoTarifaFarmacia = rsTmp.Fields!IdTipoTarifa
540                  End If

550                  mrs_Tmp.AddNew
560                  mrs_Tmp.Fields!idPartida = rsTmp.Fields!IdTipoTarifa
570                  mrs_Tmp.Fields!Codigo = rsTmp.Fields!Codigo
580                  mrs_Tmp.Fields!Descripcion = rsTmp.Fields!tipoTarifa
590                  mrs_Tmp.Update
600                  rsTmp.MoveNext
610               Loop

620               rsTmp.Close
630           Else

640               Set rsTmp = mo_reglasComunes.TiposTarifaSoloFarmacia(moConexion)
650               If rsTmp.RecordCount > 0 Then
660                  lnIdTipoTarifaFarmacia = rsTmp.Fields!IdTipoTarifa
670               End If

680               rsTmp.Close
690           End If

700           lcTexto3 = "..comienza a jalar boletas servicios..."
              'Boletas Servicios
710           Set rsReporte = mo_ReglasFacturacion.ServicioConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
720           lRecordCount = rsReporte.RecordCount
730           If lRecordCount > 0 Then
740              Set rsTmp999 = CajaComprobantesPagoParaTipoTarifaFechas(mda_FechaInicio, mda_FechaFin, moConexion)
750              lnLineas = 0
                 '
760              rsReporte.MoveFirst
770              Do While Not rsReporte.EOF
780                 lbSigueOK = True
790                 If lnIdCajero > 0 Then
800                      If rsReporte.Fields!IdCajero <> lnIdCajero Then
810                         lbSigueOK = False
820                      End If
830                 End If



840                 If lbSigueOK = True Then
                         
850                     lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento
860   If Val(lcTexto2) = 16201663 Then
870   lnPagoCta = 0
880   End If
890                     lnPagoCta = rsReporte.Fields!Adelantos
900                     lnExoneracion = rsReporte.Fields!exoneraciones
                        lnIdEstadoComprobante = rsReporte.Fields!idEstadoComprobante
910                     If rsReporte.Fields!idEstadoComprobante = 9 Then
920                        lnAnulado = rsReporte.Fields!totalPagado
930                        lnImpTotal = 0
940                     Else
950                        lnAnulado = 0



960                        lnImpTotal = rsReporte.Fields!totalPagado
970                     End If

                        '
980                     lnIdTipoServicio = 0
990                     lbPrimeraVez = True
                        '
1000                    Set oRsExoneradoBoleta = FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
                        '
1010                    lnBoletaExo = rsReporte.Fields!exoneraciones
1020                    lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
1030                    Select Case rsReporte.Fields!idEstadoComprobante
                        Case 6   '***Devolucion
1040                         lnBoletaTot = -rsReporte.Fields!totalPagado
1050                    Case 9   '***anulado
1060                         lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
1070                    End Select

1080                    lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
                        '
1090                    Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento)
1100                          lnPrecioBoleta = rsReporte!Precio
1110                          lnLineas = lnLineas + 1
                              '
1120                          If lbPrimeraVez = True Then
1130                             lbPrimeraVez = False
                                 
                                 '*****Emergencia
1140                             lcTexto3 = "..comienza: Emergencia..."
1150                             lnConsulta = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                                 'Set rsTmp = mo_ReglasReportes.CajaComprobantesPagoParaTipoTarifa(rsReporte.Fields!NroSerie, rsReporte.Fields!NroDocumento, moConexion)
1160                             rsTmp999.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
1170                                If rsTmp999.RecordCount > 0 Then
                                       '
1180                                   lnExoneracionConsultaSP = 0: lnExoneracionRestoSP = 0
1190                                   lnExoneracionLaboratorioSP = 0: lnExoneracionImagenesSP = 0
                                       '
1200                                   lbEncontroDato = True
1210                                   rsTmp999.MoveFirst


1220                                   Do While Not rsTmp999.EOF
'1230  If lcTexto2 = 16201663 Then
'1240  lnDescripcionDetall = ""
'1250  End If
                                          '
1260                                      lnExoneracionItemSP = 0

1270                                      If oRsExoneradoBoleta.RecordCount > 0 Then
1280                                         oRsExoneradoBoleta.MoveFirst
1290                                         Do While Not oRsExoneradoBoleta.EOF
1300                                            If oRsExoneradoBoleta.Fields!IdOrden = rsTmp999.Fields!IdOrden And oRsExoneradoBoleta.Fields!idProducto = rsTmp999.Fields!idProducto Then
1310                                               lnExoneracionItemSP = lnExoneracionItemSP + oRsExoneradoBoleta.Fields!TotalFinanciado
1320                                            End If

1330                                            oRsExoneradoBoleta.MoveNext
1340                                         Loop

1350                                      ElseIf oRsExoneradoBoleta.RecordCount = 0 And lnBoletaExo > 0 And lbConProrrateoEnExoneraciones = True Then
1360                                         lnExoneracionItemSP = Round((rsTmp999.Fields!totalPorPagar * lnBoletaExo) / (lnBoletaTot + lnBoletaExo), 2)
1370                                      End If

                                          '
1380                                      lnConsulta = lnConsulta + rsTmp999.Fields!total - lnExoneracionItemSP
1390                                      lnExoneracionConsultaSP = lnExoneracionConsultaSP + lnExoneracionItemSP
                                          '
1400                                      If lnExoneracionItemSP > 0 Then
1410                                         lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
1420                                      Else
1430                                         lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
1440                                      End If
1450                                      Select Case rsReporte.Fields!idEstadoComprobante
                                          Case 6        '***Devolucion
1460                                         lnImptotalDetall = -(rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!cantidad
1470                                         lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
1480                                      Case 9        '***Anulado
1490                                         lnImptotalDetall = 0: lnCanTotalDetall = 0
1500                                         lnAnuladoImpDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnAnuladoCanDetall = rsTmp999.Fields!cantidad

                                             lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
1510                                      Case Else     '***Pagado









1520                                         lnImptotalDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!cantidad
1530                                         lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
1540                                         If lnPagoCta > 0 Then

1550                                            'lnImptotalDetall = lnImptotalDetall - lnPagoCta
1560                                            'lnPagoCta = 0

                                                If lnPagoCta > lnImptotalDetall Then
                                                    lnPagoCta = lnPagoCta - lnImptotalDetall
                                                    lnImptotalDetall = 0
                                                Else
                                                    lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                                    lnPagoCta = 0
                                                End If
1570                                         End If
1580                                      End Select


1590                                      lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo

1600                                      lnDescripcionDetall = rsTmp999.Fields!nombreProducto: lnPrecioDetall = rsTmp999.Fields!Precio
                                          '
1610                                      lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
1620                                      lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
1630                                      lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                                          '
1640                                      lnIdPartida100 = rsTmp999.Fields!IdTipoTarifa
1650                                      lnIdProducto100 = lnIdProductoDetall
1660                                      lcCodigo100 = lnCodigoDetall
1670                                      lcProducto100 = lnDescripcionDetall
1680                                      lnPrecioVenta100 = lnPrecioDetall
1690                                      lnCantidadPagar100 = rsTmp999.Fields!cantidad
1700                                      lnTotalPagar100 = rsTmp999.Fields!total
                                          '
1710                                      Select Case rsTmp999.Fields!idProducto
                                          Case lnIdServiciosHospitalarios And Not IsNull(rsReporte.Fields!idCuentaAtencion)
                                               '******Servicios Hospitalarios, habrá que desagregarlos en 2 Partes: Farmacia y el resto en ServiciosHospitalarios
1720                                           mo_ReglasReportes.DesagregaCuentasItemsXServHosp oRsItemsXCuenta, rsReporte.Fields!idCuentaAtencion, _
                                                                                                rsTmp999.Fields!total, lnIdServiciosHospitalarios, moConexion
1730                                           If oRsItemsXCuenta.RecordCount > 0 Then
1740                                              oRsItemsXCuenta.MoveFirst

1750                                              Do While Not oRsItemsXCuenta.EOF
1760                                                  lnExoneracionImpDetall = 0
1770                                                  lnExoneracionCanDetall = 0


1780                                                  Select Case rsReporte.Fields!idEstadoComprobante
                                                      Case 6        '***Devolucion
1790                                                  Case 9        '***Anulado
1800                                                     lnImptotalDetall = 0
1810                                                     lnCanTotalDetall = 0



1820                                                     lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
1830                                                     lnAnuladoCanDetall = oRsItemsXCuenta.Fields!cantidad
1840                                                  Case Else     '***Pagado

1850                                                     lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
1860                                                     lnCanTotalDetall = oRsItemsXCuenta.Fields!cantidad
1870                                                     lnAnuladoImpDetall = 0
1880                                                     lnAnuladoCanDetall = 0
1890                                                  End Select
1900                                                  If ml_IdPartidaFiltro > 0 Then




1910                                                      lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
1920                                                      lnCodigoDetall = oRsItemsXCuenta.Fields!Codigo
1930                                                      lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
1940                                                      lnPrecioDetall = oRsItemsXCuenta.Fields!Precio
1950                                                      GrabaDetalleEnTmp ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lbConsiderarTodosItems
1960                                                  Else

1970      If oRsItemsXCuenta.Fields!idPartida = 16 Then
1980         lcTexto3 = ""
1990      End If
2000                                                      mrs_Tmp.MoveFirst


2010                                                      mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!idPartida
2020                                                      If Not mrs_Tmp.EOF Then

2030                                                         mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
2040                                                         mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
2050                                                         mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
2060                                                         mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
2070                                                         mrs_Tmp.Update
2080                                                      End If
2090                                                  End If
2100                                                  oRsItemsXCuenta.MoveNext
2110                                              Loop
2120                                           End If






2130                                      Case lnIdReembolsosServicios
                                               '******Reembolsos de Servicios, habrá que desagregarlos por Cuenta->Procedimientos
2140                                           mo_ReglasReportes.DesagregaCuentasItemsXreembolso oRsItemsXCuenta, rsReporte.Fields!IdComprobantePago, _
                                                                                                rsTmp999.Fields!total, moConexion
2150                                           If oRsItemsXCuenta.RecordCount > 0 Then
2160                                              oRsItemsXCuenta.MoveFirst

2170                                              Do While Not oRsItemsXCuenta.EOF
2180                                                  lnExoneracionImpDetall = 0
2190                                                  lnExoneracionCanDetall = 0


2200                                                  Select Case rsReporte.Fields!idEstadoComprobante
                                                      Case 6        '***Devolucion
2210                                                  Case 9        '***Anulado
2220                                                     lnImptotalDetall = 0
2230                                                     lnCanTotalDetall = 0



2240                                                     lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
2250                                                     lnAnuladoCanDetall = oRsItemsXCuenta.Fields!cantidad
2260                                                  Case Else     '***Pagado

2270                                                     lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
2280                                                     lnCanTotalDetall = oRsItemsXCuenta.Fields!cantidad
2290                                                     lnAnuladoImpDetall = 0
2300                                                     lnAnuladoCanDetall = 0
2310                                                  End Select
2320                                                  If ml_IdPartidaFiltro > 0 Then




2330                                                      lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
2340                                                      lnCodigoDetall = oRsItemsXCuenta.Fields!Codigo
2350                                                      lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
2360                                                      lnPrecioDetall = oRsItemsXCuenta.Fields!Precio
2370                                                      GrabaDetalleEnTmp ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lbConsiderarTodosItems
2380                                                  Else

2390      If oRsItemsXCuenta.Fields!idPartida = 16 Then
2400         lcTexto3 = ""
2410      End If
2420                                                      mrs_Tmp.MoveFirst


2430                                                      mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!idPartida
2440                                                      If Not mrs_Tmp.EOF Then

2450                                                         mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
2460                                                         mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
2470                                                         mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
2480                                                         mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
2490                                                         mrs_Tmp.Update
2500                                                      End If
2510                                                  End If
2520                                                  oRsItemsXCuenta.MoveNext
2530                                              Loop
2540                                           End If
2550                                      Case Else







2560                                           If ml_IdPartidaFiltro > 0 Then
2570                                              GrabaDetalleEnTmp ml_IdPartidaFiltro, rsTmp999.Fields!IdTipoTarifa, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!cantidad, rsTmp999.Fields!total, lbConsiderarTodosItems
2580                                           Else

2590      If rsTmp999.Fields!IdTipoTarifa = 16 Then
2600         lcTexto3 = ""
2610      End If
2620                                              mrs_Tmp.MoveFirst


2630                                              mrs_Tmp.Find "idPartida=" & rsTmp999.Fields!IdTipoTarifa
2640                                              If Not mrs_Tmp.EOF Then

2650                                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
2660                                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
2670                                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
2680                                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
2690                                                mrs_Tmp.Update
2700                                              End If
2710                                           End If
2720                                      End Select
2730                                      rsTmp999.MoveNext
2740                                   Loop
2750                                End If







                                    'rsTmp.Close
2760                          End If

2770                          rsReporte.MoveNext
2780                          If rsReporte.EOF Then
2790                             Exit Do
2800                          End If
2810                    Loop



                        '
2820                    oRsExoneradoBoleta.Close
                        If lnIdEstadoComprobante = 9 Then
                            lnBoletaExo1 = lnBoletaExo
                        End If

                        'Chequea Redondeo, para que cuadre con Totales de la  Boleta
2830                    If lnBoletaExo <> lnBoletaExo1 Or _
                           lnBoletaTot <> lnBoletaTot1 Or _
                           lnBoletaAnu <> lnBoletaAnu1 Then
2840                          lnAnuladoImpDetall = 0
2850                          lnExoneracionImpDetall = 0
2860                          lnImptotalDetall = 0
2870
                              If lnBoletaExo <> lnBoletaExo1 Then
2880                                lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
2890                          End If

                              If lnBoletaTot <> lnBoletaTot1 Then
2910                                lnImptotalDetall = lnBoletaTot - lnBoletaTot1
2920                          End If

                              If lnBoletaAnu <> lnBoletaAnu1 Then
2930
2940                                lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
2950                          End If

2960                          If ml_IdPartidaFiltro > 0 Then
2970                             GrabaDetalleEnTmp ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lbConsiderarTodosItems
2980                          Else

2990      If lnIdPartida100 = 16 Then
3000         lcTexto3 = ""
3010      End If

                                 
3020                             mrs_Tmp.MoveFirst
3030                             mrs_Tmp.Find "idPartida=" & lnIdPartida100
3040                             If Not mrs_Tmp.EOF Then
3050                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
3060                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
3070                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
3080                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
3090                                mrs_Tmp.Update
3100                             End If
3110                          End If
3120                    End If




                        '
3130                 Else

3140                    rsReporte.MoveNext
3150                 End If
3160             Loop
3170          End If



              'Farmacia
3180          lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
3190          If lbConsiderarTodosItems = False Then
3200              Set rsReporte = Nothing
3210              Set rsReporte = mo_ReglasFacturacion.FarmaciaConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
3220              lRecordCount = rsReporte.RecordCount
3230              If lRecordCount > 0 Then
3240                 Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(moConexion)
3250                 lnLineas = 0

                     '
3260                 rsReporte.MoveFirst
3270                 lnExoneracion = 0: lnAnulado = 0: lnImpTotal = 0
3280                 Do While Not rsReporte.EOF
3290                    lbSigueOK = True
3300                    If lnIdCajero > 0 Then
3310                       If rsReporte.Fields!IdCajero <> lnIdCajero Then
3320                           lbSigueOK = False
3330                       End If
3340                    End If



3350                    If lbSigueOK = True Then
3360                            lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento
3370                            lnPagoCta = rsReporte.Fields!Adelantos
3380                            lnExoneracion = lnExoneracion + rsReporte.Fields!exoneraciones
                                lnIdEstadoComprobante = rsReporte.Fields!idEstadoComprobante
3390                            If rsReporte.Fields!idEstadoComprobante = 9 Then
3400                               lnAnulado = lnAnulado + rsReporte.Fields!totalPagado
3410                            Else

3420                               lnImpTotal = lnImpTotal + rsReporte.Fields!totalPagado
3430                            End If

                                'Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
3440                            oRsExoneradoBoleta.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
                                '
3450                            lnBoletaExo = rsReporte.Fields!exoneraciones
3460                            lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
3470                            If rsReporte.Fields!idEstadoComprobante = 9 Then
3480                               lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
3490                            End If

3500                            lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
                                '
3510                            Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento)
3520                                  lnIdPartida100 = lnIdTipoTarifaFarmacia
3530                                  lnIdProducto100 = rsReporte.Fields!idProducto
3540                                  lcCodigo100 = rsReporte.Fields!Codigo
3550                                  lcProducto100 = rsReporte.Fields!Producto
3560                                  lnPrecioVenta100 = rsReporte.Fields!PrecioVenta
3570                                  lnCantidadPagar100 = rsReporte.Fields!CantidadPagar
3580                                  lnTotalPagar100 = rsReporte.Fields!TotalPagar
3590                                  lnLineas = lnLineas + 1

3600                                  lnExoneracionImpDetall = 0
3610                                  lnExoneracionCanDetall = 0
3620                                  If oRsExoneradoBoleta.RecordCount > 0 Then
3630                                     oRsExoneradoBoleta.MoveFirst
3640                                     Do While Not oRsExoneradoBoleta.EOF
3650                                        If oRsExoneradoBoleta.Fields!MovTipo = rsReporte.Fields!MovTipo And oRsExoneradoBoleta.Fields!movNumero = rsReporte.Fields!movNumero And oRsExoneradoBoleta.Fields!idProducto = rsReporte.Fields!idProducto Then
3660                                           lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!TotalFinanciado
3670                                           lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
3680                                        End If

3690                                        oRsExoneradoBoleta.MoveNext
3700                                     Loop
3710                                  End If


                                      '
3720                                  lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
3730                                  lnImptotalDetall = 0: lnCanTotalDetall = 0
                                      
3740                                  Select Case rsReporte.Fields!idEstadoComprobante
                                      Case 9      '***anulado
3750                                      lnAnuladoImpDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
3760                                      lnAnuladoCanDetall = rsReporte.Fields!CantidadPagar
                                          lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
3770                                  Case Else   '***pagado

3780                                      lnImptotalDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
3790                                      lnCanTotalDetall = rsReporte.Fields!CantidadPagar
3800                                      If lnPagoCta > 0 Then

3810                                         'lnImptotalDetall = lnImptotalDetall - lnPagoCta
3820                                         'lnPagoCta = 0

                                             If lnPagoCta > lnImptotalDetall Then
                                                lnPagoCta = lnPagoCta - lnImptotalDetall
                                                lnImptotalDetall = 0
                                             Else
                                                lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                                lnPagoCta = 0
                                             End If

3830                                      End If
3840                                  End Select


                                      '
3850                                  lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
3860                                  lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
3870                                  lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                                      '
3880                                  If ml_IdPartidaFiltro > 0 Then
3890                                     If lbConsiderarTodosItems = True Then
3900                                          lnIdProducto100 = 0
3910                                          lcCodigo100 = "00000"
3920                                          lcProducto100 = "FARMACIA"
3930                                     End If




3940                                     GrabaDetalleEnTmp ml_IdPartidaFiltro, lnIdTipoTarifaFarmacia, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, rsReporte.Fields!PrecioVenta, rsReporte.Fields!CantidadPagar, rsReporte.Fields!TotalPagar, lbConsiderarTodosItems
3950                                  Else
3960                                     mrs_Tmp.MoveFirst


3970                                     mrs_Tmp.Find "idPartida=" & lnIdTipoTarifaFarmacia
3980                                     If Not mrs_Tmp.EOF Then

3990                                        mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
4000                                        mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
4010                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
4020                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
4030                                        mrs_Tmp.Update
4040                                     End If
4050                                  End If



                                      
                                      '
4060                                  rsReporte.MoveNext
4070                                  If rsReporte.EOF Then
4080                                     Exit Do
4090                                  End If
4100                            Loop
                                If lnIdEstadoComprobante = 9 Then
                                    lnBoletaExo1 = lnBoletaExo
                                 End If






                                'oRsExoneradoBoleta.Close
                                'Chequea Redondeo, para que cuadre con Totales de la  Boleta
                                If lnBoletaExo <> lnBoletaExo1 Or lnBoletaTot <> lnBoletaTot1 Or lnBoletaAnu <> lnBoletaAnu1 Then
4120                                  lnAnuladoImpDetall = 0




4130                                  lnExoneracionImpDetall = 0
4140                                  lnImptotalDetall = 0

                                      If lnBoletaExo <> lnBoletaExo1 Then
4160                                        lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
4170                                  End If

                                      If lnBoletaTot <> lnBoletaTot1 Then
4190                                        lnImptotalDetall = lnBoletaTot - lnBoletaTot1
4200                                  End If

                                      If lnBoletaAnu <> lnBoletaAnu1 Then
4220                                        lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
4230                                  End If

4240                                  If ml_IdPartidaFiltro > 0 Then
4250                                     GrabaDetalleEnTmp ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lbConsiderarTodosItems
4260                                  Else
4270                                     mrs_Tmp.MoveFirst


4280                                     mrs_Tmp.Find "idPartida=" & lnIdPartida100
4290                                     If Not mrs_Tmp.EOF Then

4300                                        mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
4310                                        mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
4320                                        mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
4330                                        mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
4340                                        mrs_Tmp.Update
4350                                     End If
4360                                  End If
4370                            End If




                                '
4380                    Else

4390                        rsReporte.MoveNext
4400                    End If
4410                 Loop
4420              End If
4430          End If




              '
4440          If mrs_Tmp.RecordCount > 0 Then
4450             mrs_Tmp.MoveFirst
4460             Do While Not mrs_Tmp.EOF
4470                mrs_Tmp!ImpSubTotal = mrs_Tmp!ImpNormal - (mrs_Tmp!impAnulado + mrs_Tmp!impExonerado)
4480                If ml_IdPartidaFiltro = 0 Then
4490                Else

4500                  If mrs_Tmp!identificador > 0 Then
4510                     mrs_Tmp!total = Round(mrs_Tmp!cantidad * mrs_Tmp!Precio, 2)
4520                  Else

4530                     mrs_Tmp!total = 0
4540                     mrs_Tmp!Precio = 0
4550                  End If
4560                End If


4570                If lbConsiderarTodosItems = True Then
4580                      mrs_Tmp!idFuenteFinanciamiento = 1
4590                      mrs_Tmp!FuenteFinanciamiento = "Particular"
4600                      mrs_Tmp.Update
4610                End If

4620                mrs_Tmp.MoveNext
4630             Loop

4640             If lbConsiderarTodosItems = True Then
4650                mrs_Tmp.Sort = "DESCRIPCION,fuenteFinanciamiento"
4660             End If

4670             mrs_Tmp.MoveFirst
4680          End If

              '***********************notas de Credito ***********   'kike 2017
              If ml_IdPartidaFiltro = 0 Then
                    Set rsTmp = mo_ReglasCaja.NotaCreditoDevueltosPorNumYFecha("", "", mda_FechaInicio, mda_FechaFin)
                    LcTexto1 = ""
                    If lnIdCajero > 0 Then
                       LcTexto1 = "idCajero=" & lnIdCajero
                    End If
                    If LcTexto1 <> "" Then
                       rsTmp.Filter = LcTexto1
                    End If
                    If rsTmp.RecordCount > 0 Then
                        lnTotalPagar100 = 0
                        rsTmp.MoveFirst
                        Do While Not rsTmp.EOF
                           lnTotalPagar100 = lnTotalPagar100 + rsTmp!total
                           rsTmp.MoveNext
                        Loop
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!idPartida = 12345
                        mrs_Tmp.Fields!Codigo = "00000"
                        mrs_Tmp.Fields!Descripcion = "NOTAS DE CREDITO"
                        mrs_Tmp.Fields!impAnulado = 0
                        mrs_Tmp.Fields!impExonerado = 0
                        mrs_Tmp.Fields!ImpNormal = -lnTotalPagar100
                        mrs_Tmp.Fields!ImpCancelado = -lnTotalPagar100
                        mrs_Tmp.Update
                    End If
                    rsTmp.Close
              End If


              
4690          Set rsTmp = Nothing
4700          Set rsReporte = Nothing
4710          Set oRsExoneradoBoleta = Nothing
4720          Set oRsItemsXCuenta = Nothing
4730          Set rsParamet = Nothing
4740          Set rsTmp1 = Nothing
4750          Set rsTmp2 = Nothing
4760          Set mo_reglasComunes = Nothing
4770          Set mo_ReglasFacturacion = Nothing
4780          Set mo_ReglasReportes = Nothing
4790          Set rsTmp999 = Nothing
              Set mo_ReglasCaja = Nothing
              
4800          Exit Sub

ErrRXCC:
              If Err.Number = 94 Then
                 MsgBox Err.Description & Chr(13) & Chr(13) & "Boleta N° " & lcTexto2 & "     facturacionServiciosPagos.idProducto: " & Trim(Str(lnIdProductoDetall))
              Else
                     MsgBox Err.Description & Chr(13) & Chr(13) & "Boleta N° " & lcTexto2 & _
                                                         sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                                   "Sub ReporteTipoTarifa", "ReglasCaja.cls")  'debb-02/05/2016
              End If
                                                                          
4820          Exit Sub
4830          Resume



End Sub

'debb-21/03/2015
Sub GrabaDetalleEnTmp(mlIdPartidaFiltro As Long, lnIdPartida As Long, mrs_Tmp As Recordset, _
                      lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long, _
                      lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long, lnImptotalDetall As Double, _
                      lnCanTotalDetall As Long, lnIdProducto1 As Long, lcCodigo1 As String, _
                      lcDescripcion1 As String, lnPrecio1 As Double, lnCantidad1 As Long, _
                      lnTotal1 As Double, lbConsiderarTodosItems As Boolean)
        Dim lbPrimeraVez As Boolean, lbContinuar As Boolean
        lbContinuar = False
        If lbConsiderarTodosItems = True Then
           lbContinuar = True
        ElseIf mlIdPartidaFiltro = lnIdPartida Then
           lbContinuar = True
        End If
        If lbContinuar = True Then
        'If mlIdPartidaFiltro = lnIdPartida Then
            lbPrimeraVez = True
            If mrs_Tmp.RecordCount > 0 Then
               mrs_Tmp.MoveFirst
               Do While Not mrs_Tmp.EOF
                  If mrs_Tmp.Fields!identificador = lnIdProducto1 Then
                     lbPrimeraVez = False
                     Exit Do
                  End If
                  mrs_Tmp.MoveNext
               Loop
            End If
            If lbPrimeraVez = True Then
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!identificador = lnIdProducto1
                  mrs_Tmp.Fields!Descripcion = lcDescripcion1
                  mrs_Tmp.Fields!Codigo = lcCodigo1
                  mrs_Tmp.Fields!Precio = lnPrecio1
            End If
            mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
            mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
            mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
            mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
            mrs_Tmp.Fields!cantidad = mrs_Tmp.Fields!cantidad + lnCantidad1
            mrs_Tmp.Update
        End If
End Sub


'debb-02/05/2016
Sub ReportePartidaREsumen(ByRef mrs_Tmp As Recordset, mda_FechaInicio As Date, mda_FechaFin As Date, _
                          ml_IdPartidaFiltro As Long, lnTipoProceso As sghTipoProceso, _
                          moConexion As Connection, lbSeCreaTemporal As Boolean)
              Dim lnExoneracionCamaSP As Double, lnExoneracionSOPSP As Double, lnExoneracionRestoSP As Double, lnExoneracionItemSP As Double, lnExoneracionConsultaSP As Double
              Dim lcTexto3 As String, lcTexto2 As String, LcTexto1 As String
              Dim lRecordCount As Long, lnLineas As Long, lnIdTipoServicio As Long, lnIdCentroCosto As Long
              Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double, lnImpTotal As Double
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
              Dim rsReporte As New Recordset
              Dim oRsExoneradoBoleta As New Recordset
              Dim oRsItemsXCuenta As New Recordset
              Dim rsParamet As New Recordset
              Dim rsTmp1 As New Recordset
              Dim rsTmp2 As New Recordset
              Dim lcBuscaParametro As New SIGHDatos.Parametros
              Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
              Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
              Dim rsTmp999 As New Recordset
              
              Dim lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long
              Dim lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long
              Dim lnImptotalDetall As Double, lnCanTotalDetall As Long
              Dim lnIdProductoDetall As Long, lnCodigoDetall As String, lnDescripcionDetall As String, lnPrecioDetall As Double
              Dim lbHallo As Boolean, lcPie As String
              Dim lnBoletaExo As Double, lnBoletaTot As Double, lnBoletaAnu As Double
              Dim lnBoletaExo1 As Double, lnBoletaTot1 As Double, lnBoletaAnu1 As Double
              Dim lnIdPartida100 As Long, lnIdProducto100 As Long, lcCodigo100 As String, lcProducto100 As String
              Dim lnPrecioVenta100 As Double, lnCantidadPagar100 As Long, lnTotalPagar100 As Double
              Dim lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
              '
              On Error GoTo ErrRXCC
              '
              '
              lnIdServiciosHospitalarios = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
              lnIdReembolsosServicios = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
              '
              lcTexto3 = "..comienza a generar Temporal..."
              If lbSeCreaTemporal = True Then
                  If ml_IdPartidaFiltro = 0 Then
                      With mrs_Tmp
                         .Fields.Append "idPartida", adInteger
                         .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                         .Fields.Append "Descripcion", adVarChar, 150, adFldIsNullable
                         .Fields.Append "ImpAnulado", adDouble
                         .Fields.Append "ImpExonerado", adDouble
                         .Fields.Append "ImpNormal", adDouble
                         .Fields.Append "ImpCancelado", adDouble
                         .Fields.Append "ImpSubTotal", adDouble
                         .LockType = adLockOptimistic
                         .Open
                      End With
                  Else
                      With mrs_Tmp
                         .Fields.Append "Identificador", adInteger
                         .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
                         .Fields.Append "Descripcion", adVarChar, 300, adFldIsNullable
                         .Fields.Append "ImpAnulado", adDouble
                         .Fields.Append "ImpExonerado", adDouble
                         .Fields.Append "ImpNormal", adDouble
                         .Fields.Append "ImpCancelado", adDouble
                         .Fields.Append "idPartida", adInteger
                         .Fields.Append "ImpSubTotal", adDouble
                         .LockType = adLockOptimistic
                         .Open
                      End With
                  End If
              End If
              '
              lcTexto3 = "..llena Temporal..."
              If ml_IdPartidaFiltro = 0 And mrs_Tmp.RecordCount = 0 Then
                  'Llena Temporal con Partidas
                  Set rsTmp = mo_reglasComunes.FactPartidasPresupuestalesSeleccionarTodas(moConexion)
                  If rsTmp.RecordCount = 0 Then
                     MsgBox "Llene tabla PARTIDAS"
                     Exit Sub
                  End If
                  rsTmp.MoveFirst
                  Do While Not rsTmp.EOF
                     mrs_Tmp.AddNew
                     mrs_Tmp.Fields!idPartida = rsTmp.Fields!IdPartidaPresupuestal
                     mrs_Tmp.Fields!Codigo = rsTmp.Fields!Codigo
                     mrs_Tmp.Fields!Descripcion = rsTmp.Fields!Descripcion
                     mrs_Tmp.Update
                     rsTmp.MoveNext
                  Loop
                  rsTmp.Close
              End If
              lcTexto3 = "..comienza a jalar boletas servicios..."
              'Boletas Servicios
              
              Set rsReporte = mo_ReglasFacturacion.ServicioConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
'rsReporte.Filter = "nrodocumento='690051' and Codigo='D7242'"
              lRecordCount = rsReporte.RecordCount
              If lRecordCount > 0 Then
                 Set rsTmp999 = CajaComprobantesPagoXnroSerieYDocumentoFechas(mda_FechaInicio, mda_FechaFin, moConexion)
                 lnLineas = 0
                 '
                 rsReporte.MoveFirst
                 Do While Not rsReporte.EOF

                    lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento
                    lnPagoCta = rsReporte.Fields!Adelantos
                    lnExoneracion = rsReporte.Fields!exoneraciones
                    If rsReporte.Fields!idEstadoComprobante = 9 Then
                       lnAnulado = rsReporte.Fields!totalPagado
                       lnImpTotal = 0
                    Else
                       lnAnulado = 0
                       lnImpTotal = rsReporte.Fields!totalPagado
                    End If
                    '
                    lnIdTipoServicio = 0
                    lbPrimeraVez = True
                    '
                    Set oRsExoneradoBoleta = FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
                    '
                    lnBoletaExo = rsReporte.Fields!exoneraciones
                    lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
                    Select Case rsReporte.Fields!idEstadoComprobante
                    Case 6   '***Devolucion
                         lnBoletaTot = -rsReporte.Fields!totalPagado
                    Case 9   '***anulado
                         lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
                    End Select
                    lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
                    '
                    Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento)
      If Val(rsReporte.Fields!NroSerie) = 999 And Val(rsReporte.Fields!NroDocumento) = 345123 Then
      LcTexto1 = ""
      End If
                          lnLineas = lnLineas + 1
                          '
                          If lbPrimeraVez = True Then
                             lbPrimeraVez = False
                             
                             '*****Emergencia
                             lcTexto3 = "..comienza: Emergencia..."
                             lnConsulta = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                             'Set rsTmp = CajaComprobantesPagoXnroSerieYDocumento(rsReporte.Fields!NroSerie, rsReporte.Fields!NroDocumento, moConexion)
                             rsTmp999.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
                             
                             
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
                                      lnConsulta = lnConsulta + rsTmp999.Fields!total - lnExoneracionItemSP
                                      lnExoneracionConsultaSP = lnExoneracionConsultaSP + lnExoneracionItemSP
                                      '
                                      If lnExoneracionItemSP > 0 Then
                                         lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                      Else
                                         lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                      End If
                                      Select Case rsReporte.Fields!idEstadoComprobante
                                      Case 6        '***Devolucion
                                         lnImptotalDetall = -(rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!cantidad
                                         lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                      Case 9        '***Anulado
                                         lnImptotalDetall = 0: lnCanTotalDetall = 0
                                         lnAnuladoImpDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnAnuladoCanDetall = rsTmp999.Fields!cantidad
                                      Case Else     '***Pagado
                                         lnImptotalDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!cantidad
                                         lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                         If lnPagoCta > 0 Then

                                            'lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                            'lnPagoCta = 0
                                            If lnPagoCta > lnImptotalDetall Then
                                                lnPagoCta = lnPagoCta - lnImptotalDetall
                                                lnImptotalDetall = 0
                                            Else
                                                lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                                lnPagoCta = 0
                                            End If
                                            
                                         End If
                                      End Select
                                      lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!Codigo
                                      lnDescripcionDetall = rsTmp999.Fields!nombreProducto: lnPrecioDetall = rsTmp999.Fields!Precio
                                      '
                                      lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                                      lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                                      lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                                      '
                                      lnIdPartida100 = rsTmp999.Fields!idPartida
                                      lnIdProducto100 = lnIdProductoDetall
                                      lcCodigo100 = lnCodigoDetall
                                      lcProducto100 = lnDescripcionDetall
                                      lnPrecioVenta100 = lnPrecioDetall
                                      lnCantidadPagar100 = rsTmp999.Fields!cantidad
                                      lnTotalPagar100 = rsTmp999.Fields!total
                                      '
                                      Select Case rsTmp999.Fields!idProducto
                                      Case lnIdServiciosHospitalarios And Not IsNull(rsReporte.Fields!idCuentaAtencion)
                                           '******Servicios Hospitalarios, habrá que desagregarlos en 2 Partes: Farmacia y el resto en ServiciosHospitalarios
                                           mo_ReglasFacturacion.DesagregaCuentasItemsXServHosp oRsItemsXCuenta, _
                                                                rsReporte.Fields!idCuentaAtencion, rsTmp999.Fields!total, moConexion, _
                                                                lnIdServiciosHospitalarios
                                           If oRsItemsXCuenta.RecordCount > 0 Then
                                              oRsItemsXCuenta.MoveFirst
                                              Do While Not oRsItemsXCuenta.EOF
                                                  lnExoneracionImpDetall = 0
                                                  lnExoneracionCanDetall = 0
                                                  Select Case rsReporte.Fields!idEstadoComprobante
                                                  Case 6        '***Devolucion
                                                  Case 9        '***Anulado
                                                     lnImptotalDetall = 0
                                                     lnCanTotalDetall = 0
                                                     lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                                     lnAnuladoCanDetall = oRsItemsXCuenta.Fields!cantidad
                                                  Case Else     '***Pagado
                                                     lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                                     lnCanTotalDetall = oRsItemsXCuenta.Fields!cantidad
                                                     lnAnuladoImpDetall = 0
                                                     lnAnuladoCanDetall = 0
                                                  End Select
                                                  If ml_IdPartidaFiltro > 0 Then
                                                      lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                      lnCodigoDetall = oRsItemsXCuenta.Fields!Codigo
                                                      lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                      lnPrecioDetall = oRsItemsXCuenta.Fields!Precio
                                                      GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                                  Else
                                                      mrs_Tmp.MoveFirst
                                                      mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!idPartida
                                                      If Not mrs_Tmp.EOF Then
                                                         mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                         mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                         mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                         mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                         mrs_Tmp.Update
                                                      End If
                                                  End If
                                                  oRsItemsXCuenta.MoveNext
                                              Loop
                                           End If
                                      Case lnIdReembolsosServicios
                                           '******Reembolsos de Servicios, habrá que desagregarlos por Cuenta->Procedimientos
                                           mo_ReglasFacturacion.DesagregaCuentasItemsXreembolso oRsItemsXCuenta, rsReporte.Fields!IdComprobantePago, _
                                                                rsTmp999.Fields!total, moConexion
                                           If oRsItemsXCuenta.RecordCount > 0 Then
                                              oRsItemsXCuenta.MoveFirst
                                              Do While Not oRsItemsXCuenta.EOF
                                                  lnExoneracionImpDetall = 0
                                                  lnExoneracionCanDetall = 0
                                                  Select Case rsReporte.Fields!idEstadoComprobante
                                                  Case 6        '***Devolucion
                                                  Case 9        '***Anulado
                                                     lnImptotalDetall = 0
                                                     lnCanTotalDetall = 0
                                                     lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                                     lnAnuladoCanDetall = oRsItemsXCuenta.Fields!cantidad
                                                  Case Else     '***Pagado
                                                     lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                                     lnCanTotalDetall = oRsItemsXCuenta.Fields!cantidad
                                                     lnAnuladoImpDetall = 0
                                                     lnAnuladoCanDetall = 0
                                                  End Select
                                                  If ml_IdPartidaFiltro > 0 Then
                                                      lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                      lnCodigoDetall = oRsItemsXCuenta.Fields!Codigo
                                                      lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                      lnPrecioDetall = oRsItemsXCuenta.Fields!Precio
                                                      GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                                  Else
                                                      mrs_Tmp.MoveFirst
                                                      mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!idPartida
                                                      If Not mrs_Tmp.EOF Then
                                                         mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                         mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                         mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                         mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                         mrs_Tmp.Update
                                                      End If
                                                  End If
                                                  oRsItemsXCuenta.MoveNext
                                              Loop
                                           End If
                                      Case Else
                                           If ml_IdPartidaFiltro > 0 Then
                                              GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, rsTmp999.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!cantidad, rsTmp999.Fields!total, lnTipoProceso
                                           Else
                                              mrs_Tmp.MoveFirst
                                              mrs_Tmp.Find "idPartida=" & rsTmp999.Fields!idPartida
                                              If Not mrs_Tmp.EOF Then
                                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                mrs_Tmp.Update
                                              End If
                                           End If
                                      End Select
                                      rsTmp999.MoveNext
                                   Loop
                                End If
                               ' rsTmp.Close
                          End If
                          rsReporte.MoveNext
                          If rsReporte.EOF Then
                             Exit Do
                          End If
                    Loop
                    '
                    oRsExoneradoBoleta.Close
                    'Chequea Redondeo, para que cuadre con Totales de la  Boleta
                    If CCur(Str(lnBoletaExo)) <> CCur(Str(lnBoletaExo1)) Or _
                        CCur(Str(lnBoletaTot)) <> CCur(Str(sighentidades.DevuelveNumeroRedondeado(lnBoletaTot1))) Or _
                        CCur(Str(lnBoletaAnu)) <> CCur(Str(lnBoletaAnu1)) Then
                          lnAnuladoImpDetall = 0
                          lnExoneracionImpDetall = 0
                          lnImptotalDetall = 0
                          If CCur(Str(lnBoletaExo)) <> CCur(Str(lnBoletaExo1)) Then
                                lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                          End If
                          If CCur(Str(lnBoletaTot)) <> CCur(Str(sighentidades.DevuelveNumeroRedondeado(lnBoletaTot1))) Then
                                lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                          End If
                          If CCur(Str(lnBoletaAnu)) <> CCur(Str(lnBoletaAnu1)) Then
                                lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                          End If
                          If ml_IdPartidaFiltro > 0 Then
                             GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                          Else
                             mrs_Tmp.MoveFirst
                             mrs_Tmp.Find "idPartida=" & lnIdPartida100
                             If Not mrs_Tmp.EOF Then
                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                mrs_Tmp.Update
                             End If
                          End If
                    End If
                    '
                 Loop
              End If
              'Farmacia
              lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
              Set rsReporte = Nothing
              Set rsReporte = mo_ReglasFacturacion.FarmaciaConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
              lRecordCount = rsReporte.RecordCount
              If lRecordCount > 0 Then
                 Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(moConexion)
                 lnLineas = 0
                 '
                 rsReporte.MoveFirst
                 lnExoneracion = 0: lnAnulado = 0: lnImpTotal = 0
                 Do While Not rsReporte.EOF
                    lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento
                    lnPagoCta = rsReporte.Fields!Adelantos
                    lnExoneracion = lnExoneracion + rsReporte.Fields!exoneraciones
                    If rsReporte.Fields!idEstadoComprobante = 9 Then
                       lnAnulado = lnAnulado + rsReporte.Fields!totalPagado
                    Else
                       lnImpTotal = lnImpTotal + rsReporte.Fields!totalPagado
                    End If
                    

                    'Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
                    oRsExoneradoBoleta.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
                   '
                    lnBoletaExo = rsReporte.Fields!exoneraciones
                    lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
                    If rsReporte.Fields!idEstadoComprobante = 9 Then
                       lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
                    End If
                    lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
                    '
                    Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!NroDocumento)
                          lnIdPartida100 = rsReporte.Fields!idPartida
                          lnIdProducto100 = rsReporte.Fields!idProducto
                          lcCodigo100 = rsReporte.Fields!Codigo
                          lcProducto100 = rsReporte.Fields!Producto
                          lnPrecioVenta100 = rsReporte.Fields!PrecioVenta
                          lnCantidadPagar100 = rsReporte.Fields!CantidadPagar
                          lnTotalPagar100 = rsReporte.Fields!TotalPagar
                          lnLineas = lnLineas + 1
                          lnExoneracionImpDetall = 0
                          lnExoneracionCanDetall = 0
                          If oRsExoneradoBoleta.RecordCount > 0 Then
                             oRsExoneradoBoleta.MoveFirst
                             Do While Not oRsExoneradoBoleta.EOF
                                If oRsExoneradoBoleta.Fields!MovTipo = rsReporte.Fields!MovTipo And oRsExoneradoBoleta.Fields!movNumero = rsReporte.Fields!movNumero And oRsExoneradoBoleta.Fields!idProducto = rsReporte.Fields!idProducto Then
                                   lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!TotalFinanciado
                                   lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
                                End If
                                oRsExoneradoBoleta.MoveNext
                             Loop
                          End If
                          '
                          lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                          lnImptotalDetall = 0: lnCanTotalDetall = 0
                          
                          Select Case rsReporte.Fields!idEstadoComprobante
                          Case 9      '***anulado
                              lnAnuladoImpDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                              lnAnuladoCanDetall = rsReporte.Fields!CantidadPagar
                          Case Else   '***pagado
                              lnImptotalDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                              lnCanTotalDetall = rsReporte.Fields!CantidadPagar
                              If lnPagoCta > 0 Then

                                    'lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                    'lnPagoCta = 0
                                    If lnPagoCta > lnImptotalDetall Then
                                        lnPagoCta = lnPagoCta - lnImptotalDetall
                                        lnImptotalDetall = 0
                                    Else
                                        lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                        lnPagoCta = 0
                                    End If

                              End If
                          End Select
                          '
                          lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                          lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                          lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                          '
                          If ml_IdPartidaFiltro > 0 Then
                             GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, rsReporte.Fields!idPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, rsReporte.Fields!idProducto, rsReporte.Fields!Codigo, rsReporte.Fields!Producto, rsReporte.Fields!PrecioVenta, rsReporte.Fields!CantidadPagar, rsReporte.Fields!TotalPagar, lnTipoProceso
                          Else
                             mrs_Tmp.MoveFirst
                             mrs_Tmp.Find "idPartida=" & rsReporte.Fields!idPartida
                             If Not mrs_Tmp.EOF Then
                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                mrs_Tmp.Update
                             End If
                          End If
                          
                          '
                          rsReporte.MoveNext
                          If rsReporte.EOF Then
                             Exit Do
                          End If
                    Loop
                    'oRsExoneradoBoleta.Close
                    'Chequea Redondeo, para que cuadre con Totales de la  Boleta
                    If CCur(Str(lnBoletaExo)) <> CCur(Str(lnBoletaExo1)) Or _
                       CCur(Str(lnBoletaTot)) <> CCur(Str(sighentidades.DevuelveNumeroRedondeado(lnBoletaTot1))) Or _
                       CCur(Str(lnBoletaAnu)) <> CCur(Str(lnBoletaAnu1)) Then
                          lnAnuladoImpDetall = 0
                          lnExoneracionImpDetall = 0
                          lnImptotalDetall = 0
                          If CCur(Str(lnBoletaExo)) <> CCur(Str(lnBoletaExo1)) Then
                                lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                          End If
                          If CCur(Str(lnBoletaTot)) <> CCur(Str(sighentidades.DevuelveNumeroRedondeado(lnBoletaTot1))) Then
                                lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                          End If
                          If CCur(Str(lnBoletaAnu)) <> CCur(Str(lnBoletaAnu1)) Then
                                lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                          End If
                          If ml_IdPartidaFiltro > 0 Then
                             GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                          Else
                             mrs_Tmp.MoveFirst
                             mrs_Tmp.Find "idPartida=" & lnIdPartida100
                             If Not mrs_Tmp.EOF Then
                                mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                mrs_Tmp.Update
                             End If
                          End If
                    End If
                    '
                 Loop
              End If
              If mrs_Tmp.RecordCount > 0 Then
                 mrs_Tmp.MoveFirst
                 Do While Not mrs_Tmp.EOF
                    mrs_Tmp!ImpSubTotal = mrs_Tmp!ImpNormal - (mrs_Tmp!impAnulado + mrs_Tmp!impExonerado)
                    mrs_Tmp.Update
                    mrs_Tmp.MoveNext
                 Loop
              End If
              Set rsTmp = Nothing
              Set rsReporte = Nothing
              Set oRsExoneradoBoleta = Nothing
              Set oRsItemsXCuenta = Nothing
              Set rsParamet = Nothing
              Set rsTmp1 = Nothing
              Set rsTmp2 = Nothing
              Set lcBuscaParametro = Nothing
              'Set moConexion = Nothing
              Set mo_reglasComunes = Nothing
              Set rsTmp999 = Nothing
              Exit Sub
ErrRXCC:
              MsgBox Err.Description & Chr(13) & "Boleta: " & lcTexto2 & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                                          "Sub ReportePartidaREsumen", "ReglasCaja.cls")  'debb-02/05/2016
End Sub


Sub GrabaDetalleEnTmpPartida(mlIdPartidaFiltro As Long, lnIdPartida As Long, mrs_Tmp As Recordset, _
                             lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long, _
                             lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long, _
                             lnImptotalDetall As Double, lnCanTotalDetall As Long, lnIdProducto1 As Long, _
                             lcCodigo1 As String, lcDescripcion1 As String, lnPrecio1 As Double, _
                             lnCantidad1 As Long, lnTotal1 As Double, lnTipoProceso As sghTipoProceso)
        Dim lbPrimeraVez As Boolean
        Dim lbContinuaTMP As Boolean
        lbContinuaTMP = False
        If lnTipoProceso = sghProcesaYgraba Then
               lbContinuaTMP = True
        ElseIf lnTipoProceso = sghSoloParaReporte Then
            If mlIdPartidaFiltro = lnIdPartida Then
               lbContinuaTMP = True
            End If
        End If
        If lbContinuaTMP = True Then
            lbPrimeraVez = True
            If mrs_Tmp.RecordCount > 0 Then
               mrs_Tmp.MoveFirst
               Do While Not mrs_Tmp.EOF
                  If mrs_Tmp.Fields!identificador = lnIdProducto1 And mrs_Tmp.Fields!idPartida = lnIdPartida Then
                     lbPrimeraVez = False
                     Exit Do
                  End If
                  mrs_Tmp.MoveNext
               Loop
            End If
            If lbPrimeraVez = True Then
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!identificador = lnIdProducto1
                  mrs_Tmp.Fields!Descripcion = Left(lcDescripcion1, 150)
                  mrs_Tmp.Fields!Codigo = lcCodigo1
                  mrs_Tmp.Fields!idPartida = lnIdPartida
            End If
            mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
            mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
            mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
            mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
            mrs_Tmp.Update
        End If
End Sub




Function FactPartidasPresupuestalesXMesSelecionaUltimoProceso(mda_FechaInicio As Date, oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesXMesSelecionaUltimoProceso"
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set FactPartidasPresupuestalesXMesSelecionaUltimoProceso = oRsTmp1
    Set oCommand = Nothing
End Function

Sub FactPartidasPresupuestalesXMesEliminar(mda_FechaInicio As Date, oConexion As Connection)
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesXMesEliminar"
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
End Sub


Sub FactPartidasPresupuestalesXMesAgregar(ldFecha As Date, lnIdPartida As Long, _
                                         lnIdProducto As Long, lnImpAnulado As Double, _
                                         lnImpExonerado As Double, lnImpNormal As Double, _
                                         lnImpCancelado As Double, oConexion As Connection)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesXMesAgregar"
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, ldFecha): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPartida", adInteger, adParamInput, 0, lnIdPartida): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ImpAnulado", adCurrency, adParamInput, 0, lnImpAnulado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@impExonerado", adCurrency, adParamInput, 0, lnImpExonerado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ImpNormal", adCurrency, adParamInput, 0, lnImpNormal): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ImpCancelado", adCurrency, adParamInput, 0, lnImpCancelado): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
End Sub

Function CajaComprobantePagoUltimaBoletaDelAnio(mda_FechaInicio As Date, oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantePagoUltimaBoletaDelAnio"
        Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set CajaComprobantePagoUltimaBoletaDelAnio = oRsTmp1
    Set oCommand = Nothing
End Function

Function FactPartidasPresupuestalesXMesResumen(mda_FechaInicio As Date, mda_FechaFin As Date, lnIdPartida As Long, _
                                               lbEsFarmacia As Boolean, oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesXMesResumen"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPartida", adInteger, adParamInput, 0, lnIdPartida): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@esMedicamento", adInteger, adParamInput, 0, IIf(lbEsFarmacia = True, 1, 0)): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set FactPartidasPresupuestalesXMesResumen = oRsTmp1
    Set oCommand = Nothing
End Function
Function CajaComprobantesPagoXnroSerieYDocumentoFechas(mda_FechaInicio As Date, mda_FechaFin As Date, _
                                                       oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoXnroSerieYDocumentoFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set CajaComprobantesPagoXnroSerieYDocumentoFechas = oRsTmp1
    Set oCommand = Nothing
End Function

Function FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos"
        Set oRsTmp1 = .Execute
    End With
    Set FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos = oRsTmp1
    Set oCommand = Nothing
End Function

Function CajaComprobantesPagoParaTipoTarifaFechas(mda_FechaInicio As Date, mda_FechaFin As Date, _
                                                       oConexion As Connection) As Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoParaTipoTarifaFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, mda_FechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, mda_FechaFin): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set CajaComprobantesPagoParaTipoTarifaFechas = oRsTmp1
    Set oCommand = Nothing
End Function

Function FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo(lcCodigo As String, lcNombre As String, esCpt As Long, _
                Optional oConexion As Connection) As Recordset
    Dim oTabla As New CatalogoServicios
    Set FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo = oTabla.FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo(lcCodigo, lcNombre, esCpt, oConexion)
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
End Function

'********************************************************************************
'*****************************************INO************************************
'********************************************************************************

Function CajaDevolucionesAgregar(oDoCajaDevoluciones As DOCajaDevoluciones) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaDevoluciones As New cajadevoluciones

    CajaDevolucionesAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaDevoluciones.Conexion = oConexion
       
    If oCajaDevoluciones.Insertar(oDoCajaDevoluciones) Then
        CajaDevolucionesAgregar = True
    Else
        ms_MensajeError = oCajaDevoluciones.MensajeError
        
        End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaDevoluciones = Nothing
    
End Function


Function CajaDevolucionesModificar(oDoCajaDevoluciones As DOCajaDevoluciones) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaDevoluciones As New cajadevoluciones
'Dim total As Double
    
    CajaDevolucionesModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaDevoluciones.Conexion = oConexion
    
    'total = oCajaGestion.TotalPorGestion(oDoCajaGestion.IdGestionCaja)
    'oDoCajaGestion.TotalCobrado = total
    If oCajaDevoluciones.modificar(oDoCajaDevoluciones) Then
        CajaDevolucionesModificar = True
    Else
        ms_MensajeError = oCajaDevoluciones.MensajeError
        
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaDevoluciones = Nothing
End Function

Function CajaDevolucionesEliminar(oDoCajaDevoluciones As DOCajaDevoluciones) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCajaDevoluciones As New cajadevoluciones
    
    CajaDevolucionesEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCajaDevoluciones.Conexion = oConexion
        
    If oCajaDevoluciones.Eliminar(oDoCajaDevoluciones) Then
        CajaDevolucionesEliminar = True
    Else
        ms_MensajeError = oCajaDevoluciones.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaDevoluciones = Nothing
    
End Function


Function CajaDevolucionSeleccionarPorId(IdCajaDevolucion As Long) As DOCajaDevoluciones
Dim oConexion As New ADODB.Connection
Dim oCajaDevoluciones As New cajadevoluciones
Dim oDoCajaDevoluciones As New DOCajaDevoluciones
    
    Set CajaDevolucionSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajaDevoluciones.Conexion = oConexion
    oDoCajaDevoluciones.idDevolucion = IdCajaDevolucion
    If oCajaDevoluciones.SeleccionarPorId(oDoCajaDevoluciones) Then
        Set CajaDevolucionSeleccionarPorId = oDoCajaDevoluciones
    End If
    oConexion.Close
    ms_MensajeError = oCajaDevoluciones.MensajeError
    Set oConexion = Nothing
    Set oCajaDevoluciones = Nothing
End Function


'Function CajaBuscarDevolucion(NroSerie As String, NroDocumento As String) As Recordset
'
'Dim oRsTmp1 As New Recordset
'Dim oConexion As New ADODB.Connection
'    Dim oCommand As New ADODB.Command
'    Dim oParameter As ADODB.Parameter
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = oConexion
'        .CommandTimeout = 150
'        .CommandText = "CajaBuscarDevueltos"
'        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, NroSerie): .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 7, NroDocumento): .Parameters.Append oParameter
'        Set oRsTmp1 = .Execute
'    End With
'    Set CajaBuscarDevolucion = oRsTmp1
'    Set oCommand = Nothing
'End Function


Function BuscarTriaje(idAtencion As String) As Recordset

Dim oRsTmp1 As New Recordset
Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesCESeleccionarPorId"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set BuscarTriaje = oRsTmp1
    Set oCommand = Nothing
End Function

Function CajaDevolucionesSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaDevolucionesConsultarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaDevolucionesSeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function BuscarDevoluciones(idComprobante As Long) As Recordset

Dim oRsTmp1 As New Recordset
Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaDevolucionesBuscar"
        Set oParameter = .CreateParameter("@idComprobantePago", adVarChar, adParamInput, 10, idComprobante): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set BuscarDevoluciones = oRsTmp1
    Set oCommand = Nothing
End Function
Function BuscarDevolucionesApellido(lApellidoPaterno As String) As Recordset

Dim oRsTmp1 As New Recordset
Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaDevolucionesBuscarApellido"
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 100, lApellidoPaterno): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set BuscarDevolucionesApellido = oRsTmp1
    Set oCommand = Nothing
End Function
Function BuscarDevolucionesApellido2(lApellidoMaterno As String) As Recordset

Dim oRsTmp1 As New Recordset
Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaDevolucionesBuscarApellido2"
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 100, lApellidoMaterno): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set BuscarDevolucionesApellido2 = oRsTmp1
    Set oCommand = Nothing
End Function


'reportes
Function CajaDevolucionesPorFechas(lcFechaIni As Date, lcFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaDevolucionesPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaDevolucionesPorFechasUsuario"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, lcFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, lcFechaFin): .Parameters.Append oParameter
        'Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, lIdUsuarioSistema): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaDevolucionesPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'*****Prorrateo en un ITEM de la Boleta, si el personal de SERVICIO SOCIAL no lo hizo en ESTADO DE CUENTA ***************
Function BoletaCajaProrrateaExoneracionXitem(lnTotalBoleta As Double, lnExoneracionBoleta As Double, _
                                             lnTotalItem As Double) As Double
    BoletaCajaProrrateaExoneracionXitem = Round(lnTotalItem * lnExoneracionBoleta / (lnTotalBoleta + lnExoneracionBoleta), 2)
End Function
'*****Prorrateo en un ITEM de la Boleta, si el personal de SERVICIO SOCIAL no lo hizo en ESTADO DE CUENTA ***************
Function BoletaCajaCalculaExoneracionXitem(lnTotalBoleta As Double, lnExoneracionBoleta As Double, _
                                           lnIdProductoAprorratear, lnIdOrden As Long, lnIdCajaComprobantePago, _
                                           lnIdCuentaAtencion As Long, oConexion As Connection, lnTotalItem As Double) As Double
    BoletaCajaCalculaExoneracionXitem = BoletaCajaProrrateaExoneracionXitem(lnTotalBoleta, lnExoneracionBoleta, lnTotalItem)
    If lnIdCuentaAtencion > 0 Then
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim lnExonerado As Double
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "FacturacionServicioFinanciamientosExonerados"
             Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdCajaComprobantePago): .Parameters.Append oParameter
             Set oRecordset = .Execute
        End With
        oRecordset.Filter = "totalfinanciado>0"
        If oRecordset.RecordCount > 0 Then
           oRecordset.Filter = "idProducto=" & lnIdProductoAprorratear & " and idOrden=" & lnIdOrden
           If oRecordset.RecordCount > 0 Then
              lnExonerado = 0
              oRecordset.MoveFirst
              Do While Not oRecordset.EOF
                 lnExonerado = lnExonerado + oRecordset.Fields!TotalFinanciado
                 oRecordset.MoveNext
              Loop
              BoletaCajaCalculaExoneracionXitem = lnExonerado
           Else
              BoletaCajaCalculaExoneracionXitem = 0
           End If
        End If
        Set oCommand = Nothing
        Set oRecordset = Nothing
    End If
End Function
Function CajaComprobantesPagoSoloServicioXitem(lnIdProducto As Long, dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
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
        .CommandText = "CajaComprobantesPagoSoloServicioXitem"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoSoloServicioXitem = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Nota de Credito
Function NotaCreditoDebitoCargarEstadoNotaCredito() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoDebitoCargarEstadoNotaCredito = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoDebitoCargarEstadoNotaCredito"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoDebitoCargarEstadoNotaCredito = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function NotaCreditoDebitoTipoNotaSeleccionarPorTipo(IdTipoNota As Integer, TipoCompNota As String) As DONotaCreditoDebitoTipoNota
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebitoTipoNota As New NotaCreditoDebitoTipoNota
    Dim oDONotaCreditoDebitoTipoNota As New DONotaCreditoDebitoTipoNota
    
    Set NotaCreditoDebitoTipoNotaSeleccionarPorTipo = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oNotaCreditoDebitoTipoNota.Conexion = oConexion
    oDONotaCreditoDebitoTipoNota.IdTipoNota = IdTipoNota
    oDONotaCreditoDebitoTipoNota.TipoNota = TipoCompNota
    If oNotaCreditoDebitoTipoNota.SeleccionarPorIdTipoNota(oDONotaCreditoDebitoTipoNota) Then
        Set NotaCreditoDebitoTipoNotaSeleccionarPorTipo = oDONotaCreditoDebitoTipoNota
    Else
        Set NotaCreditoDebitoTipoNotaSeleccionarPorTipo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oNotaCreditoDebitoTipoNota.MensajeError
    Set oConexion = Nothing
    Set oNotaCreditoDebitoTipoNota = Nothing
End Function

Function NotaCreditoDebitoAgregar(oDoNotaCreditoDebito As DoNotaCreditoDebito, TipoOrden As Integer, TipoNota As Integer) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebito As New NotaCreditoDebito
    Dim oDOCita As New doCita: Dim oCita As New Citas
    Dim oDoHistoria As New DOHistoriaClinica: Dim oHistoriasSolicitadas As New HistoriasSolicitadas
    Dim oDoAtencion As New DOAtencion: Dim oAtenciones As New Atenciones
    Dim oDONotaCreditoDebitoTipoNota As New DONotaCreditoDebitoTipoNota
    Dim oNotaCreditoDebitoTipoNota As New NotaCreditoDebitoTipoNota
    Dim orsTemp As New Recordset
    Dim orsTemp2 As New Recordset
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim bProcesoOK As Boolean
    oConexion.Open sighentidades.CadenaConexion
    
    NotaCreditoDebitoAgregar = False
    oConexion.BeginTrans
    
    Set oNotaCreditoDebito.Conexion = oConexion
    Set oHistoriasSolicitadas.Conexion = oConexion
    Set oCita.Conexion = oConexion
    Set oAtenciones.Conexion = oConexion
    Set oNotaCreditoDebitoTipoNota.Conexion = oConexion
    
    bProcesoOK = True
    If TipoOrden = sghTipoPaqueteSoloServicio Then
        Set orsTemp = mo_ReglasCaja.NotaCreditoConsultarOrdenServicio(oDoNotaCreditoDebito.IdComprobantePago)
        If orsTemp.RecordCount > 0 Then
            If orsTemp.Fields!IdPuntoCarga = 6 Then 'CE
                Set orsTemp2 = mo_ReglasCaja.NotaCreditoConsultarCitaPorNCuenta(orsTemp.Fields!idCuentaAtencion)
                If orsTemp2.RecordCount > 0 Then
                    If MsgBox("La cita ligada al comprobante de pago será anulada. ¿Está seguro de continuar con la operación?", vbCritical + vbYesNo, "Nota de crédito") = vbNo Then
                        Set orsTemp = Nothing: Set orsTemp2 = Nothing
                        bProcesoOK = False
                    Else
                        If MsgBox("La historia solicitada para la cita será anulada. ¿Está seguro de continuar con la operación?", vbCritical + vbYesNo, "Nota de crédito") = vbNo Then
                            Set orsTemp = Nothing: Set orsTemp2 = Nothing
                            bProcesoOK = False
                        Else
                            oDOCita.IdCita = orsTemp2.Fields!IdCita
                            oDOCita.idAtencion = orsTemp2.Fields!idAtencion
                            oDOCita.idPaciente = orsTemp.Fields!idPaciente
                            oDOCita.fecha = orsTemp2.Fields!fecha
                            oDOCita.HoraInicio = orsTemp2.Fields!HoraInicio
                            'Eliminar solicitud de historias
                            If Not oHistoriasSolicitadas.EliminarPorCita(oDOCita) Then
                                Set orsTemp = Nothing: Set orsTemp2 = Nothing
                                bProcesoOK = False: ms_MensajeError = oHistoriasSolicitadas.MensajeError: GoTo ErrorManager
                            End If
                            'Eliminar la cita
                            If Not oCita.Eliminar(oDOCita) Then
                                Set orsTemp = Nothing: Set orsTemp2 = Nothing
                                bProcesoOK = False: ms_MensajeError = oCita.MensajeError: GoTo ErrorManager
                            End If
                            'Anular la atención
                            oDoAtencion.idAtencion = orsTemp2.Fields!idAtencion
                            If oAtenciones.SeleccionarPorId(oDoAtencion) Then
                                oDoAtencion.IdEstadoAtencion = sghEstadoTabla.sghAnulado
                                If Not oAtenciones.modificar(oDoAtencion) Then
                                    bProcesoOK = False: ms_MensajeError = oAtenciones.MensajeError: GoTo ErrorManager
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Set orsTemp = Nothing
        Set orsTemp2 = Nothing
    End If

    If bProcesoOK = True Then
        'debb-14/03/2017 (inicio)
        Dim lcNroDocumento As String
        lcNroDocumento = NotaCreditoDebitoUltimoDctoMasUNO(oDoNotaCreditoDebito.NroSerie, oConexion)
        If lcNroDocumento <> "" Then
           oDoNotaCreditoDebito.NroDocumento = lcNroDocumento
        End If
        'debb-14/03/2017 (fin)
        If Not oNotaCreditoDebito.Insertar(oDoNotaCreditoDebito) Then
             bProcesoOK = False: ms_MensajeError = oNotaCreditoDebito.MensajeError: GoTo ErrorManager
        End If
               
        oDONotaCreditoDebitoTipoNota.IdTipoNota = oDoNotaCreditoDebito.IdTipoNota
        oDONotaCreditoDebitoTipoNota.TipoNota = TipoNota
        If oNotaCreditoDebitoTipoNota.SeleccionarPorIdTipoNota(oDONotaCreditoDebitoTipoNota) Then
            oDONotaCreditoDebitoTipoNota.NroDocumento = oDoNotaCreditoDebito.NroDocumento
            If Not oNotaCreditoDebitoTipoNota.modificar(oDONotaCreditoDebitoTipoNota) Then
                bProcesoOK = False: ms_MensajeError = oNotaCreditoDebitoTipoNota.MensajeError: GoTo ErrorManager
            End If
        End If
        
        oConexion.CommitTrans
    Else
        GoTo ErrorManager
        ms_MensajeError = "No se pudo grabar la Nota de Crédito"
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = ""
    NotaCreditoDebitoAgregar = True
ErrorManager:
    If NotaCreditoDebitoAgregar = False Then
        oConexion.RollbackTrans
        oConexion.Close
        Set oConexion = Nothing
    End If
    Set oNotaCreditoDebito = Nothing
    Set oHistoriasSolicitadas = Nothing
    Set oNotaCreditoDebitoTipoNota = Nothing
    Set oCita = Nothing
End Function

Function NotaCreditoDebitoModificar(oDoNotaCreditoDebito As DoNotaCreditoDebito) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebito As New NotaCreditoDebito
    
    NotaCreditoDebitoModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oNotaCreditoDebito.Conexion = oConexion
    If oNotaCreditoDebito.modificar(oDoNotaCreditoDebito) Then
        NotaCreditoDebitoModificar = True
    Else
        ms_MensajeError = oNotaCreditoDebito.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oNotaCreditoDebito = Nothing
End Function

Function NotaCreditoDebitoEliminar(oDoNotaCreditoDebito As DoNotaCreditoDebito) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebito As New NotaCreditoDebito
    
    NotaCreditoDebitoEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oNotaCreditoDebito.Conexion = oConexion
    If oNotaCreditoDebito.Eliminar(oDoNotaCreditoDebito) Then
        NotaCreditoDebitoEliminar = True
    Else
        ms_MensajeError = oNotaCreditoDebito.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oNotaCreditoDebito = Nothing
End Function

Function NotaCreditoRegistrosTotalesPorNumYFecha(lcNserie As String, lcNdocumento As String, ldIdEstado As Integer, lcRazonSocial As String, dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoRegistrosTotalesPorNumYFecha = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoRegistrosTotalesPorNumYFecha"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lcNserie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lcNdocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 20, lcRazonSocial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If ldIdEstado <> 0 Then
        oRecordset.Filter = "IdEstadoNota=" & ldIdEstado
   End If
   Set NotaCreditoRegistrosTotalesPorNumYFecha = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function NotaCreditoDebitoSeleccionarPorId(lnIdNota As Long) As DoNotaCreditoDebito
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebito As New NotaCreditoDebito
    Dim oDoNotaCreditoDebito As New DoNotaCreditoDebito
    
    Set NotaCreditoDebitoSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oNotaCreditoDebito.Conexion = oConexion
    oDoNotaCreditoDebito.IdNota = lnIdNota
    If oNotaCreditoDebito.SeleccionarPorId(oDoNotaCreditoDebito) Then
        Set NotaCreditoDebitoSeleccionarPorId = oDoNotaCreditoDebito
    Else
        Set NotaCreditoDebitoSeleccionarPorId = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oNotaCreditoDebito.MensajeError
    Set oConexion = Nothing
    Set oNotaCreditoDebito = Nothing
End Function

Function NotaCreditoDebitoCargarMotivo() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoDebitoCargarMotivo = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoDebitoCargarMotivo"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoDebitoCargarMotivo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function NotaCreditoFarmNotaIngreso(lcDocumentoNumero As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoFarmNotaIngreso = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoFarmNotaIngreso"
        Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, lcDocumentoNumero): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoFarmNotaIngreso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function NotaCreditoBuscaPorIdComprobante(lnIdComprobante As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoBuscaPorIdComprobante = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoBuscaPorIdComprobante"
        Set oParameter = .CreateParameter("@IdComprobante", adInteger, adParamInput, 0, lnIdComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoBuscaPorIdComprobante = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function NotaCreditoConsultarOrdenServicio(lnIdComprobante As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoConsultarOrdenServicio = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoConsultarOrdenServicio"
        Set oParameter = .CreateParameter("@IdComprobante", adInteger, adParamInput, 0, lnIdComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoConsultarOrdenServicio = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function NotaCreditoConsultarCitaPorNCuenta(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoConsultarCitaPorNCuenta = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoConsultarCitaPorNCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoConsultarCitaPorNCuenta = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'SUNAT
Function NotaCreditoRegistrosTotalesPorNumero(lcNserie As String, lcNdocumento As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoRegistrosTotalesPorNumero = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoRegistrosTotalesPorNumero"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lcNserie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lcNdocumento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoRegistrosTotalesPorNumero = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'SUNAT

Function NotaCreditoDevueltosPorNumYFecha(lcNserie As String, lcNdocumento As String, dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoDevueltosPorNumYFecha = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoDevueltosPorNumYFecha"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lcNserie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lcNdocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoDevueltosPorNumYFecha = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'debb-17/03/2016
Function CajaComprobantesPagoPorFechasConNombreCajas(dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoPorFechasConNombreCajas = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoPorFechasConNombreCajas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoPorFechasConNombreCajas = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



'debb-21/03/2016
Function CajaComprobantesPagoXfechasExoneradosServicios(dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoXfechasExoneradosServicios = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoXfechasExoneradosServicios"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoXfechasExoneradosServicios = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-21/03/2016
Function CajaComprobantesPagoXfechasExoneradosFarmacia(dtFdesde As Date, dtFHasta As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantesPagoXfechasExoneradosFarmacia = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoXfechasExoneradosFarmacia"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantesPagoXfechasExoneradosFarmacia = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function DevuelvePrecioCPTporProductoTfinanciamiento(lnIdProducto As Long, lnIdtipoFinanciamiento As Long, oConexion As Connection) As Double
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim oRsTmp841 As New Recordset
    Set oRsTmp841 = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(lnIdProducto, _
                                                                                           lnIdtipoFinanciamiento, _
                                                                                            oConexion)
    DevuelvePrecioCPTporProductoTfinanciamiento = 0
    If oRsTmp841.RecordCount > 0 Then
       DevuelvePrecioCPTporProductoTfinanciamiento = oRsTmp841!PrecioUnitario
    End If
    oRsTmp841.Close

    Set oRsTmp841 = Nothing
    Set mo_ReglasFacturacion = Nothing
End Function


Function CajaComprobantesPagoSeleccionarPorMovnumero(lcMovNumero As String, lcMovTipo As String, oConexion As Connection) As Recordset
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
        .CommandText = "CajaComprobantesPagoSeleccionarPorMovnumero"

        Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute

   End With
   Set CajaComprobantesPagoSeleccionarPorMovnumero = oRecordset




   Set oCommand = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
   Exit Function
End Function
'SUNAT
Function CajaComprobantePagoProductosPorIdComprobante(lnIdComprobante As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoProductosPorIdComprobante = Nothing
    ms_MensajeError = ""
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoProductosPorIdComprobante"
        'Facturador Sunat
        Set oParameter = .CreateParameter("@IdComprobante", adInteger, adParamInput, 0, lnIdComprobante): .Parameters.Append oParameter
        'Facturador Sunat
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoProductosPorIdComprobante = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function
'SUNAT
'SUNAT
Function CajaComprobantePagoServiciosPorIdComprobante(lnIdComprobante As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoServiciosPorIdComprobante = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoServiciosPorIdComprobantePago"
        'Facturador Sunat
        Set oParameter = .CreateParameter("@IdComprobante", adInteger, adParamInput, 0, lnIdComprobante): .Parameters.Append oParameter
        'Facturador Sunat
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoServiciosPorIdComprobante = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function
'SUNAT
'SUNAT   kike 2017
Function NotaCreditoParaSunatPorNumYFecha(dtFdesde As Date, dtFHasta As Date, lcNserie As String, lcNdocumento As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoParaSunatPorNumYFecha = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoParaSunatPorNumYFecha"
        Set oParameter = .CreateParameter("@FDesde", adDBTimeStamp, adParamInput, 0, dtFdesde): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FHasta", adDBTimeStamp, adParamInput, 0, dtFHasta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nSerie", adVarChar, adParamInput, 4, lcNserie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nDocumento", adVarChar, adParamInput, 12, lcNdocumento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoParaSunatPorNumYFecha = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SUNAT
Function NotaCreditoDebitoTipoNotaModificar(oDONotaCreditoDebitoTipoNota As DONotaCreditoDebitoTipoNota) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oNotaCreditoDebitoTipoNota As New NotaCreditoDebitoTipoNota
    
    NotaCreditoDebitoTipoNotaModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oNotaCreditoDebitoTipoNota.Conexion = oConexion
    If oNotaCreditoDebitoTipoNota.modificar(oDONotaCreditoDebitoTipoNota) Then
        NotaCreditoDebitoTipoNotaModificar = True
    Else
        ms_MensajeError = oNotaCreditoDebitoTipoNota.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oNotaCreditoDebitoTipoNota = Nothing
End Function
'SUNAT
'SUNAT
Function NotaCreditoDebitoSeleccionarPorNroSerie(lcNroSerie As String, lnIdTipoComprobante As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set NotaCreditoDebitoSeleccionarPorNroSerie = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoDebitoSeleccionarPorNroSerie"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lcNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, lnIdTipoComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set NotaCreditoDebitoSeleccionarPorNroSerie = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SUNAT

Function CajaNroDocumentoSeleccionarPorId(idCaja As Long, idTipoComprobante As Long) As DOCajaNroDocumento
Dim oConexion As New ADODB.Connection
Dim oCajaNroDocumento As New CajaNroDocumento
Dim oDoCajaNroDocumento As New DOCajaNroDocumento
    Set CajaNroDocumentoSeleccionarPorId = Nothing
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCajaNroDocumento.Conexion = oConexion
    oDoCajaNroDocumento.idCaja = idCaja
    oDoCajaNroDocumento.idTipoComprobante = idTipoComprobante
    If oCajaNroDocumento.SeleccionarPorId(oDoCajaNroDocumento) Then
        Set CajaNroDocumentoSeleccionarPorId = oDoCajaNroDocumento
    End If
    oConexion.Close
    ms_MensajeError = oCajaNroDocumento.MensajeError
    Set oConexion = Nothing
    Set oCajaNroDocumento = Nothing
End Function

Function CajaNroDocumentoModificar(oDoCajaNroDocumento As DOCajaNroDocumento) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oCajaNroDocumento As New CajaNroDocumento
    
    CajaNroDocumentoModificar = False
    oConexion.Open sighentidades.CadenaConexion
    
    Set oCajaNroDocumento.Conexion = oConexion
    If oCajaNroDocumento.modificar(oDoCajaNroDocumento) Then
        CajaNroDocumentoModificar = True
    Else
        ms_MensajeError = oCajaNroDocumento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCajaNroDocumento = Nothing
End Function

Function CajaNroDocumentoFarmVentaSeleccionarPorNroSerie(lcNroSerie As String, lnIdTipoComprobante As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaNroDocumentoFarmVentaSeleccionarPorNroSerie = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaNroDocumentoFarmVentaSeleccionarNroMaximo"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lcNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@DocumentoIdtipo", adInteger, adParamInput, 0, lnIdTipoComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaNroDocumentoFarmVentaSeleccionarPorNroSerie = oRecordset

   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'kike 2017
Function DevuelveCorrelativoDeFacturaBoletaAnulada(lcNroSerie As String, lcNroDocumento As String) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnCorr1 As Integer
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantePagoAnuladosSunat"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   DevuelveCorrelativoDeFacturaBoletaAnulada = "000"
   lnCorr1 = 1
   If oRecordset.RecordCount > 0 Then
      Do While Not oRecordset.EOF
         If oRecordset!NroSerie = lcNroSerie And oRecordset!NroDocumento = lcNroDocumento Then
            DevuelveCorrelativoDeFacturaBoletaAnulada = Right("000" & Trim(Str(lnCorr1)), 3)
            Exit Do
         End If
         lnCorr1 = lnCorr1 + 1
         oRecordset.MoveNext
      Loop
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


'debb-14/03/2017
Function NotaCreditoDebitoUltimoDctoMasUNO(lcSerie As String, oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim lnLen As Integer
    ms_MensajeError = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "NotaCreditoDebitoUltimoDcto"
        Set oParameter = .CreateParameter("@nroSerie", adVarChar, adParamInput, 4, lcSerie): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount = 0 Then
      NotaCreditoDebitoUltimoDctoMasUNO = ""
   Else
      lnLen = Len(oRecordset!NroDocumento)
      NotaCreditoDebitoUltimoDctoMasUNO = Right("000000000" & Trim(Str(Val(oRecordset!NroDocumento) + 1)), lnLen)
   End If
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
Resume
End Function


Function CajaCajaSegunFiltro(lcFiltro As String) As Recordset
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
          .CommandText = "CajaCajaSegunFiltro"
          Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        Set CajaCajaSegunFiltro = oRecordset
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Function


Function CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision(lnNroSerie As String, lnNroDcto As String, ldFechaInicio As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion    'lcFechaInicio='22/01/2009 00:00:00'
        .CommandTimeout = 150                'lcFechaFin='22/01/2009 23:59:00'
        .CommandText = "CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision"
        Set oParameter = .CreateParameter("@NroSerie", adVarChar, adParamInput, 4, lnNroSerie): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 8, lnNroDcto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function



Function CajaComprobantesPagoActualizaFechaCobranza(lnIdComprobantePago As Long, ldFechaCobranza As Date) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    CajaComprobantesPagoActualizaFechaCobranza = False
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoActualizaFechaCobranza"
        Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranza", adDBTimeStamp, adParamInput, 0, IIf(ldFechaCobranza = 0, Null, ldFechaCobranza)): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   CajaComprobantesPagoActualizaFechaCobranza = True
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function




Function LaFacturaEsUnReemboloso(lnIdComprobantePago As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    LaFacturaEsUnReemboloso = False
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactReembolsosDocumentosXIdComprobantePago"
        Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
      LaFacturaEsUnReemboloso = True
   End If
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function




Function DevuelveMovimientoEnLabImaFar(lnIdOrdenPago As Long, lnIdTipoOrden As Long, lnIdCuenta As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    DevuelveMovimientoEnLabImaFar = ""
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveMovimientoEnLabImaFar"
        Set oParameter = .CreateParameter("@movimiento", adVarChar, adParamOutput, 20, ""): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoOrden", adInteger, adParamInput, 0, lnIdTipoOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
        .Execute
        If .Parameters("@movimiento") <> "." Then
           DevuelveMovimientoEnLabImaFar = " >> " & .Parameters("@movimiento")
        End If
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function




Function NotaCreditoDebitoSeleccionarPorIdComprobante(lnIdComprobantePago As Long, oConexion As Connection) As Recordset
        Dim oRecordset As New ADODB.Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "NotaCreditoDebitoSeleccionarPorIdComprobante"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRecordset = .Execute
        End With
        Set NotaCreditoDebitoSeleccionarPorIdComprobante = oRecordset
        Set oCommand = Nothing
End Function

Function FactCatalogoServiciosSeleccionarPorId(lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosSeleccionarPorId"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactCatalogoServiciosSeleccionarPorId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function



Function sunat_onuBuscarPorCodigoDescripcion(lcCodigo As String, lcDescripcion As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "sunat_onuBuscarPorCodigoDescripcion"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@descripcion", adVarChar, adParamInput, 100, lcDescripcion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set sunat_onuBuscarPorCodigoDescripcion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function



Function sunat_ResumenDiarioUltimoCorrelativo(oConexion As Connection) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim lnCorr1 As Integer
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "sunat_ResumenDiarioUltimoCorrelativo"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   sunat_ResumenDiarioUltimoCorrelativo = "001"
   lnCorr1 = 1
   If oRecordset.RecordCount > 0 Then
      sunat_ResumenDiarioUltimoCorrelativo = Right("000" & Trim(Str(Val(Right(oRecordset!DctoRDI, 3)) + 1)), 3)
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub CajaComprobantesPagoActualizaDatosSunat(lnIdComprobantePago As Long, TotalVentasInaf As Double, SumatoriaIGV As Double, _
                                             oConexion As Connection)
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
        .CommandText = "CajaComprobantesPagoActualizaDatosSunat"
        Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SunatTotalIgv", adCurrency, adParamInput, 0, SumatoriaIGV): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SunatOpeInafectas", adCurrency, adParamInput, 0, TotalVentasInaf): .Parameters.Append oParameter
        .Execute
   End With
   Set oCommand = Nothing
   Set oRecordset = Nothing
   
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Sub


Function Sunat_ResumenDiarioSeleccionarPorFechaEmision(lcFechaEmision As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunat_ResumenDiarioSeleccionarPorFechaEmision"
        Set oParameter = .CreateParameter("@FechaEmision", adVarChar, adParamInput, 10, Left(lcFechaEmision, 10)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Sunat_ResumenDiarioSeleccionarPorFechaEmision = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function


Sub Sunat_ResumenDiarioEliminarPorFechaEmision(lcFechaEmision As String)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "Sunat_ResumenDiarioEliminarPorFechaEmision"
        Set oParameter = .CreateParameter("@FechaEmision", adVarChar, adParamInput, 10, Left(lcFechaEmision, 10)): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Sub

Function CajaComprobantesPagoActualizaTieneCredito(lnIdComprobantePago As Long, lcTieneCredito As String) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    CajaComprobantesPagoActualizaTieneCredito = False
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantesPagoActualizaTieneCredito"
        Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TieneCredito", adVarChar, adParamInput, 1, IIf(lcTieneCredito = "", Null, lcTieneCredito)): .Parameters.Append oParameter
        .Execute
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   CajaComprobantesPagoActualizaTieneCredito = True
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"

End Function

Function CajaComprobantePagoXmovnumero(lcMovNumero As String, lcMovTipo As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    ms_MensajeError = ""
    sighentidades.AbreConexionSIGH oConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CajaComprobantePagoXmovnumero"
        Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@movtipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CajaComprobantePagoXmovnumero = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function
