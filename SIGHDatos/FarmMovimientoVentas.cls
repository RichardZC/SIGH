VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmMovimientoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FarmMovimientoVentas
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmMovimientoVentas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmMovimientoVentas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoVentasAgregar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tipoVenta", adVarChar, adParamInput, 1, IIf(oTabla.tipoVenta = "", Null, oTabla.tipoVenta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPreVenta", adInteger, adParamInput, 0, IIf(oTabla.idPreventa = 0, Null, oTabla.idPreventa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adInteger, adParamInput, 0, IIf(oTabla.idPrescriptor = 0, Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoReceta", adInteger, adParamInput, 0, oTabla.idTipoReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPaciente = 0, Null, oTabla.IdServicioPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oTabla.idFuenteFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla farmMovimientoVentas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmMovimientoVentas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoVentasModificar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tipoVenta", adVarChar, adParamInput, 1, IIf(oTabla.tipoVenta = "", Null, oTabla.tipoVenta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPreVenta", adInteger, adParamInput, 0, IIf(oTabla.idPreventa = 0, Null, oTabla.idPreventa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adInteger, adParamInput, 0, IIf(oTabla.idPrescriptor = 0, Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoReceta", adInteger, adParamInput, 0, oTabla.idTipoReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicioPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPaciente = 0, Null, oTabla.IdServicioPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla farmMovimientoVentas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmMovimientoVentas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoVentasEliminar"
   Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla farmMovimientoVentas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmMovimientoVentas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmMovimientoVentasSeleccionarPorId"
     Set oParameter = .CreateParameter("@movNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@movTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.tipoVenta = IIf(IsNull(oRecordset!tipoVenta), "", oRecordset!tipoVenta)
       oTabla.idPreventa = IIf(IsNull(oRecordset!idPreventa), 0, oRecordset!idPreventa)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.idPrescriptor = IIf(IsNull(oRecordset!idPrescriptor), 0, oRecordset!idPrescriptor)
       oTabla.idTipoReceta = IIf(IsNull(oRecordset!idTipoReceta), 0, oRecordset!idTipoReceta)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdServicioPaciente = IIf(IsNull(oRecordset!IdServicioPaciente), 0, oRecordset!IdServicioPaciente)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.FechaHoraPrescribe = IIf(IsNull(oRecordset!FechaHoraPrescribe), 0, oRecordset!FechaHoraPrescribe)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmMovimientoVentas)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.tipoVenta = ""
   oTabla.idPreventa = 0
   oTabla.idTipoFinanciamiento = 0
   oTabla.idPrescriptor = 0
   oTabla.idTipoReceta = 0
   oTabla.IdDiagnostico = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.IdServicioPaciente = 0
   oTabla.IdPaciente = 0
End Sub



 
Function SeleccionarPorCuentaAtencion(lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarPorCuentaAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorCuentaAtencion = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
