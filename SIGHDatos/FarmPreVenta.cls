VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmPreVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FarmPreventa
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
'   Descripción:    Inserta un registro a la tabla farmPreVenta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmPreVenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmPreVentaAgregar"
   Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, IIf(oTabla.idAlmacen = 0, Null, oTabla.idAlmacen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idVendedor", adInteger, adParamInput, 0, IIf(oTabla.idVendedor = 0, Null, oTabla.idVendedor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.TOTAL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoReceta", adInteger, adParamInput, 0, oTabla.idTipoReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adInteger, adParamInput, 0, IIf(oTabla.idPrescriptor = 0, Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraCreacion", adVarChar, adParamInput, 5, IIf(oTabla.HoraCreacion = "", Null, oTabla.HoraCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoPreventa", adInteger, adParamInput, 0, oTabla.idEstadoPreventa)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dni", adVarChar, adParamInput, 8, IIf(oTabla.dni = "", Null, oTabla.dni))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
   .Parameters.Append oParameter
   
   
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idPreventa = .Parameters("@idPreventa")
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
'   Descripción:    modifica un registro a la tabla farmPreVenta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmPreVenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmPreVentaModificar"
   Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, IIf(oTabla.idPreventa = 0, Null, oTabla.idPreventa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, IIf(oTabla.idAlmacen = 0, Null, oTabla.idAlmacen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idVendedor", adInteger, adParamInput, 0, IIf(oTabla.idVendedor = 0, Null, oTabla.idVendedor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.TOTAL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoReceta", adInteger, adParamInput, 0, oTabla.idTipoReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPrescriptor", adInteger, adParamInput, 0, IIf(oTabla.idPrescriptor = 0, Null, oTabla.idPrescriptor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HoraCreacion", adVarChar, adParamInput, 5, IIf(oTabla.HoraCreacion = "", Null, oTabla.HoraCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioModifica", adInteger, adParamInput, 0, IIf(oTabla.idUsuarioModifica = 0, Null, oTabla.idUsuarioModifica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoPreventa", adInteger, adParamInput, 0, oTabla.idEstadoPreventa)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraPrescribe", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraPrescribe = 0, Null, oTabla.FechaHoraPrescribe))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dni", adVarChar, adParamInput, 8, IIf(oTabla.dni = "", Null, oTabla.dni))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
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
'   Descripción:    elimia o borra un registro a la tabla farmPreVenta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmPreVenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmPreVentaEliminar"
   Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, oTabla.idPreventa)
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
'   Descripción:    Selecciona un unico registro de la tabla farmPreVenta
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmPreVenta) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmPreVentaSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, oTabla.idPreventa)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idPreventa = IIf(IsNull(oRecordset!idPreventa), 0, oRecordset!idPreventa)
       oTabla.idAlmacen = IIf(IsNull(oRecordset!idAlmacen), 0, oRecordset!idAlmacen)
       oTabla.idVendedor = IIf(IsNull(oRecordset!idVendedor), 0, oRecordset!idVendedor)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.TOTAL = oRecordset!TOTAL
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.idTipoReceta = IIf(IsNull(oRecordset!idTipoReceta), 0, oRecordset!idTipoReceta)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.idPrescriptor = IIf(IsNull(oRecordset!idPrescriptor), 0, oRecordset!idPrescriptor)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.HoraCreacion = IIf(IsNull(oRecordset!HoraCreacion), "", oRecordset!HoraCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.idUsuarioModifica = IIf(IsNull(oRecordset!idUsuarioModifica), 0, oRecordset!idUsuarioModifica)
       oTabla.idEstadoPreventa = IIf(IsNull(oRecordset!idEstadoPreventa), 0, oRecordset!idEstadoPreventa)
       oTabla.FechaHoraPrescribe = IIf(IsNull(oRecordset!FechaHoraPrescribe), 0, oRecordset!FechaHoraPrescribe)
       oTabla.Paciente = IIf(IsNull(oRecordset!Paciente), "", oRecordset!Paciente)
       oTabla.dni = IIf(IsNull(oRecordset!dni), "", oRecordset!dni)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmPreVenta)
   oTabla.idPreventa = 0
   oTabla.idAlmacen = 0
   oTabla.idVendedor = 0
   oTabla.IdPaciente = 0
   oTabla.idTipoFinanciamiento = 0
   oTabla.TOTAL = 0
   oTabla.IdDiagnostico = 0
   oTabla.idTipoReceta = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.idPrescriptor = 0
   oTabla.FechaCreacion = Now
   oTabla.HoraCreacion = ""
   oTabla.IdUsuario = 0
   oTabla.FechaModificacion = Now
   oTabla.idUsuarioModifica = 0
   oTabla.idEstadoPreventa = 0
End Sub

