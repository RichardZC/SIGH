VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enfermeria_Visitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Enfermeria_Visitas
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla Enfermeria_Visitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOEnfermeria_Visitas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_VisitasAgregar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, oTabla.IdVisita)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraVisita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraVisita = 0, Null, oTabla.FechaHoraVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 1000, oTabla.Observaciones)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, oTabla.IdCama)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoEnfermera", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoEnfermera = 0, Null, oTabla.IdEmpleadoEnfermera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IngresoValorizacion", adBoolean, adParamInput, 0, oTabla.IngresoValorizacion)
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
'   Descripción:    modifica un registro a la tabla Enfermeria_Visitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEnfermeria_Visitas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_VisitasModificar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, IIf(oTabla.IdVisita = 0, Null, oTabla.IdVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraVisita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraVisita = 0, Null, oTabla.FechaHoraVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 1000, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEmpleadoEnfermera", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoEnfermera = 0, Null, oTabla.IdEmpleadoEnfermera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IngresoValorizacion", adBoolean, adParamInput, 0, oTabla.IngresoValorizacion)
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
'   Descripción:    elimia o borra un registro a la tabla Enfermeria_Visitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal lnIdCuentaAtencion As Long, ByVal lnIdVisita As Integer) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_VisitasEliminar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, lnIdVisita)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 0)
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
'   Descripción:    Selecciona un unico registro de la tabla Enfermeria_Visitas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEnfermeria_Visitas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Enfermeria_VisitasSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdVisita = IIf(IsNull(oRecordset!IdVisita), 0, oRecordset!IdVisita)
       oTabla.FechaHoraVisita = IIf(IsNull(oRecordset!FechaHoraVisita), 0, oRecordset!FechaHoraVisita)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       oTabla.IdCama = IIf(IsNull(oRecordset!IdCama), 0, oRecordset!IdCama)
       oTabla.IdEmpleadoEnfermera = IIf(IsNull(oRecordset!IdEmpleadoEnfermera), 0, oRecordset!IdEmpleadoEnfermera)
       oTabla.IngresoValorizacion = IIf(IsNull(oRecordset!IngresoValorizacion), 0, oRecordset!IngresoValorizacion)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOEnfermeria_Visitas)
   oTabla.IdCuentaAtencion = 0
   oTabla.IdVisita = 0
   oTabla.FechaHoraVisita = Now
   oTabla.IdServicio = 0
   oTabla.Observaciones = ""
   oTabla.IdCama = 0
   oTabla.IdEmpleadoEnfermera = 0
   oTabla.IngresoValorizacion = True
End Sub
