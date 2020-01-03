VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_DatosEstablecimiento"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_DatosEstablecimiento
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdDatoEstablec As Long
Dim ml_IdEstablecimiento As Long
Dim ms_Color As String
Dim ml_Turnos As Long
Dim ml_UltimoNroFormatoHIS As Long

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_DatosEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_DatosEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DatosEstablecimientoAgregar"
   Set oParameter = .CreateParameter("@IdDatoEstablec", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Color", adVarChar, adParamInput, 30, IIf(oTabla.Color = "", Null, oTabla.Color))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Turnos", adInteger, adParamInput, 0, IIf(oTabla.Turnos = 0, Null, oTabla.Turnos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@UltimoNroFormatoHIS", adInteger, adParamInput, 0, IIf(oTabla.UltimoNroFormatoHIS = 0, Null, oTabla.UltimoNroFormatoHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdDatoEstablec = .Parameters("@IdDatoEstablec")
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
'   Descripción:    modifica un registro a la tabla HIS_DatosEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_DatosEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DatosEstablecimientoModificar"
   Set oParameter = .CreateParameter("@IdDatoEstablec", adInteger, adParamInput, 0, IIf(oTabla.IdDatoEstablec = 0, Null, oTabla.IdDatoEstablec))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Color", adVarChar, adParamInput, 30, IIf(oTabla.Color = "", Null, oTabla.Color))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Turnos", adInteger, adParamInput, 0, IIf(oTabla.Turnos = 0, Null, oTabla.Turnos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@UltimoNroFormatoHIS", adInteger, adParamInput, 0, IIf(oTabla.UltimoNroFormatoHIS = 0, Null, oTabla.UltimoNroFormatoHIS))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_DatosEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_DatosEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DatosEstablecimientoEliminar"
   Set oParameter = .CreateParameter("@IdDatoEstablec", adInteger, adParamInput, 0, oTabla.IdDatoEstablec)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_DatosEstablecimiento
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_DatosEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_DatosEstablecimientoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdDatoEstablec", adInteger, adParamInput, 0, oTabla.IdDatoEstablec)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdDatoEstablec = IIf(IsNull(oRecordset!IdDatoEstablec), 0, oRecordset!IdDatoEstablec)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.Color = IIf(IsNull(oRecordset!Color), "", oRecordset!Color)
       oTabla.Turnos = IIf(IsNull(oRecordset!Turnos), 0, oRecordset!Turnos)
       oTabla.UltimoNroFormatoHIS = IIf(IsNull(oRecordset!UltimoNroFormatoHIS), 0, oRecordset!UltimoNroFormatoHIS)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_DatosEstablecimiento)
   oTabla.IdDatoEstablec = 0
   oTabla.IdEstablecimiento = 0
   oTabla.Color = ""
   oTabla.Turnos = 0
   oTabla.UltimoNroFormatoHIS = 0
End Sub
