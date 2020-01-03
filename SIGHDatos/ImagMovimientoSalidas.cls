VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagMovimientoSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ImagMovimientoSalidas
'        Programado por: Barrantes D
'        Fecha: Abril 2009
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
'   Descripción:    Inserta un registro a la tabla ImagMovimientoSalidas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOImagMovimientoSalidas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoSalidasAgregar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, IIf(oTabla.IdResponsable = 0, Null, oTabla.IdResponsable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivoSalida", adInteger, adParamInput, 0, IIf(oTabla.IdMotivoSalida = 0, Null, oTabla.IdMotivoSalida))
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
'   Descripción:    modifica un registro a la tabla ImagMovimientoSalidas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOImagMovimientoSalidas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoSalidasModificar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsable", adInteger, adParamInput, 0, IIf(oTabla.IdResponsable = 0, Null, oTabla.IdResponsable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMotivoSalida", adInteger, adParamInput, 0, IIf(oTabla.IdMotivoSalida = 0, Null, oTabla.IdMotivoSalida))
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
'   Descripción:    elimia o borra un registro a la tabla ImagMovimientoSalidas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOImagMovimientoSalidas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoSalidasEliminar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla ImagMovimientoSalidas
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOImagMovimientoSalidas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ImagMovimientoSalidasSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
       oTabla.IdResponsable = IIf(IsNull(oRecordset!IdResponsable), 0, oRecordset!IdResponsable)
       oTabla.IdMotivoSalida = IIf(IsNull(oRecordset!IdMotivoSalida), "", oRecordset!IdMotivoSalida)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOImagMovimientoSalidas)
   oTabla.IdMovimiento = 0
   oTabla.IdResponsable = 0
   oTabla.IdMotivoSalida = 0
End Sub


