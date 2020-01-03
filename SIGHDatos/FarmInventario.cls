VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmInventario
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
'   Descripción:    Inserta un registro a la tabla farmInventario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmInventario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioAgregar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, IIf(oTabla.idAlmacen = 0, Null, oTabla.idAlmacen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@numeroInventario", adVarChar, adParamInput, 4, IIf(oTabla.numeroInventario = "", Null, oTabla.numeroInventario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoInventario", adInteger, adParamInput, 0, oTabla.idEstadoInventario)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoInventario", adInteger, adParamInput, 0, oTabla.idTipoInventario)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idInventario = .Parameters("@idInventario")
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
'   Descripción:    modifica un registro a la tabla farmInventario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmInventario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioModificar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, IIf(oTabla.idInventario = 0, Null, oTabla.idInventario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, IIf(oTabla.idAlmacen = 0, Null, oTabla.idAlmacen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@numeroInventario", adVarChar, adParamInput, 4, IIf(oTabla.numeroInventario = "", Null, oTabla.numeroInventario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModificacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModificacion = 0, Null, oTabla.FechaModificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoInventario", adInteger, adParamInput, 0, oTabla.idEstadoInventario)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoInventario", adInteger, adParamInput, 0, oTabla.idTipoInventario)
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
'   Descripción:    elimia o borra un registro a la tabla farmInventario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmInventario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioEliminar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, oTabla.idInventario)
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
'   Descripción:    Selecciona un unico registro de la tabla farmInventario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmInventario) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmInventarioSeleccionarPorId"
     Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, oTabla.idInventario)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idInventario = IIf(IsNull(oRecordset!idInventario), 0, oRecordset!idInventario)
       oTabla.idAlmacen = IIf(IsNull(oRecordset!idAlmacen), 0, oRecordset!idAlmacen)
       oTabla.numeroInventario = IIf(IsNull(oRecordset!numeroInventario), "", oRecordset!numeroInventario)
       oTabla.FechaCierre = IIf(IsNull(oRecordset!FechaCierre), 0, oRecordset!FechaCierre)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.FechaModificacion = IIf(IsNull(oRecordset!FechaModificacion), 0, oRecordset!FechaModificacion)
       oTabla.idEstadoInventario = IIf(IsNull(oRecordset!idEstadoInventario), 0, oRecordset!idEstadoInventario)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.idTipoInventario = IIf(IsNull(oRecordset!idTipoInventario), 0, oRecordset!idTipoInventario)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmInventario)
   oTabla.idInventario = 0
   oTabla.idAlmacen = 0
   oTabla.numeroInventario = ""
   oTabla.FechaCierre = Now
   oTabla.FechaCreacion = Now
   oTabla.FechaModificacion = Now
   oTabla.idEstadoInventario = 0
   oTabla.IdUsuario = 0
   oTabla.idTipoInventario = 1
End Sub

Function DevuelveListaInventarios() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "farmInventarioDevuelveListaInventarios"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveListaInventarios = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



