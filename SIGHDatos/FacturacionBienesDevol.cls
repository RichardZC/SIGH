VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionBienesDevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionBienesDevoluciones
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionBienesDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesDevolucionesAgregar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadAdevolver", adInteger, adParamInput, 0, IIf(oTabla.CantidadAdevolver = 0, Null, oTabla.CantidadAdevolver))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoDevolucion", adInteger, adParamInput, 0, oTabla.IdEstadoDevolucion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovNumeroE", adVarChar, adParamInput, 9, IIf(oTabla.MovNumeroE = "", Null, oTabla.MovNumeroE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipoE", adVarChar, adParamInput, 1, IIf(oTabla.MovTipoE = "", Null, oTabla.MovTipoE))
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
'   Descripción:    modifica un registro a la tabla FacturacionBienesDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesDevolucionesModificar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadAdevolver", adInteger, adParamInput, 0, IIf(oTabla.CantidadAdevolver = 0, Null, oTabla.CantidadAdevolver))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoDevolucion", adInteger, adParamInput, 0, oTabla.IdEstadoDevolucion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAutoriza", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAutoriza = 0, Null, oTabla.FechaAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAutoriza", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioAutoriza = 0, Null, oTabla.IdUsuarioAutoriza))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovNumeroE", adVarChar, adParamInput, 9, IIf(oTabla.MovNumeroE = "", Null, oTabla.MovNumeroE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipoE", adVarChar, adParamInput, 1, IIf(oTabla.MovTipoE = "", Null, oTabla.MovTipoE))
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionBienesDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesDevolucionesEliminar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionBienesDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionBienesDevolucionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.CantidadAdevolver = IIf(IsNull(oRecordset!CantidadAdevolver), 0, oRecordset!CantidadAdevolver)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdEstadoDevolucion = IIf(IsNull(oRecordset!IdEstadoDevolucion), 0, oRecordset!IdEstadoDevolucion)
       oTabla.FechaAutoriza = IIf(IsNull(oRecordset!FechaAutoriza), 0, oRecordset!FechaAutoriza)
       oTabla.IdUsuarioAutoriza = IIf(IsNull(oRecordset!IdUsuarioAutoriza), 0, oRecordset!IdUsuarioAutoriza)
       oTabla.MovNumeroE = IIf(IsNull(oRecordset!MovNumeroE), "", oRecordset!MovNumeroE)
       oTabla.MovTipoE = IIf(IsNull(oRecordset!MovTipoE), "", oRecordset!MovTipoE)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFacturacionBienesDevol)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.idProducto = 0
   oTabla.CantidadAdevolver = 0
   oTabla.IdComprobantePago = 0
   oTabla.IdEstadoDevolucion = 0
   oTabla.FechaAutoriza = Now
   oTabla.IdUsuarioAutoriza = 0
End Sub

Function SeleccionarPorIdProducto(LcMovNumero As String, LcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarPorIdProducto = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesDevolucionesSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdProducto = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'RHA 07/10/20 Cambio33 Inicio
Function SeleccionarPorIdProducto1(LcMovNumero As String, LcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set SeleccionarPorIdProducto1 = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_FacturacionBienesDevolucionesSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdProducto1 = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'RHA 07/10/20 Cambio33 Fin



Function EliminarPorMovNumeroE(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarPorMovNumeroE = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionBienesDevolucionesEliminarPorMovNumeroE"
   Set oParameter = .CreateParameter("@MovNumeroE", adVarChar, adParamInput, 9, oTabla.MovNumeroE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipoE", adVarChar, adParamInput, 1, IIf(oTabla.MovTipoE = "", Null, oTabla.MovTipoE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarPorMovNumeroE = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarPorMovNumeroE(ByVal oTabla As DoFacturacionBienesDevol) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   SeleccionarPorMovNumeroE = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionBienesDevolucionesSeleccionarPorMovNumeroE"
     Set oParameter = .CreateParameter("@MovNumeroE", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipoE", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.CantidadAdevolver = IIf(IsNull(oRecordset!CantidadAdevolver), 0, oRecordset!CantidadAdevolver)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdEstadoDevolucion = IIf(IsNull(oRecordset!IdEstadoDevolucion), 0, oRecordset!IdEstadoDevolucion)
       oTabla.FechaAutoriza = IIf(IsNull(oRecordset!FechaAutoriza), 0, oRecordset!FechaAutoriza)
       oTabla.IdUsuarioAutoriza = IIf(IsNull(oRecordset!IdUsuarioAutoriza), 0, oRecordset!IdUsuarioAutoriza)
       oTabla.MovNumeroE = IIf(IsNull(oRecordset!MovNumeroE), "", oRecordset!MovNumeroE)
       oTabla.MovTipoE = IIf(IsNull(oRecordset!MovTipoE), "", oRecordset!MovTipoE)
   SeleccionarPorMovNumeroE = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
