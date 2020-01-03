VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipoFinanciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TipoFinanciador
'        Programado por: Garay M
'        Fecha: Marzo 2015
'
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla TipoFinanciador
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoTipoFinanciador) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TipoFinanciadorAgregar"
   Set oParameter = .CreateParameter("@idTipoFinanciador", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 50, IIf(oTabla.nombre = "", Null, oTabla.nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@denominacion", adVarChar, adParamInput, 50, IIf(oTabla.denominacion = "", Null, oTabla.denominacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 11, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idTipoFinanciador = .Parameters("@idTipoFinanciador")
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
'   Descripción:    modifica un registro a la tabla TipoFinanciador
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoTipoFinanciador) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TipoFinanciadorModificar"
   Set oParameter = .CreateParameter("@idTipoFinanciador", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciador = 0, Null, oTabla.idTipoFinanciador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 50, IIf(oTabla.nombre = "", Null, oTabla.nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@denominacion", adVarChar, adParamInput, 50, IIf(oTabla.denominacion = "", Null, oTabla.denominacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 11, IIf(oTabla.codigo = "", Null, oTabla.codigo))
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
'   Descripción:    elimia o borra un registro a la tabla TipoFinanciador
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoTipoFinanciador) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TipoFinanciadorEliminar"
   Set oParameter = .CreateParameter("@idTipoFinanciador", adInteger, adParamInput, 0, oTabla.idTipoFinanciador)
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
'   Descripción:    Selecciona un unico registro de la tabla TipoFinanciador
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoTipoFinanciador) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TipoFinanciadorSeleccionarPorId"
     Set oParameter = .CreateParameter("@idTipoFinanciador", adInteger, adParamInput, 0, oTabla.idTipoFinanciador)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idTipoFinanciador = IIf(IsNull(oRecordset!idTipoFinanciador), 0, oRecordset!idTipoFinanciador)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.denominacion = IIf(IsNull(oRecordset!denominacion), "", oRecordset!denominacion)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TipoFinanciadorSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
    Set oCommand = Nothing
   Set oRecordset = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub SetDefaults(ByVal oTabla As DoTipoFinanciador)
   oTabla.idTipoFinanciador = 0
   oTabla.nombre = ""
   oTabla.denominacion = ""
   oTabla.codigo = ""
End Sub
