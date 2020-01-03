VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla CajaCajero
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
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
'   Descripción:    Inserta un registro a la tabla CajaCajero
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaCajero) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroAgregar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoCajero", adChar, adParamInput, 1, IIf(oTabla.EstadoCajero = 0, Null, oTabla.EstadoCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCajero = .Parameters("@IdCajero")
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
'   Descripción:    modifica un registro a la tabla CajaCajero
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaCajero) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroModificar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoCajero", adChar, adParamInput, 1, IIf(oTabla.EstadoCajero = 0, Null, oTabla.EstadoCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaCajero
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaCajero) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajeroEliminar"
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaCajero
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaCajero) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaCajeroSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, oTabla.IdCajero): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEmpleado = IIf(IsNull(oRecordset!IdEmpleado), 0, oRecordset!IdEmpleado)
       oTabla.EstadoCajero = IIf(IsNull(oRecordset!EstadoCajero), 0, oRecordset!EstadoCajero)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaCajeroSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarTodosParaLista() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaCajeroSeleccionarTodosParaLista"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodosParaLista = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function SeleccionarSegunUsuario(IdEmpleado As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "cajacajeroSeleccionarXidEmpleado"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IdEmpleado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarSegunUsuario = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function RealizarFiltro(oEmpleado As dOEmpleado) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String, sWhere As String
    SQL = ""
    sWhere = ""
    If oEmpleado.ApellidoPaterno <> "" Then
         SQL = SQL + " and e.ApellidoPaterno like '%" & oEmpleado.ApellidoPaterno & "%'"
    End If
    If oEmpleado.ApellidoMaterno <> "" Then
         SQL = SQL + " and e.ApellidoMaterno like '%" & oEmpleado.ApellidoMaterno & "%'"
    End If
    If oEmpleado.Nombres <> "" Then
         SQL = SQL + " and e.Nombres like '%" & oEmpleado.Nombres & "%'"
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosNoMinsaFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, SQL): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RealizarFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function SeleccionarSegunUsuarioYCaja(doCajero As DOCajaCajero) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "cajacajeroSeleccionarSegunUsuarioYCaja"
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, doCajero.IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, doCajero.IdEmpleado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   SeleccionarSegunUsuarioYCaja = False
   If (oRecordset.RecordCount > 0) Then
        SeleccionarSegunUsuarioYCaja = True
   End If
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


