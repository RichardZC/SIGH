VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstadosAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla EstadosAtencion
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
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
'   Funci�n:        Insertar
'   Descripci�n:    Inserta un registro a la tabla EstadosAtencion
'   Par�metros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEstadosAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstadosAtencionAgregar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 20, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoAtencion", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.idEstadoAtencion = .Parameters("@IdEstadoAtencion")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Funci�n:        Modificar
'   Descripci�n:    modifica un registro a la tabla EstadosAtencion
'   Par�metros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEstadosAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstadosAtencionModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 20, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoAtencion", adInteger, adParamInput, 0, IIf(oTabla.idEstadoAtencion = 0, Null, oTabla.idEstadoAtencion)): .Parameters.Append oParameter
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
'   Funci�n:        Eliminar
'   Descripci�n:    elimia o borra un registro a la tabla EstadosAtencion
'   Par�metros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEstadosAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstadosAtencionEliminar"
           Set oParameter = .CreateParameter("@IdEstadoAtencion", adInteger, adParamInput, 0, IIf(oTabla.idEstadoAtencion = 0, Null, oTabla.idEstadoAtencion)): .Parameters.Append oParameter
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
'   Funci�n:        SeleccionarPorId
'   Descripci�n:    Selecciona un unico registro de la tabla EstadosAtencion
'   Par�metros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEstadosAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EstadosAtencionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEstadoAtencion", adInteger, adParamInput, 0, oTabla.idEstadoAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.idEstadoAtencion = IIf(IsNull(oRecordset!idEstadoAtencion), 0, oRecordset!idEstadoAtencion)
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


Function SeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstadosAtencionSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

