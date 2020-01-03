VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionPagosACuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FacturacionPagosAcuenta
'        Autor: William Castro Grijalva
'        Fecha: 14/04/2006 07:06:56 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla FacturacionPagosAcuenta
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFacturacionPAgosACuenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionPagosAcuentaAgregar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPagosACuenta", adInteger, adParamOutput, 0, IIf(oTabla.IdPagosACuenta = 0, Null, oTabla.IdPagosACuenta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPago", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPago = 0, Null, oTabla.FechaPago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoCajero", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoCajero = 0, Null, oTabla.IdEmpleadoCajero)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdPagosACuenta = .Parameters("@IdPagosACuenta")
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
'   Descripción:    modifica un registro a la tabla FacturacionPagosAcuenta
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionPAgosACuenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionPagosAcuentaModificar"
           Set oParameter = .CreateParameter("@IdPagosACuenta", adInteger, adParamInput, 0, IIf(oTabla.IdPagosACuenta = 0, Null, oTabla.IdPagosACuenta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPago", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPago = 0, Null, oTabla.FechaPago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoCajero", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoCajero = 0, Null, oTabla.IdEmpleadoCajero)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionPagosAcuenta
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionPAgosACuenta) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionPagosAcuentaEliminar"
           Set oParameter = .CreateParameter("@IdPagosACuenta", adInteger, adParamInput, 0, IIf(oTabla.IdPagosACuenta = 0, Null, oTabla.IdPagosACuenta)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionPagosAcuenta
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionPAgosACuenta) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionPagosAcuentaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPagosACuenta", adInteger, adParamInput, 0, oTabla.IdPagosACuenta): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdPagosACuenta = IIf(IsNull(oRecordset!IdExoneracion), 0, oRecordset!IdExoneracion)
       oTabla.IdAtencion = IIf(IsNull(oRecordset!IdAtencion), 0, oRecordset!IdAtencion)
       oTabla.TotalPagado = IIf(IsNull(oRecordset!TotalPagado), 0, oRecordset!TotalPagado)
       oTabla.FechaPago = IIf(IsNull(oRecordset!FechaPago), "", oRecordset!FechaPago)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
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


Function SeleccionarPorCuentaAtencion(IdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String
Dim sOrder As String
    
    sOrder = ""
    sSQL = " select idPagosACuenta as id, a.idAtencion, "
    sSQL = sSQL + " a.fechaPago, a.IdComprobantePago,a.IdEmpleadoCajero, "
    sSQL = sSQL + " b.ApellidoPaterno as Empleado, "
    sSQL = sSQL + " b.ApellidoPaterno + ' ' +  b.ApellidoMaterno + ', ' + b.Nombres as NombresEmpleado,a.totalPagado "
    
    sSQL = sSQL + " from facturacionPagosACuenta a, empleados b "
    sSQL = sSQL + " where idatencion in (select idatencion"
    sSQL = sSQL + " from atenciones where idcuentaatencion = " & IdCuentaAtencion & ")"
    sSQL = sSQL + " and a.idEmpleadoCajero = b.idempleado "
    sSQL = sSQL + " order by fechaPago "
    
    Set SeleccionarPorCuentaAtencion = Nothing


   oRecordset.Open sSQL, mo_Conexion, adOpenDynamic, adLockOptimistic
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function
