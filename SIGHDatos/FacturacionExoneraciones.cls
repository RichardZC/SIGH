VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionExoneraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FacturacionExoneraciones
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
'   Descripción:    Inserta un registro a la tabla FacturacionExoneraciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFacturacionExoneraciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionExoneracionesAgregar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdExoneracion", adInteger, adParamOutput, 0, IIf(oTabla.IdExoneracion = 0, Null, oTabla.IdExoneracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaExoneracion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaExoneracion = 0, Null, oTabla.FechaExoneracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoExonera", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoExonera = 0, Null, oTabla.IdEmpleadoExonera)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdExoneracion = .Parameters("@IdExoneracion")
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
'   Descripción:    modifica un registro a la tabla FacturacionExoneraciones
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionExoneraciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionExoneracionesModificar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdAtencion = 0, Null, oTabla.IdAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleadoExonera", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleadoExonera = 0, Null, oTabla.IdEmpleadoExonera)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaExoneracion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaExoneracion = 0, Null, oTabla.FechaExoneracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdExoneracion", adInteger, adParamInput, 0, IIf(oTabla.IdExoneracion = 0, Null, oTabla.IdExoneracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionExoneraciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionExoneraciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionExoneracionesEliminar"
           Set oParameter = .CreateParameter("@IdExoneracion", adInteger, adParamInput, 0, IIf(oTabla.IdExoneracion = 0, Null, oTabla.IdExoneracion)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionExoneraciones
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionExoneraciones) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionExoneracionesSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdExoneracion", adInteger, adParamInput, 0, oTabla.IdExoneracion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdExoneracion = IIf(IsNull(oRecordset!IdExoneracion), 0, oRecordset!IdExoneracion)
       oTabla.IdAtencion = IIf(IsNull(oRecordset!IdAtencion), 0, oRecordset!IdAtencion)
       oTabla.FechaExoneracion = IIf(IsNull(oRecordset!FechaExoneracion), "", oRecordset!FechaExoneracion)
       oTabla.IdEmpleadoExonera = IIf(IsNull(oRecordset!IdEmpleadoExonera), 0, oRecordset!IdEmpleadoExonera)
       oTabla.TotalExonerado = IIf(IsNull(oRecordset!TotalExonerado), 0, oRecordset!TotalExonerado)
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
    sSQL = " select idExoneracion as idExoneracion, a.idAtencion, "
    sSQL = sSQL + " a.fechaExoneracion, a.IdEmpleadoExonera,"
    sSQL = sSQL + " b.ApellidoPaterno + ' ' + b.ApellidoMaterno + ', ' + b.Nombres as NombreEmpExonera, a.TotalExonerado "
    sSQL = sSQL + " from facturacionExoneraciones a, empleados b "
    sSQL = sSQL + " where idatencion in (select idatencion"
    sSQL = sSQL + " from atenciones where idcuentaatencion = " & IdCuentaAtencion & ")"
    sSQL = sSQL + " and a.idEmpleadoExonera = b.idempleado "
    sSQL = sSQL + " order by fechaExoneracion "
    
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

