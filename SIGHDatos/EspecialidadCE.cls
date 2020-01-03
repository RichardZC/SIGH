VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EspecialidadCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla EspecialidadCE
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
'   Descripción:    Inserta un registro a la tabla EspecialidadCE
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEspecialidadCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EspecialidadCEAgregar"
           Set oParameter = .CreateParameter("@IdProductoInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdProductoInterconsulta = 0, Null, oTabla.IdProductoInterconsulta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProductoConsulta", adInteger, adParamInput, 0, IIf(oTabla.IdProductoConsulta = 0, Null, oTabla.IdProductoConsulta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidadCE", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdEspecialidadCE = .Parameters("@IdEspecialidadCE")
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
'   Descripción:    modifica un registro a la tabla EspecialidadCE
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEspecialidadCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EspecialidadCEModificar"
           Set oParameter = .CreateParameter("@IdProductoInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdProductoInterconsulta = 0, Null, oTabla.IdProductoInterconsulta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProductoConsulta", adInteger, adParamInput, 0, IIf(oTabla.IdProductoConsulta = 0, Null, oTabla.IdProductoConsulta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TiempoPromedioAtencion", adInteger, adParamInput, 0, IIf(oTabla.TiempoPromedioAtencion = 0, Null, oTabla.TiempoPromedioAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidadCE", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadCE = 0, Null, oTabla.IdEspecialidadCE)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla EspecialidadCE
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEspecialidadCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EspecialidadCEEliminar"
           Set oParameter = .CreateParameter("@IdEspecialidadCE", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidadCE = 0, Null, oTabla.IdEspecialidadCE)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla EspecialidadCE
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEspecialidadCE) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EspecialidadCESeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEspecialidadCE", adInteger, adParamInput, 0, oTabla.IdEspecialidadCE): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdProductoInterconsulta = IIf(IsNull(oRecordset!IdProductoInterconsulta), 0, oRecordset!IdProductoInterconsulta)
       oTabla.IdProductoConsulta = IIf(IsNull(oRecordset!IdProductoConsulta), 0, oRecordset!IdProductoConsulta)
       oTabla.TiempoPromedioAtencion = IIf(IsNull(oRecordset!TiempoPromedioAtencion), 0, oRecordset!TiempoPromedioAtencion)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.IdEspecialidadCE = IIf(IsNull(oRecordset!IdEspecialidadCE), 0, oRecordset!IdEspecialidadCE)
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


Function SeleccionarPorIdEspecialidad(ByVal oTabla As DOEspecialidadCE) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdEspecialidad = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EspecialidadCESeleccionarPorIdEspecialidad"
       Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, oTabla.IdEspecialidad): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorIdEspecialidad = True
       oTabla.IdProductoInterconsulta = IIf(IsNull(oRecordset!IdProductoInterconsulta), 0, oRecordset!IdProductoInterconsulta)
       oTabla.IdProductoConsulta = IIf(IsNull(oRecordset!IdProductoConsulta), 0, oRecordset!IdProductoConsulta)
       oTabla.TiempoPromedioAtencion = IIf(IsNull(oRecordset!TiempoPromedioAtencion), 0, oRecordset!TiempoPromedioAtencion)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.IdEspecialidadCE = IIf(IsNull(oRecordset!IdEspecialidadCE), 0, oRecordset!IdEspecialidadCE)
   Else
       SeleccionarPorIdEspecialidad = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


