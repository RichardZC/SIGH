VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaDevoluciones
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/01/2015 12:38:14
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
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
'   Descripción:    Inserta un registro a la tabla CajaDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOCajaDevoluciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CajaDevolucionesAgregar"
   Set oParameter = .CreateParameter("@idDevolucion", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@montoDevuelto", adCurrency, adParamInput, 0, oTabla.montoDevuelto)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@montoTotal", adCurrency, adParamInput, 0, oTabla.montoTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaDevolucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaDevolucion = 0, Null, oTabla.fechaDevolucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@motivo", adVarChar, adParamInput, 2000, IIf(oTabla.IdUsuario = 0, Null, oTabla.mMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idDevolucion = .Parameters("@idDevolucion")
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
'   Descripción:    modifica un registro a la tabla CajaDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaDevoluciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CajaDevolucionesModificar"
   Set oParameter = .CreateParameter("@idDevolucion", adInteger, adParamInput, 0, IIf(oTabla.idDevolucion = 0, Null, oTabla.idDevolucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@montoDevuelto", adCurrency, adParamInput, 0, oTabla.montoDevuelto)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@montoTotal", adCurrency, adParamInput, 0, oTabla.montoTotal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaDevolucion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaDevolucion = 0, Null, oTabla.fechaDevolucion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
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
'   Descripción:    elimia o borra un registro a la tabla CajaDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaDevoluciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "CajaDevolucionesEliminar"
   Set oParameter = .CreateParameter("@idcomprobantepago", adInteger, adParamInput, 0, oTabla.IdComprobantePago)
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
'   Descripción:    Selecciona un unico registro de la tabla CajaDevoluciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaDevoluciones) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CajaDevolucionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idDevolucion", adInteger, adParamInput, 0, oTabla.idDevolucion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idDevolucion = IIf(IsNull(oRecordset!idDevolucion), 0, oRecordset!idDevolucion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.montoDevuelto = oRecordset!montoDevuelto
       oTabla.montoTotal = oRecordset!montoTotal
       oTabla.fechaDevolucion = IIf(IsNull(oRecordset!fechaDevolucion), 0, oRecordset!fechaDevolucion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOCajaDevoluciones)
   oTabla.idDevolucion = 0
   oTabla.IdComprobantePago = 0
   oTabla.montoDevuelto = 0
   oTabla.montoTotal = 0
   oTabla.fechaDevolucion = Now
   oTabla.IdUsuario = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: CajaDevoluciones
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: CajaDevoluciones
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/01/2015 12:38:14
'------------------------------------------------------------------------------------
 


