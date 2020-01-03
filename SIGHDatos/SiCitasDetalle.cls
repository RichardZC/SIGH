VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiCitasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiCitasDetalle
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/11/2018 09:19:28 a.m.
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
'   Descripción:    Inserta un registro a la tabla SiCitasDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOSiCitasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasDetalleAgregar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, IIf(oTabla.idCitaSI = 0, Null, oTabla.idCitaSI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.cantidad = 0, Null, oTabla.cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 200, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, IIf(oTabla.llaveTicket = "", Null, oTabla.llaveTicket))
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
'   Descripción:    modifica un registro a la tabla SiCitasDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOSiCitasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasDetalleModificar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, IIf(oTabla.idCitaSI = 0, Null, oTabla.idCitaSI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.cantidad = 0, Null, oTabla.cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 5, IIf(oTabla.Observaciones = "", Null, oTabla.Observaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@llaveTicket", adVarChar, adParamInput, 20, IIf(oTabla.llaveTicket = "", Null, oTabla.llaveTicket))
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
'   Descripción:    elimia o borra un registro a la tabla SiCitasDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOSiCitasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SiCitasDetalleEliminar"
   Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, oTabla.idCitaSI)
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
'   Descripción:    Selecciona un unico registro de la tabla SiCitasDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOSiCitasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SiCitasDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@idCitaSI", adInteger, adParamInput, 0, oTabla.idCitaSI)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idCitaSI = IIf(IsNull(oRecordset!idCitaSI), 0, oRecordset!idCitaSI)
       oTabla.idProductoCpt = IIf(IsNull(oRecordset!idProductoCpt), 0, oRecordset!idProductoCpt)
       oTabla.cantidad = IIf(IsNull(oRecordset!cantidad), 0, oRecordset!cantidad)
       oTabla.precio = IIf(IsNull(oRecordset!precio), 0, oRecordset!precio)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), 0, oRecordset!Observaciones)
       oTabla.llaveTicket = IIf(IsNull(oRecordset!llaveTicket), "", oRecordset!llaveTicket)
       
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOSiCitasDetalle)
   oTabla.idCitaSI = 0
   oTabla.idProductoCpt = 0
   oTabla.cantidad = 0
   oTabla.precio = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SiCitasDetalle
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiCitasDetalle
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/11/2018 09:19:28 a.m.
'------------------------------------------------------------------------------------
 

