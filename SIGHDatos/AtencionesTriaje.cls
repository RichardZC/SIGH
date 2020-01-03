VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesTriaje
'        Autor: Daniel Enrique Barrantes
'        Fecha: 28/02/2011 02:52:13 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionesTriaje
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtencionTriaje) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesTriajeAgregar"
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Presion", adVarChar, adParamInput, 20, IIf(oTabla.Presion = "", Null, oTabla.Presion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Temperatura", adVarChar, adParamInput, 10, IIf(oTabla.Temperatura = "", Null, oTabla.Temperatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adVarChar, adParamInput, 10, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adVarChar, adParamInput, 10, IIf(oTabla.Talla = "", Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaTriaje", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaTriaje = 0, Null, oTabla.FechaTriaje))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioCreo", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCreo = 0, Null, oTabla.IdUsuarioCreo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModifico", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModifico = 0, Null, oTabla.FechaModifico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioModifico", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioModifico = 0, Null, oTabla.IdUsuarioModifico))
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
'   Descripción:    modifica un registro a la tabla AtencionesTriaje
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionTriaje) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesTriajeModificar"
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Presion", adVarChar, adParamInput, 20, IIf(oTabla.Presion = "", Null, oTabla.Presion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Temperatura", adVarChar, adParamInput, 10, IIf(oTabla.Temperatura = "", Null, oTabla.Temperatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adVarChar, adParamInput, 10, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adVarChar, adParamInput, 10, IIf(oTabla.Talla = "", Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaTriaje", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaTriaje = 0, Null, oTabla.FechaTriaje))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioCreo", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioCreo = 0, Null, oTabla.IdUsuarioCreo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaModifico", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaModifico = 0, Null, oTabla.FechaModifico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioModifico", adInteger, adParamInput, 0, IIf(oTabla.IdUsuarioModifico = 0, Null, oTabla.IdUsuarioModifico))
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesTriaje
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionTriaje) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesTriajeEliminar"
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesTriaje
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionTriaje) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesTriajeSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.Presion = IIf(IsNull(oRecordset!Presion), "", oRecordset!Presion)
       oTabla.Temperatura = IIf(IsNull(oRecordset!Temperatura), "", oRecordset!Temperatura)
       oTabla.Peso = IIf(IsNull(oRecordset!Peso), "", oRecordset!Peso)
       oTabla.Talla = IIf(IsNull(oRecordset!Talla), "", oRecordset!Talla)
       oTabla.FechaTriaje = IIf(IsNull(oRecordset!FechaTriaje), 0, oRecordset!FechaTriaje)
       oTabla.IdUsuarioCreo = IIf(IsNull(oRecordset!IdUsuarioCreo), 0, oRecordset!IdUsuarioCreo)
       oTabla.FechaModifico = IIf(IsNull(oRecordset!FechaModifico), 0, oRecordset!FechaModifico)
       oTabla.IdUsuarioModifico = IIf(IsNull(oRecordset!IdUsuarioModifico), 0, oRecordset!IdUsuarioModifico)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtencionTriaje)
   oTabla.idAtencion = 0
   oTabla.Presion = ""
   oTabla.Temperatura = ""
   oTabla.Peso = ""
   oTabla.Talla = ""
   oTabla.FechaTriaje = Now
   oTabla.IdUsuarioCreo = 0
   oTabla.FechaModifico = Now
   oTabla.IdUsuarioModifico = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: AtencionesTriaje
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: AtencionesTriaje
'        Autor:   DAniel Barrantes B.
'        Fecha: 28/02/2011 02:52:13 p.m.
'------------------------------------------------------------------------------------
 

