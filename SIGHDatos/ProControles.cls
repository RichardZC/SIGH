VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ProControles
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

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
'   Descripción:    Inserta un registro a la tabla ProControles
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOProControles) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlesAgregar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaControl", adVarChar, adParamInput, 100, IIf(oTabla.FechaControl = "", Null, oTabla.FechaControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ControlOtroEESS", adBoolean, adParamInput, 0, IIf(oTabla.ControlOtroEESS = False, False, oTabla.ControlOtroEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
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
'   Descripción:    modifica un registro a la tabla ProControles
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProControles) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlesModificar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaControl", adVarChar, adParamInput, 100, IIf(oTabla.FechaControl = "", Null, oTabla.FechaControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ControlOtroEESS", adBoolean, adParamInput, 0, IIf(oTabla.ControlOtroEESS = False, False, oTabla.ControlOtroEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
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
'   Descripción:    elimia o borra un registro a la tabla ProControles
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProControles) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlesEliminar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oTabla.IdProCabecera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oTabla.IdControl)
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
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla ProControles
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function EliminarDatosControl(ByVal oTabla As DOProControles) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarDatosControl = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlDatoEliminarTodo"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oTabla.IdProCabecera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oTabla.IdControl)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarDatosControl = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla ProControles
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProControles) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProControlesSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oTabla.IdProCabecera)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oTabla.IdControl)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.FechaControl = IIf(IsNull(oRecordset!FechaControl), 0, oRecordset!FechaControl)
       oTabla.ControlOtroEESS = IIf(IsNull(oRecordset!ControlOtroEESS), False, oRecordset!ControlOtroEESS)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOProControles)
   oTabla.IdPrograma = 0
   oTabla.IdProCabecera = 0
   oTabla.IdControl = 0
   oTabla.idAtencion = 0
   oTabla.FechaControl = Now
   oTabla.ControlOtroEESS = False
   oTabla.IdEstablecimiento = 0
End Sub


