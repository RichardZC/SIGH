VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProControlDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ProControlDato
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
'   Descripción:    Inserta un registro a la tabla ProControlDato
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOProControlDato) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlDatoAgregar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControlDato", adInteger, adParamInput, 0, IIf(oTabla.IdControlDato = 0, Null, oTabla.IdControlDato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ControlDato", adVarChar, adParamInput, 200, IIf(oTabla.ControlDato = "", Null, oTabla.ControlDato))
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
'   Descripción:    modifica un registro a la tabla ProControlDato
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProControlDato) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlDatoModificar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControlDato", adInteger, adParamInput, 0, IIf(oTabla.IdControlDato = 0, Null, oTabla.IdControlDato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ControlDato", adVarChar, adParamInput, 200, IIf(oTabla.ControlDato = "", Null, oTabla.ControlDato))
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
'   Descripción:    elimia o borra un registro a la tabla ProControlDato
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProControlDato) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProControlDatoEliminar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oTabla.IdProCabecera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oTabla.IdControl)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControlDato", adInteger, adParamInput, 0, oTabla.IdControlDato)
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
'   Descripción:    Selecciona un unico registro de la tabla ProControlDato
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProControlDato) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProControlDatoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
     .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, oTabla.IdProCabecera)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, oTabla.IdControl)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdControlDato", adInteger, adParamInput, 0, oTabla.IdControlDato)
    .Parameters.Append oParameter
     
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.ControlDato = IIf(IsNull(oRecordset!ControlDato), "", oRecordset!ControlDato)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOProControlDato)
   oTabla.IdPrograma = 0
   oTabla.IdProCabecera = 0
   oTabla.IdControl = 0
   oTabla.IdControlDato = 0
   oTabla.ControlDato = ""
End Sub

