VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enfermeria_TratamientoDosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Enfermeria_TratamientoDosis
'        Programado por: Cachay F
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
'   Descripción:    Inserta un registro a la tabla Enfermeria_TratamientoDosis
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOEnfermeria_TratamientoDosis) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_TratamientoDosisAgregar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, IIf(oTabla.IdVisita = 0, Null, oTabla.IdVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiaVisita", adInteger, adParamInput, 0, IIf(oTabla.IdDiaVisita = 0, Null, oTabla.IdDiaVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Dosis", adInteger, adParamInput, 0, IIf(oTabla.Dosis = 0, Null, oTabla.Dosis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DatoProrenata", adInteger, adParamInput, 0, IIf(oTabla.DatoProrenata = 0, Null, oTabla.DatoProrenata))
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
'   Descripción:    modifica un registro a la tabla Enfermeria_TratamientoDosis
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEnfermeria_TratamientoDosis) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_TratamientoDosisModificar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdVisita", adInteger, adParamInput, 0, IIf(oTabla.IdVisita = 0, Null, oTabla.IdVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiaVisita", adInteger, adParamInput, 0, IIf(oTabla.IdDiaVisita = 0, Null, oTabla.IdDiaVisita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Dosis", adInteger, adParamInput, 0, IIf(oTabla.Dosis = 0, Null, oTabla.Dosis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DatoProrenata", adInteger, adParamInput, 0, IIf(oTabla.DatoProrenata = 0, Null, oTabla.DatoProrenata))
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
'   Descripción:    elimia o borra un registro a la tabla Enfermeria_TratamientoDosis
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEnfermeria_TratamientoDosis) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Enfermeria_TratamientoDosisEliminar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla Enfermeria_TratamientoDosis
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEnfermeria_TratamientoDosis) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Enfermeria_TratamientoDosisSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdVisita = IIf(IsNull(oRecordset!IdVisita), 0, oRecordset!IdVisita)
       oTabla.IdDiaVisita = IIf(IsNull(oRecordset!IdDiaVisita), 0, oRecordset!IdDiaVisita)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.idItem = IIf(IsNull(oRecordset!idItem), 0, oRecordset!idItem)
       oTabla.Dosis = IIf(IsNull(oRecordset!Dosis), 0, oRecordset!Dosis)
       oTabla.DatoProrenata = IIf(IsNull(oRecordset!DatoProrenata), 0, oRecordset!DatoProrenata)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOEnfermeria_TratamientoDosis)
   oTabla.IdCuentaAtencion = 0
   oTabla.IdVisita = 0
   oTabla.IdDiaVisita = 0
   oTabla.IdReceta = 0
   oTabla.idItem = 0
   oTabla.Dosis = 0
   oTabla.DatoProrenata = 0
End Sub

