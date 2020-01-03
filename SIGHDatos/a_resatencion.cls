VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_resatencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla a_resatencion
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
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
'   Descripción:    Inserta un registro a la tabla a_resatencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As Doa_resatencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_resatencionAgregar"
   Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, IIf(oTabla.pers_IdResAtencion = "", Null, oTabla.pers_IdResAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, IIf(oTabla.pers_IdTipoDocumento = "", Null, oTabla.pers_IdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_ApePaterno", adVarChar, adParamInput, 40, IIf(oTabla.pers_ApePaterno = "", Null, oTabla.pers_ApePaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_ApeMaterno", adVarChar, adParamInput, 40, IIf(oTabla.pers_ApeMaterno = "", Null, oTabla.pers_ApeMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_PriNombre", adVarChar, adParamInput, 35, IIf(oTabla.pers_PriNombre = "", Null, oTabla.pers_PriNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_OtrNombre", adVarChar, adParamInput, 35, IIf(oTabla.pers_OtrNombre = "", Null, oTabla.pers_OtrNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdTipoPersonalSalud", adVarChar, adParamInput, 2, IIf(oTabla.pers_IdTipoPersonalSalud = "", Null, oTabla.pers_IdTipoPersonalSalud))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_Colegiatura", adVarChar, adParamInput, 20, IIf(oTabla.pers_Colegiatura = "", Null, oTabla.pers_Colegiatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdEspecialidad", adVarChar, adParamInput, 2, IIf(Val(oTabla.pers_IdEspecialidad) = 0, Null, oTabla.pers_IdEspecialidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_NroEspecialidad", adVarChar, adParamInput, 10, IIf(oTabla.pers_NroEspecialidad = "", Null, oTabla.pers_NroEspecialidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.pers_IdEstado = "", Null, oTabla.pers_IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   If Err.Number = -2147217873 Then
      Modificar oTabla
   Else
      ms_MensajeError = Err.Number & " " + Err.Description
   End If
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla a_resatencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Doa_resatencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_resatencionModificar"
   Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, IIf(oTabla.pers_IdResAtencion = "", Null, oTabla.pers_IdResAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, IIf(oTabla.pers_IdTipoDocumento = "", Null, oTabla.pers_IdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_ApePaterno", adVarChar, adParamInput, 40, IIf(oTabla.pers_ApePaterno = "", Null, oTabla.pers_ApePaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_ApeMaterno", adVarChar, adParamInput, 40, IIf(oTabla.pers_ApeMaterno = "", Null, oTabla.pers_ApeMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_PriNombre", adVarChar, adParamInput, 35, IIf(oTabla.pers_PriNombre = "", Null, oTabla.pers_PriNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_OtrNombre", adVarChar, adParamInput, 35, IIf(oTabla.pers_OtrNombre = "", Null, oTabla.pers_OtrNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdTipoPersonalSalud", adVarChar, adParamInput, 2, IIf(oTabla.pers_IdTipoPersonalSalud = "", Null, oTabla.pers_IdTipoPersonalSalud))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_Colegiatura", adVarChar, adParamInput, 20, IIf(oTabla.pers_Colegiatura = "", Null, oTabla.pers_Colegiatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdEspecialidad", adVarChar, adParamInput, 2, IIf(Val(oTabla.pers_IdEspecialidad) = 0, Null, oTabla.pers_IdEspecialidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_NroEspecialidad", adVarChar, adParamInput, 10, IIf(oTabla.pers_NroEspecialidad = "", Null, oTabla.pers_NroEspecialidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.pers_IdEstado = "", Null, oTabla.pers_IdEstado))
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
'   Descripción:    elimia o borra un registro a la tabla a_resatencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Doa_resatencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_resatencionEliminar"
   Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, oTabla.pers_IdResAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, oTabla.pers_IdTipoDocumento)
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
'   Descripción:    Selecciona un unico registro de la tabla a_resatencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Doa_resatencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "a_resatencionSeleccionarPorId"
     Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, oTabla.pers_IdResAtencion)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, oTabla.pers_IdTipoDocumento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.pers_IdResAtencion = IIf(IsNull(oRecordset!pers_IdResAtencion), "", oRecordset!pers_IdResAtencion)
       oTabla.pers_IdTipoDocumento = IIf(IsNull(oRecordset!pers_IdTipoDocumento), "", oRecordset!pers_IdTipoDocumento)
       oTabla.pers_ApePaterno = IIf(IsNull(oRecordset!pers_ApePaterno), "", oRecordset!pers_ApePaterno)
       oTabla.pers_ApeMaterno = IIf(IsNull(oRecordset!pers_ApeMaterno), "", oRecordset!pers_ApeMaterno)
       oTabla.pers_PriNombre = IIf(IsNull(oRecordset!pers_PriNombre), "", oRecordset!pers_PriNombre)
       oTabla.pers_OtrNombre = IIf(IsNull(oRecordset!pers_OtrNombre), "", oRecordset!pers_OtrNombre)
       oTabla.pers_IdTipoPersonalSalud = IIf(IsNull(oRecordset!pers_IdTipoPersonalSalud), "", oRecordset!pers_IdTipoPersonalSalud)
       oTabla.pers_Colegiatura = IIf(IsNull(oRecordset!pers_Colegiatura), "", oRecordset!pers_Colegiatura)
       oTabla.pers_IdEspecialidad = IIf(IsNull(oRecordset!pers_IdEspecialidad), "", oRecordset!pers_IdEspecialidad)
       oTabla.pers_NroEspecialidad = IIf(IsNull(oRecordset!pers_NroEspecialidad), "", oRecordset!pers_NroEspecialidad)
       oTabla.pers_IdEstado = IIf(IsNull(oRecordset!pers_IdEstado), "", oRecordset!pers_IdEstado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Doa_resatencion)
   oTabla.pers_IdResAtencion = ""
   oTabla.pers_IdTipoDocumento = ""
   oTabla.pers_ApePaterno = ""
   oTabla.pers_ApeMaterno = ""
   oTabla.pers_PriNombre = ""
   oTabla.pers_OtrNombre = ""
   oTabla.pers_IdTipoPersonalSalud = ""
   oTabla.pers_Colegiatura = ""
   oTabla.pers_IdEspecialidad = ""
   oTabla.pers_NroEspecialidad = ""
   oTabla.pers_IdEstado = ""
End Sub


