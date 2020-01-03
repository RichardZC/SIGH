VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtenInteGrupo
'        Programado por: Garay M
'        Fecha: Julio 2014
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
'   Descripción:    Inserta un registro a la tabla AtenInteGrupo
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenInteGrupo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoAgregar"
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AtencionIntegralGrupo", adVarChar, adParamInput, 50, IIf(oTabla.AtencionIntegralGrupo = "", Null, oTabla.AtencionIntegralGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeAnio", adInteger, adParamInput, 0, IIf(oTabla.DesdeAnio = 0, Null, oTabla.DesdeAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeMes", adInteger, adParamInput, 0, oTabla.DesdeMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeDia", adInteger, adParamInput, 0, oTabla.DesdeDia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaAnio", adInteger, adParamInput, 0, IIf(oTabla.HastaAnio = 0, Null, oTabla.HastaAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaMes", adInteger, adParamInput, 0, oTabla.HastaMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaDia", adInteger, adParamInput, 0, oTabla.HastaDia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdAtenInteGrupo = .Parameters("@IdAtenInteGrupo")
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
'   Descripción:    modifica un registro a la tabla AtenInteGrupo
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteGrupo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoModificar"
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AtencionIntegralGrupo", adVarChar, adParamInput, 50, IIf(oTabla.AtencionIntegralGrupo = "", Null, oTabla.AtencionIntegralGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeAnio", adInteger, adParamInput, 0, IIf(oTabla.DesdeAnio = 0, Null, oTabla.DesdeAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeMes", adInteger, adParamInput, 0, oTabla.DesdeMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DesdeDia", adInteger, adParamInput, 0, oTabla.DesdeDia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaAnio", adInteger, adParamInput, 0, IIf(oTabla.HastaAnio = 0, Null, oTabla.HastaAnio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaMes", adInteger, adParamInput, 0, oTabla.HastaMes)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HastaDia", adInteger, adParamInput, 0, oTabla.HastaDia)
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteGrupo
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteGrupo) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteGrupoEliminar"
   Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteGrupo
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteGrupo) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteGrupoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdAtenInteGrupo", adInteger, adParamInput, 0, oTabla.IdAtenInteGrupo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdAtenInteGrupo = oRecordset!IdAtenInteGrupo
       oTabla.AtencionIntegralGrupo = IIf(IsNull(oRecordset!AtencionIntegralGrupo), "", oRecordset!AtencionIntegralGrupo)
       oTabla.DesdeAnio = IIf(IsNull(oRecordset!DesdeAnio), 0, oRecordset!DesdeAnio)
       oTabla.DesdeMes = oRecordset!DesdeMes
       oTabla.DesdeDia = oRecordset!DesdeDia
       oTabla.HastaAnio = IIf(IsNull(oRecordset!HastaAnio), 0, oRecordset!HastaAnio)
       oTabla.HastaMes = oRecordset!HastaMes
       oTabla.HastaDia = oRecordset!HastaDia
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenInteGrupo)
   oTabla.IdAtenInteGrupo = 0
   oTabla.AtencionIntegralGrupo = ""
   oTabla.DesdeAnio = 0
   oTabla.DesdeMes = 0
   oTabla.DesdeDia = 0
   oTabla.HastaAnio = 0
   oTabla.HastaMes = 0
   oTabla.HastaDia = 0
End Sub
