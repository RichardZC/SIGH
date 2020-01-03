VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla SisFuaResumen
'        Programado por: Barrantes D
'        Fecha: Marzo 2014
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
'   Descripción:    Inserta un registro a la tabla SisFuaResumen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOSisFuaResumen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaResumenAgregar"
   Set oParameter = .CreateParameter("@idResumen", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 10, IIf(oTabla.Anio = "", Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 10, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroEnvio", adVarChar, adParamInput, 2, IIf(oTabla.NroEnvio = "", Null, oTabla.NroEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NomPaquete", adVarChar, adParamInput, 18, IIf(oTabla.NomPaquete = "", Null, oTabla.NomPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VersionGTI", adVarChar, adParamInput, 10, IIf(oTabla.VersionGTI = "", Null, oTabla.VersionGTI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilATE", adInteger, adParamInput, 0, IIf(oTabla.CantFilATE = 0, Null, oTabla.CantFilATE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilSMI", adInteger, adParamInput, 0, IIf(oTabla.CantFilSMI = 0, Null, oTabla.CantFilSMI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilDIA", adInteger, adParamInput, 0, IIf(oTabla.CantFilDIA = 0, Null, oTabla.CantFilDIA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilMED", adInteger, adParamInput, 0, IIf(oTabla.CantFilMED = 0, Null, oTabla.CantFilMED))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilINS", adInteger, adParamInput, 0, IIf(oTabla.CantFilINS = 0, Null, oTabla.CantFilINS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilPRO", adInteger, adParamInput, 0, IIf(oTabla.CantFilPRO = 0, Null, oTabla.CantFilPRO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilUSU", adInteger, adParamInput, 0, IIf(oTabla.CantFilUSU = 0, Null, oTabla.CantFilUSU))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idResumen = .Parameters("@idResumen")
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
'   Descripción:    modifica un registro a la tabla SisFuaResumen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOSisFuaResumen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaResumenModificar"
   Set oParameter = .CreateParameter("@idResumen", adInteger, adParamInput, 0, IIf(oTabla.idResumen = 0, Null, oTabla.idResumen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 10, IIf(oTabla.Anio = "", Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 10, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroEnvio", adVarChar, adParamInput, 2, IIf(oTabla.NroEnvio = "", Null, oTabla.NroEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NomPaquete", adVarChar, adParamInput, 18, IIf(oTabla.NomPaquete = "", Null, oTabla.NomPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VersionGTI", adVarChar, adParamInput, 10, IIf(oTabla.VersionGTI = "", Null, oTabla.VersionGTI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilATE", adInteger, adParamInput, 0, IIf(oTabla.CantFilATE = 0, Null, oTabla.CantFilATE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilSMI", adInteger, adParamInput, 0, IIf(oTabla.CantFilSMI = 0, Null, oTabla.CantFilSMI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilDIA", adInteger, adParamInput, 0, IIf(oTabla.CantFilDIA = 0, Null, oTabla.CantFilDIA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilMED", adInteger, adParamInput, 0, IIf(oTabla.CantFilMED = 0, Null, oTabla.CantFilMED))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilINS", adInteger, adParamInput, 0, IIf(oTabla.CantFilINS = 0, Null, oTabla.CantFilINS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilPRO", adInteger, adParamInput, 0, IIf(oTabla.CantFilPRO = 0, Null, oTabla.CantFilPRO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantFilUSU", adInteger, adParamInput, 0, IIf(oTabla.CantFilUSU = 0, Null, oTabla.CantFilUSU))
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
'   Descripción:    elimia o borra un registro a la tabla SisFuaResumen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOSisFuaResumen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaResumenEliminar"
   Set oParameter = .CreateParameter("@idResumen", adInteger, adParamInput, 0, oTabla.idResumen)
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
'   Descripción:    Selecciona un unico registro de la tabla SisFuaResumen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOSisFuaResumen) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaResumenSeleccionarPorId"
     Set oParameter = .CreateParameter("@idResumen", adInteger, adParamInput, 0, oTabla.idResumen)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idResumen = IIf(IsNull(oRecordset!idResumen), 0, oRecordset!idResumen)
       oTabla.Anio = IIf(IsNull(oRecordset!Anio), "", oRecordset!Anio)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), "", oRecordset!Mes)
       oTabla.NroEnvio = IIf(IsNull(oRecordset!NroEnvio), "", oRecordset!NroEnvio)
       oTabla.NomPaquete = IIf(IsNull(oRecordset!NomPaquete), "", oRecordset!NomPaquete)
       oTabla.VersionGTI = IIf(IsNull(oRecordset!VersionGTI), "", oRecordset!VersionGTI)
       oTabla.CantFilATE = IIf(IsNull(oRecordset!CantFilATE), 0, oRecordset!CantFilATE)
       oTabla.CantFilSMI = IIf(IsNull(oRecordset!CantFilSMI), 0, oRecordset!CantFilSMI)
       oTabla.CantFilDIA = IIf(IsNull(oRecordset!CantFilDIA), 0, oRecordset!CantFilDIA)
       oTabla.CantFilMED = IIf(IsNull(oRecordset!CantFilMED), 0, oRecordset!CantFilMED)
       oTabla.CantFilINS = IIf(IsNull(oRecordset!CantFilINS), 0, oRecordset!CantFilINS)
       oTabla.CantFilPRO = IIf(IsNull(oRecordset!CantFilPRO), 0, oRecordset!CantFilPRO)
       oTabla.CantFilUSU = IIf(IsNull(oRecordset!CantFilUSU), 0, oRecordset!CantFilUSU)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOSisFuaResumen)
   oTabla.idResumen = 0
   oTabla.Anio = ""
   oTabla.Mes = ""
   oTabla.NroEnvio = ""
   oTabla.NomPaquete = ""
   oTabla.VersionGTI = ""
   oTabla.CantFilATE = 0
   oTabla.CantFilSMI = 0
   oTabla.CantFilDIA = 0
   oTabla.CantFilMED = 0
   oTabla.CantFilINS = 0
   oTabla.CantFilPRO = 0
   oTabla.CantFilUSU = 0
End Sub


