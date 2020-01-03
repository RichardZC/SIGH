VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_historicoAten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla his_historicoAtenciones
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
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
'   Descripción:    Inserta un registro a la tabla his_historicoAtenciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOhis_historicoAten) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "his_historicoAtencionesAgregar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@diagnost", adVarChar, adParamInput, 6, IIf(oTabla.diagnost = "", Null, oTabla.diagnost))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 6, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ups", adVarChar, adParamInput, 6, IIf(oTabla.ups = "", Null, oTabla.ups))
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
'   Descripción:    modifica un registro a la tabla his_historicoAtenciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOhis_historicoAten) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "his_historicoAtencionesModificar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@diagnost", adVarChar, adParamInput, 6, IIf(oTabla.diagnost = "", Null, oTabla.diagnost))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 6, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ups", adVarChar, adParamInput, 6, IIf(oTabla.ups = "", Null, oTabla.ups))
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
'   Descripción:    elimia o borra un registro a la tabla his_historicoAtenciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOhis_historicoAten) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "his_historicoAtencionesEliminar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla his_historicoAtenciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOhis_historicoAten) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "his_historicoAtencionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.fecha = IIf(IsNull(oRecordset!fecha), 0, oRecordset!fecha)
       oTabla.diagnost = IIf(IsNull(oRecordset!diagnost), "", oRecordset!diagnost)
       oTabla.cpt = IIf(IsNull(oRecordset!cpt), "", oRecordset!cpt)
       oTabla.ups = IIf(IsNull(oRecordset!ups), "", oRecordset!ups)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOhis_historicoAten)
   oTabla.IdPaciente = 0
   oTabla.fecha = Now
   oTabla.diagnost = ""
   oTabla.cpt = ""
   oTabla.ups = ""
End Sub

