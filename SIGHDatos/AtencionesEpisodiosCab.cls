VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesEpisodiosCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesEpisodiosCabecera
'        Programado por: Barrantes D
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
'   Descripción:    Inserta un registro a la tabla AtencionesEpisodiosCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtencionesEpisodiosCab) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesEpisodiosCabeceraAgregar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEpisodio", adInteger, adParamInput, 0, IIf(oTabla.idEpisodio = 0, Null, oTabla.idEpisodio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre))
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
'   Descripción:    modifica un registro a la tabla AtencionesEpisodiosCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionesEpisodiosCab) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesEpisodiosCabeceraModificar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEpisodio", adInteger, adParamInput, 0, IIf(oTabla.idEpisodio = 0, Null, oTabla.idEpisodio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre))
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesEpisodiosCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionesEpisodiosCab) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesEpisodiosCabeceraEliminar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEpisodio", adInteger, adParamInput, 0, oTabla.idEpisodio)
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesEpisodiosCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionesEpisodiosCab) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesEpisodiosCabeceraSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
     oRecordset.Filter = "idEpisodio=" & oTabla.idEpisodio
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idEpisodio = IIf(IsNull(oRecordset!idEpisodio), 0, oRecordset!idEpisodio)
       oTabla.FechaApertura = IIf(IsNull(oRecordset!FechaApertura), 0, oRecordset!FechaApertura)
       oTabla.FechaCierre = IIf(IsNull(oRecordset!FechaCierre), 0, oRecordset!FechaCierre)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtencionesEpisodiosCab)
   oTabla.IdPaciente = 0
   oTabla.idEpisodio = 0
   oTabla.FechaApertura = Now
   oTabla.FechaCierre = Now
End Sub

 

