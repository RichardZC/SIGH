VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_EstablecPacienteHIS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_EstablecPacienteHIS
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdEstablecPacienteHIS As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdHisPaciente As Long
Dim ms_NroHC_FF As String

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
'   Descripción:    Inserta un registro a la tabla HIS_EstablecPacienteHIS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_EstablecPacienteHIS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_EstablecPacienteHISAgregar"
   Set oParameter = .CreateParameter("@IdEstablecPacienteHIS", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdHisPaciente = 0, Null, oTabla.IdHisPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHC_FF", adVarChar, adParamInput, 20, IIf(oTabla.NroHC_FF = "", Null, oTabla.NroHC_FF))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdEstablecPacienteHIS = .Parameters("@IdEstablecPacienteHIS")
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
'   Descripción:    modifica un registro a la tabla HIS_EstablecPacienteHIS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_EstablecPacienteHIS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_EstablecPacienteHISModificar"
   Set oParameter = .CreateParameter("@IdEstablecPacienteHIS", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecPacienteHIS = 0, Null, oTabla.IdEstablecPacienteHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdHisPaciente = 0, Null, oTabla.IdHisPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHC_FF", adVarChar, adParamInput, 20, IIf(oTabla.NroHC_FF = "", Null, oTabla.NroHC_FF))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_EstablecPacienteHIS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_EstablecPacienteHIS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_EstablecPacienteHISEliminar"
   Set oParameter = .CreateParameter("@IdEstablecPacienteHIS", adInteger, adParamInput, 0, oTabla.IdEstablecPacienteHIS)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_EstablecPacienteHIS
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_EstablecPacienteHIS) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_EstablecPacienteHISSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdEstablecPacienteHIS", adInteger, adParamInput, 0, oTabla.IdEstablecPacienteHIS)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdEstablecPacienteHIS = IIf(IsNull(oRecordset!IdEstablecPacienteHIS), 0, oRecordset!IdEstablecPacienteHIS)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.IdHisPaciente = IIf(IsNull(oRecordset!IdHisPaciente), 0, oRecordset!IdHisPaciente)
       oTabla.NroHC_FF = IIf(IsNull(oRecordset!NroHC_FF), "", oRecordset!NroHC_FF)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_EstablecPacienteHIS)
   oTabla.IdEstablecPacienteHIS = 0
   oTabla.IdEstablecimiento = 0
   oTabla.IdHisPaciente = 0
   oTabla.NroHC_FF = ""
End Sub
