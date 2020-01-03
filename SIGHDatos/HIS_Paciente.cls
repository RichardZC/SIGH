VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_Paciente"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_Paciente
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_Paciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_Paciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_PacienteAgregar"
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sexo", adInteger, adParamInput, 0, IIf(oTabla.Sexo = 0, Null, oTabla.Sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdNacionalidad", adInteger, adParamInput, 0, IIf(oTabla.IdNacionalidad = 0, Null, oTabla.IdNacionalidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocIdentidad", adVarChar, adParamInput, 12, IIf(oTabla.NroDocIdentidad = "", Null, oTabla.NroDocIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHijo", adChar, adParamInput, 2, IIf(oTabla.NroHijo = "", Null, oTabla.NroHijo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEtnia", adChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPacienteGalenHos", adInteger, adParamInput, 0, IIf(oTabla.IdPacienteGalenHos = 0, Null, oTabla.IdPacienteGalenHos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisPaciente = .Parameters("@IdHisPaciente")
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
'   Descripción:    modifica un registro a la tabla HIS_Paciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_Paciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_PacienteModificar"
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdHisPaciente = 0, Null, oTabla.IdHisPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sexo", adInteger, adParamInput, 0, IIf(oTabla.Sexo = 0, Null, oTabla.Sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdNacionalidad", adInteger, adParamInput, 0, IIf(oTabla.IdNacionalidad = 0, Null, oTabla.IdNacionalidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocIdentidad", adVarChar, adParamInput, 12, IIf(oTabla.NroDocIdentidad = "", Null, oTabla.NroDocIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHijo", adChar, adParamInput, 2, IIf(oTabla.NroHijo = "", Null, oTabla.NroHijo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEtnia", adChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPacienteGalenHos", adInteger, adParamInput, 0, IIf(oTabla.IdPacienteGalenHos = 0, Null, oTabla.IdPacienteGalenHos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_Paciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_Paciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_PacienteEliminar"
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, oTabla.IdHisPaciente)
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

Sub ActualzaTabla(ByRef oTabla As DOHIS_Paciente, oRecordset As Recordset)
       oTabla.IdHisPaciente = IIf(IsNull(oRecordset!IdHisPaciente), 0, oRecordset!IdHisPaciente)
       oTabla.NroHC_FF = ""
       oTabla.Sexo = IIf(IsNull(oRecordset!Sexo), 0, oRecordset!Sexo)
       oTabla.IdNacionalidad = IIf(IsNull(oRecordset!IdNacionalidad), 0, oRecordset!IdNacionalidad)
       oTabla.NroDocIdentidad = IIf(IsNull(oRecordset!NroDocIdentidad), "", oRecordset!NroDocIdentidad)
       oTabla.NroHijo = IIf(IsNull(oRecordset!NroHijo), "", oRecordset!NroHijo)
       oTabla.IdEtnia = IIf(IsNull(oRecordset!IdEtnia), "", oRecordset!IdEtnia)
       oTabla.IdPacienteGalenHos = IIf(IsNull(oRecordset!IdPacienteGalenHos), 0, oRecordset!IdPacienteGalenHos)
       oTabla.IdTipoDocumento = IIf(IsNull(oRecordset!IdTipoDocumento), 0, oRecordset!IdTipoDocumento)

End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_Paciente
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_Paciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_PacienteSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, oTabla.IdHisPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
      ActualzaTabla oTabla, oRecordset
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_Paciente)
   oTabla.IdHisPaciente = 0
   'oTabla.NroHC_FF = ""
   'oTabla.TipoHistoria = 0
   oTabla.Sexo = 0
   oTabla.IdNacionalidad = 0
   oTabla.NroDocIdentidad = ""
   oTabla.NroHijo = ""
   oTabla.IdEtnia = ""
   oTabla.IdPacienteGalenHos = 0
   oTabla.IdTipoDocumento = 0
End Sub

Function VerificarPaciente(NroDocumentoIdentidad As String, IdTipoDocIdent As Integer) As DOHIS_Paciente
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_pacienteXdocumento"
        Set oParameter = .CreateParameter("@NroDocumentoIdentidad", adVarChar, adParamInput, 12, NroDocumentoIdentidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, IdTipoDocIdent): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        ActualzaTabla mo_TablaHISPaciente, oRecordset
        Set VerificarPaciente = mo_TablaHISPaciente
   Else
        mo_TablaHISPaciente.IdHisPaciente = 0
        Set VerificarPaciente = mo_TablaHISPaciente
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function VerificarPacienteNroHijo(NroDocumentoIdentidad As String, IdTipoDocIdent As Integer, NroHijo As Integer) As DOHIS_Paciente
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_pacienteXdocumentoNroHijo"
        Set oParameter = .CreateParameter("@NroDocumentoIdentidad", adVarChar, adParamInput, 12, NroDocumentoIdentidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, IdTipoDocIdent): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroHijo", adInteger, adParamInput, 0, NroHijo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        ActualzaTabla mo_TablaHISPaciente, oRecordset
        Set VerificarPacienteNroHijo = mo_TablaHISPaciente
   Else
        mo_TablaHISPaciente.IdHisPaciente = 0
        Set VerificarPacienteNroHijo = mo_TablaHISPaciente
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Actualiza todos los registro de los pacientes de HIS a el ID del Paciente GalenHos Correspondiente
Public Function ActualizarPacienteHIS(NroDocumentoIdentidad As String, IdPacienteGalenHos As Long) As Boolean
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente
    sSql = ""
    sWhere = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_PACIENTEactualizaIdPacienteGalenHos"
        Set oParameter = .CreateParameter("@IdPacienteGalenHos", adInteger, adParamInput, 0, IdPacienteGalenHos): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumentoIdentidad", adVarChar, adParamInput, 12, NroDocumentoIdentidad): .Parameters.Append oParameter
         .Execute
   End With
   ActualizarPacienteHIS = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function HISPacientesFiltraPorNroDocumentoYtipo(lcNroDocumento As String, lnIdDocIdentidad As Integer) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
     Set HISPacientesFiltraPorNroDocumentoYtipo = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
        .CommandText = "HIS_PacientesFiltraPorNroDocumentoYtipo"
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, lcNroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, lnIdDocIdentidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set HISPacientesFiltraPorNroDocumentoYtipo = oRecordset
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

