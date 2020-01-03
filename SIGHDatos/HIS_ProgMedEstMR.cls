VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_ProgMedEstMR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_ProgMedEstMR
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdHisProgMedEstMR As Long
Dim ml_IdMedico As Long
Dim ml_IdServicio As Long
Dim ml_IdEstablecimiento As Long
Dim mda_FechaProgramada As Date
Dim ml_IdTurno As Long

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
'   Descripción:    Inserta un registro a la tabla HIS_ProgMedEstMR
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_ProgMedEstMR) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ProgMedEstMRAgregar"
   Set oParameter = .CreateParameter("@IdHisProgMedEstMR", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisProgMedEstMR = .Parameters("@IdHisProgMedEstMR")
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
'   Descripción:    modifica un registro a la tabla HIS_ProgMedEstMR
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_ProgMedEstMR) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ProgMedEstMRModificar"
   Set oParameter = .CreateParameter("@IdHisProgMedEstMR", adInteger, adParamInput, 0, IIf(oTabla.IdHisProgMedEstMR = 0, Null, oTabla.IdHisProgMedEstMR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaProgramada", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaProgramada = 0, Null, oTabla.FechaProgramada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_ProgMedEstMR
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_ProgMedEstMR) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ProgMedEstMREliminar"
   Set oParameter = .CreateParameter("@IdHisProgMedEstMR", adInteger, adParamInput, 0, oTabla.IdHisProgMedEstMR)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_ProgMedEstMR
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_ProgMedEstMR) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_ProgMedEstMRSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisProgMedEstMR", adInteger, adParamInput, 0, oTabla.IdHisProgMedEstMR)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisProgMedEstMR = IIf(IsNull(oRecordset!IdHisProgMedEstMR), 0, oRecordset!IdHisProgMedEstMR)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.FechaProgramada = IIf(IsNull(oRecordset!FechaProgramada), 0, oRecordset!FechaProgramada)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_ProgMedEstMR)
   oTabla.IdHisProgMedEstMR = 0
   oTabla.IdMedico = 0
   oTabla.IdServicio = 0
   oTabla.IdEstablecimiento = 0
   oTabla.FechaProgramada = Now
   oTabla.IdTurno = 0
End Sub

Public Function ObtenerDatosProgramacionMedica(IdEstablecimiento As Long, IdServicio As Long, IdMedico As Long, Anio As Integer, IdMes As Integer, IdTurno As Integer) As Recordset
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
        .CommandText = "HIS_ProgMedEstMRxMedicoMesYanio"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IdServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMes", adInteger, adParamInput, 0, IdMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, Anio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosProgramacionMedica = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ListaServiciosPorEstablecimientoYEspecialidad(IdEspecialidad As Long, IdEstablecimiento As Long) As Recordset
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
        .CommandText = "HIS_ServEstablecimientoXespecialidadEstablecimiento"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IdEspecialidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaServiciosPorEstablecimientoYEspecialidad = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ListarProgramacionMedica_FechasMesActual(IdMedico As Long, ms_Fecha As String) As ADODB.Recordset
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
        .CommandText = "HIS_ProgMedEstMRporMedicoAnioMes"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ms_Fecha", adDBTimeStamp, adParamInput, 0, CDate(ms_Fecha)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListarProgramacionMedica_FechasMesActual = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ListarProgramacionMedicaPorMedicoYEstablecimiento(IdEstablecimiento As Long, IdMedico As Long, Mes As Integer, Anio As Integer) As Recordset
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
        .CommandText = "HIS_ProgMedEstMRporMedicoEstablecFecha"
        Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, Anio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, Mes): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListarProgramacionMedicaPorMedicoYEstablecimiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ValidarMedicoEstablecimiento(ml_IdMedico As Long, IdEstabelecimiento As Long) As Boolean
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
        .CommandText = "EMPLEADOSporIdMedicoIdEstablecimiento"
        Set oParameter = .CreateParameter("@ml_IdMedico", adInteger, adParamInput, 0, ml_IdMedico): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstabelecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount = 0 Then
        ValidarMedicoEstablecimiento = False
   Else
        ValidarMedicoEstablecimiento = True
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Frank
Function HIS_BuscaResponsableFiltro(lnIdEstablecimiento As Long, lnIdServicio As Long, LcAnio As String, Lcmes As String, lcNombre As String) As Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim sSql As String, sWhere As String
    Set HIS_BuscaResponsableFiltro = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_BuscaResponsableFiltro"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, lnIdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idservicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adVarChar, adParamInput, 4, LcAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 2, Lcmes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 200, lcNombre): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set HIS_BuscaResponsableFiltro = oRecordset
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function
