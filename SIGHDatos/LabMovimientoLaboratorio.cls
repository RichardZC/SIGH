VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabMovimientoLaboratorio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabMovimientoLaboratorio
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla LabMovimientoLaboratorio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DoLabMovimientoLaboratorio) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Insertar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoLaboratorioAgregar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@CorrelativoAnual", adInteger, adParamInput, 0, IIf(oTabla.CorrelativoAnual = 0, Null, oTabla.CorrelativoAnual))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idPersonaTomaLab", adInteger, adParamInput, 0, IIf(oTabla.idPersonaTomaLab = 0, Null, oTabla.idPersonaTomaLab))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idPersonaRecoge", adInteger, adParamInput, 0, IIf(oTabla.idPersonaRecoge = 0, Null, oTabla.idPersonaRecoge))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@EsDiagnosticoDefinitivo", adInteger, adParamInput, 0, IIf(oTabla.EsDiagnosticoDefinitivo = 0, Null, oTabla.EsDiagnosticoDefinitivo))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@OrdenaPrueba", adVarChar, adParamInput, 50, IIf(oTabla.OrdenaPrueba = "", Null, oTabla.OrdenaPrueba))
    .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.paciente = "", Null, Left(oTabla.paciente, 100)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_FUM", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Eo_FUM = 0, Null, oTabla.Eo_FUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_EG", adInteger, adParamInput, 0, IIf(oTabla.Eo_EG = 0, Null, oTabla.Eo_EG))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@colegiatura", adVarChar, adParamInput, 10, IIf(oTabla.colegiatura = "", Null, oTabla.colegiatura))
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
'   Descripción:    Modifica un registro de la tabla LabMovimientoLaboratorio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Modificar(ByVal oTabla As DoLabMovimientoLaboratorio) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Modificar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoLaboratorioModificar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@CorrelativoAnual", adInteger, adParamInput, 0, IIf(oTabla.CorrelativoAnual = 0, Null, oTabla.CorrelativoAnual))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idPersonaTomaLab", adInteger, adParamInput, 0, IIf(oTabla.idPersonaTomaLab = 0, Null, oTabla.idPersonaTomaLab))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idPersonaRecoge", adInteger, adParamInput, 0, IIf(oTabla.idPersonaRecoge = 0, Null, oTabla.idPersonaRecoge))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@EsDiagnosticoDefinitivo", adInteger, adParamInput, 0, IIf(oTabla.EsDiagnosticoDefinitivo = 0, Null, oTabla.EsDiagnosticoDefinitivo))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@OrdenaPrueba", adVarChar, adParamInput, 50, IIf(oTabla.OrdenaPrueba = "", Null, oTabla.OrdenaPrueba))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.paciente = "", Null, Left(oTabla.paciente, 100)))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
    .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_FUM", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Eo_FUM = 0, Null, oTabla.Eo_FUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_EG", adInteger, adParamInput, 0, IIf(oTabla.Eo_EG = 0, Null, oTabla.Eo_EG))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@colegiatura", adVarChar, adParamInput, 10, IIf(oTabla.colegiatura = "", Null, oTabla.colegiatura))
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
'   Descripción:    Elimina o borra un registro de la tabla LabMovimientoLaboratorio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Eliminar(ByVal oTabla As DoLabMovimientoLaboratorio) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Eliminar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoLaboratorioEliminar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
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
'   Descripción:    Selecciona un único registro de la tabla LabMovimientoLaboratorio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DoLabMovimientoLaboratorio) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  SeleccionarPorId = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoLaboratorioSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
    .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  If Not (oRecordset.EOF And oRecordset.BOF) Then
    oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
    oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
    oTabla.CorrelativoAnual = IIf(IsNull(oRecordset!CorrelativoAnual), 0, oRecordset!CorrelativoAnual)
    oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
    oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
    oTabla.idPersonaTomaLab = IIf(IsNull(oRecordset!idPersonaTomaLab), 0, oRecordset!idPersonaTomaLab)
    oTabla.idPersonaRecoge = IIf(IsNull(oRecordset!idPersonaRecoge), 0, oRecordset!idPersonaRecoge)
    oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
    oTabla.EsDiagnosticoDefinitivo = IIf(IsNull(oRecordset!EsDiagnosticoDefinitivo), 0, oRecordset!EsDiagnosticoDefinitivo)
    oTabla.OrdenaPrueba = IIf(IsNull(oRecordset!OrdenaPrueba), "", oRecordset!OrdenaPrueba)
    oTabla.paciente = IIf(IsNull(oRecordset!paciente), "", oRecordset!paciente)
    oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
    oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
    oTabla.Eo_FUM = IIf(IsNull(oRecordset!Eo_FUM), 0, oRecordset!Eo_FUM)
    oTabla.Eo_EG = IIf(IsNull(oRecordset!Eo_EG), 0, oRecordset!Eo_EG)
    oTabla.colegiatura = IIf(IsNull(oRecordset!colegiatura), "", oRecordset!colegiatura)
    SeleccionarPorId = True
    ms_MensajeError = ""
  End If
  oRecordset.Close
  Exit Function

ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description
  Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DoLabMovimientoLaboratorio)
  oTabla.IdMovimiento = 0
  oTabla.IdOrden = 0
  oTabla.CorrelativoAnual = 0
  oTabla.IdCuentaAtencion = 0
  oTabla.IdComprobantePago = 0
  oTabla.idPersonaTomaLab = 0
  oTabla.idPersonaRecoge = 0
  oTabla.IdDiagnostico = 0
  oTabla.EsDiagnosticoDefinitivo = 0
  oTabla.OrdenaPrueba = ""
  oTabla.paciente = ""
  oTabla.IdTipoSexo = 0
  oTabla.FechaNacimiento = 0
End Sub

