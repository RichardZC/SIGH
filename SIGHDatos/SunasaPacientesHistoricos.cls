VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SunasaPacientesHistoricos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla SunasaPacientesHistoricos
'        Programado por: Barrantes D
'        Fecha: Marzo 2011
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
'   Descripción:    Inserta un registro a la tabla SunasaPacientesHistoricos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSunasaPacientesHistoricos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SunasaPacientesHistoricosAgregar"
   Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoIAFA", adVarChar, adParamInput, 5, IIf(oTabla.CodigoIAFA = "", Null, oTabla.CodigoIAFA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaisTitular", adInteger, adParamInput, 0, IIf(oTabla.idPaisTitular = 0, Null, oTabla.idPaisTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoDocumentoTitular", adInteger, adParamInput, 0, IIf(oTabla.idTipoDocumentoTitular = 0, Null, oTabla.idTipoDocumentoTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoTitular", adVarChar, adParamInput, 10, IIf(oTabla.NroDocumentoTitular = "", Null, oTabla.NroDocumentoTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoCasada", adVarChar, adParamInput, 35, IIf(oTabla.ApellidoCasada = "", Null, oTabla.ApellidoCasada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValidacionRegIdentidad", adBoolean, adParamInput, 0, IIf(oTabla.ValidacionRegIdentidad = 0, Null, oTabla.ValidacionRegIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroCarnetIdentidad", adVarChar, adParamInput, 10, IIf(oTabla.NroCarnetIdentidad = "", Null, oTabla.NroCarnetIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstadoDelSeguro", adInteger, adParamInput, 0, IIf(oTabla.EstadoDelSeguro = 0, Null, oTabla.EstadoDelSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAfiliacion", adInteger, adParamInput, 0, IIf(oTabla.IdAfiliacion = 0, Null, oTabla.IdAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ProductoYplan", adVarChar, adParamInput, 7, IIf(oTabla.ProductoYplan = "", Null, oTabla.ProductoYplan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaInicioAfiliacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaInicioAfiliacion = 0, Null, oTabla.FechaInicioAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaFinalAfiliacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaFinalAfiliacion = 0, Null, oTabla.FechaFinalAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idRegimen", adInteger, adParamInput, 0, IIf(oTabla.idRegimen = 0, Null, oTabla.idRegimen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimientoIAFA", adVarChar, adParamInput, 8, IIf(oTabla.CodigoEstablecimientoIAFA = "", Null, oTabla.CodigoEstablecimientoIAFA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimientoRENAES", adVarChar, adParamInput, 8, IIf(oTabla.CodigoEstablecimientoRENAES = "", Null, oTabla.CodigoEstablecimientoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idParentesco", adInteger, adParamInput, 0, IIf(oTabla.idParentesco = 0, Null, oTabla.idParentesco))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RUCempleador", adVarChar, adParamInput, 11, IIf(oTabla.RUCempleador = "", Null, oTabla.RUCempleador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AnteriorIdTipoDocumentoAsegurado", adInteger, adParamInput, 0, IIf(oTabla.AnteriorIdTipoDocumentoAsegurado = 0, Null, oTabla.AnteriorIdTipoDocumentoAsegurado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AnteriorNroDocumentoAsegurado", adVarChar, adParamInput, 10, IIf(oTabla.AnteriorNroDocumentoAsegurado = "", Null, oTabla.AnteriorNroDocumentoAsegurado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNIusarioOperacion", adVarChar, adParamInput, 8, IIf(oTabla.DNIusarioOperacion = "", Null, oTabla.DNIusarioOperacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idOperacion", adInteger, adParamInput, 0, IIf(oTabla.idOperacion = 0, Null, oTabla.idOperacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEnvio", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEnvio = 0, Null, oTabla.FechaEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioParienteEncargado", adVarChar, adParamInput, 100, IIf(oTabla.SisSepelioParienteEncargado = "", Null, oTabla.SisSepelioParienteEncargado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioDni", adVarChar, adParamInput, 8, IIf(oTabla.SisSepelioDni = "", Null, oTabla.SisSepelioDni))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioFnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.SisSepelioFnacimiento = 0, Null, oTabla.SisSepelioFnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioSexo", adInteger, adParamInput, 0, IIf(oTabla.SisSepelioSexo = 0, Null, oTabla.SisSepelioSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisNroAfiliacion", adVarChar, adParamInput, 20, IIf(oTabla.SisNroAfiliacion = "", Null, oTabla.SisNroAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@YaNoTieneSeguro", adBoolean, adParamInput, 0, oTabla.yaNoTieneSeguro)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idSunasaPacienteHistorico = .Parameters("@idSunasaPacienteHistorico")
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
'   Descripción:    modifica un registro a la tabla SunasaPacientesHistoricos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSunasaPacientesHistoricos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SunasaPacientesHistoricosModificar"
   Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamInput, 0, IIf(oTabla.idSunasaPacienteHistorico = 0, Null, oTabla.idSunasaPacienteHistorico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoIAFA", adVarChar, adParamInput, 5, IIf(oTabla.CodigoIAFA = "", Null, oTabla.CodigoIAFA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaisTitular", adInteger, adParamInput, 0, IIf(oTabla.idPaisTitular = 0, Null, oTabla.idPaisTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoDocumentoTitular", adInteger, adParamInput, 0, IIf(oTabla.idTipoDocumentoTitular = 0, Null, oTabla.idTipoDocumentoTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocumentoTitular", adVarChar, adParamInput, 10, IIf(oTabla.NroDocumentoTitular = "", Null, oTabla.NroDocumentoTitular))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoCasada", adVarChar, adParamInput, 35, IIf(oTabla.ApellidoCasada = "", Null, oTabla.ApellidoCasada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValidacionRegIdentidad", adBoolean, adParamInput, 0, IIf(oTabla.ValidacionRegIdentidad = 0, Null, oTabla.ValidacionRegIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroCarnetIdentidad", adVarChar, adParamInput, 10, IIf(oTabla.NroCarnetIdentidad = "", Null, oTabla.NroCarnetIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstadoDelSeguro", adInteger, adParamInput, 0, IIf(oTabla.EstadoDelSeguro = 0, Null, oTabla.EstadoDelSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdAfiliacion", adInteger, adParamInput, 0, IIf(oTabla.IdAfiliacion = 0, Null, oTabla.IdAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ProductoYplan", adVarChar, adParamInput, 7, IIf(oTabla.ProductoYplan = "", Null, oTabla.ProductoYplan))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaInicioAfiliacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaInicioAfiliacion = 0, Null, oTabla.FechaInicioAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaFinalAfiliacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaFinalAfiliacion = 0, Null, oTabla.FechaFinalAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idRegimen", adInteger, adParamInput, 0, IIf(oTabla.idRegimen = 0, Null, oTabla.idRegimen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimientoIAFA", adVarChar, adParamInput, 8, IIf(oTabla.CodigoEstablecimientoIAFA = "", Null, oTabla.CodigoEstablecimientoIAFA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimientoRENAES", adVarChar, adParamInput, 8, IIf(oTabla.CodigoEstablecimientoRENAES = "", Null, oTabla.CodigoEstablecimientoRENAES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idParentesco", adInteger, adParamInput, 0, IIf(oTabla.idParentesco = 0, Null, oTabla.idParentesco))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RUCempleador", adVarChar, adParamInput, 11, IIf(oTabla.RUCempleador = "", Null, oTabla.RUCempleador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AnteriorIdTipoDocumentoAsegurado", adInteger, adParamInput, 0, IIf(oTabla.AnteriorIdTipoDocumentoAsegurado = 0, Null, oTabla.AnteriorIdTipoDocumentoAsegurado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AnteriorNroDocumentoAsegurado", adVarChar, adParamInput, 10, IIf(oTabla.AnteriorNroDocumentoAsegurado = "", Null, oTabla.AnteriorNroDocumentoAsegurado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNIusarioOperacion", adVarChar, adParamInput, 8, IIf(oTabla.DNIusarioOperacion = "", Null, oTabla.DNIusarioOperacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idOperacion", adInteger, adParamInput, 0, IIf(oTabla.idOperacion = 0, Null, oTabla.idOperacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEnvio", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEnvio = 0, Null, oTabla.FechaEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioParienteEncargado", adVarChar, adParamInput, 100, IIf(oTabla.SisSepelioParienteEncargado = "", Null, oTabla.SisSepelioParienteEncargado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioDni", adVarChar, adParamInput, 8, IIf(oTabla.SisSepelioDni = "", Null, oTabla.SisSepelioDni))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioFnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.SisSepelioFnacimiento = 0, Null, oTabla.SisSepelioFnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisSepelioSexo", adInteger, adParamInput, 0, IIf(oTabla.SisSepelioSexo = 0, Null, oTabla.SisSepelioSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SisNroAfiliacion", adVarChar, adParamInput, 20, IIf(oTabla.SisNroAfiliacion = "", Null, oTabla.SisNroAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@YaNoTieneSeguro", adBoolean, adParamInput, 0, oTabla.yaNoTieneSeguro)
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
'   Descripción:    elimia o borra un registro a la tabla SunasaPacientesHistoricos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSunasaPacientesHistoricos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SunasaPacientesHistoricosEliminar"
   Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamInput, 0, oTabla.idSunasaPacienteHistorico)
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
'   Descripción:    Selecciona un unico registro de la tabla SunasaPacientesHistoricos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSunasaPacientesHistoricos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SunasaPacientesHistoricosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idSunasaPacienteHistorico", adInteger, adParamInput, 0, oTabla.idSunasaPacienteHistorico)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       LlenaTabla oTabla, oRecordset
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSunasaPacientesHistoricos)
   oTabla.idSunasaPacienteHistorico = 0
   oTabla.IdPaciente = 0
   oTabla.CodigoIAFA = ""
   oTabla.idPaisTitular = 0
   oTabla.idTipoDocumentoTitular = 0
   oTabla.NroDocumentoTitular = ""
   oTabla.ApellidoCasada = ""
   oTabla.ValidacionRegIdentidad = False
   oTabla.NroCarnetIdentidad = ""
   oTabla.EstadoDelSeguro = 0
   oTabla.IdAfiliacion = 0
   oTabla.ProductoYplan = ""
   oTabla.FechaInicioAfiliacion = Now
   oTabla.FechaFinalAfiliacion = Now
   oTabla.idRegimen = 0
   oTabla.CodigoEstablecimientoIAFA = ""
   oTabla.CodigoEstablecimientoRENAES = ""
   oTabla.idParentesco = 0
   oTabla.RUCempleador = ""
   oTabla.AnteriorIdTipoDocumentoAsegurado = 0
   oTabla.AnteriorNroDocumentoAsegurado = ""
   oTabla.DNIusarioOperacion = ""
   oTabla.idOperacion = 0
   oTabla.FechaEnvio = Now
   oTabla.SisSepelioParienteEncargado = ""
   oTabla.SisSepelioDni = ""
   oTabla.SisSepelioFnacimiento = Now
   oTabla.SisSepelioSexo = 0
   oTabla.SisNroAfiliacion = ""
End Sub

Function SeleccionarPorIdPaciente(ByVal oTabla As DoSunasaPacientesHistoricos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdPaciente = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SunasaPacientesHistoricosSeleccionarPorIdPaciente"
     Set oParameter = .CreateParameter("@idPaciente ", adInteger, adParamInput, 0, oTabla.IdPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       LlenaTabla oTabla, oRecordset
   SeleccionarPorIdPaciente = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub LlenaTabla(ByVal oTabla As DoSunasaPacientesHistoricos, oRecordset As Recordset)
       oTabla.idSunasaPacienteHistorico = IIf(IsNull(oRecordset!idSunasaPacienteHistorico), 0, oRecordset!idSunasaPacienteHistorico)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.CodigoIAFA = IIf(IsNull(oRecordset!CodigoIAFA), "", oRecordset!CodigoIAFA)
       oTabla.idPaisTitular = IIf(IsNull(oRecordset!idPaisTitular), 0, oRecordset!idPaisTitular)
       oTabla.idTipoDocumentoTitular = IIf(IsNull(oRecordset!idTipoDocumentoTitular), 0, oRecordset!idTipoDocumentoTitular)
       oTabla.NroDocumentoTitular = IIf(IsNull(oRecordset!NroDocumentoTitular), "", oRecordset!NroDocumentoTitular)
       oTabla.ApellidoCasada = IIf(IsNull(oRecordset!ApellidoCasada), "", oRecordset!ApellidoCasada)
       oTabla.ValidacionRegIdentidad = IIf(IsNull(oRecordset!ValidacionRegIdentidad), 0, oRecordset!ValidacionRegIdentidad)
       oTabla.NroCarnetIdentidad = IIf(IsNull(oRecordset!NroCarnetIdentidad), "", oRecordset!NroCarnetIdentidad)
       oTabla.EstadoDelSeguro = IIf(IsNull(oRecordset!EstadoDelSeguro), 0, oRecordset!EstadoDelSeguro)
       oTabla.IdAfiliacion = IIf(IsNull(oRecordset!IdAfiliacion), 0, oRecordset!IdAfiliacion)
       oTabla.ProductoYplan = IIf(IsNull(oRecordset!ProductoYplan), "", oRecordset!ProductoYplan)
       oTabla.FechaInicioAfiliacion = IIf(IsNull(oRecordset!FechaInicioAfiliacion), 0, oRecordset!FechaInicioAfiliacion)
       oTabla.FechaFinalAfiliacion = IIf(IsNull(oRecordset!FechaFinalAfiliacion), 0, oRecordset!FechaFinalAfiliacion)
       oTabla.idRegimen = IIf(IsNull(oRecordset!idRegimen), 0, oRecordset!idRegimen)
       oTabla.CodigoEstablecimientoIAFA = IIf(IsNull(oRecordset!CodigoEstablecimientoIAFA), "", oRecordset!CodigoEstablecimientoIAFA)
       oTabla.CodigoEstablecimientoRENAES = IIf(IsNull(oRecordset!CodigoEstablecimientoRENAES), "", oRecordset!CodigoEstablecimientoRENAES)
       oTabla.idParentesco = IIf(IsNull(oRecordset!idParentesco), 0, oRecordset!idParentesco)
       oTabla.RUCempleador = IIf(IsNull(oRecordset!RUCempleador), "", oRecordset!RUCempleador)
       oTabla.AnteriorIdTipoDocumentoAsegurado = IIf(IsNull(oRecordset!AnteriorIdTipoDocumentoAsegurado), 0, oRecordset!AnteriorIdTipoDocumentoAsegurado)
       oTabla.AnteriorNroDocumentoAsegurado = IIf(IsNull(oRecordset!AnteriorNroDocumentoAsegurado), "", oRecordset!AnteriorNroDocumentoAsegurado)
       oTabla.DNIusarioOperacion = IIf(IsNull(oRecordset!DNIusarioOperacion), "", oRecordset!DNIusarioOperacion)
       oTabla.idOperacion = IIf(IsNull(oRecordset!idOperacion), 0, oRecordset!idOperacion)
       oTabla.FechaEnvio = IIf(IsNull(oRecordset!FechaEnvio), 0, oRecordset!FechaEnvio)
       oTabla.SisSepelioParienteEncargado = IIf(IsNull(oRecordset!SisSepelioParienteEncargado), "", oRecordset!SisSepelioParienteEncargado)
       oTabla.SisSepelioDni = IIf(IsNull(oRecordset!SisSepelioDni), "", oRecordset!SisSepelioDni)
       oTabla.SisSepelioFnacimiento = IIf(IsNull(oRecordset!SisSepelioFnacimiento), 0, oRecordset!SisSepelioFnacimiento)
       oTabla.SisSepelioSexo = IIf(IsNull(oRecordset!SisSepelioSexo), 0, oRecordset!SisSepelioSexo)
       oTabla.SisNroAfiliacion = IIf(IsNull(oRecordset!SisNroAfiliacion), "", oRecordset!SisNroAfiliacion)
       oTabla.yaNoTieneSeguro = oRecordset!yaNoTieneSeguro
End Sub

