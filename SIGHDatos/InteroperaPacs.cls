VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InteroperaPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InteroperaPacs
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 05:52:12 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla InteroperaPacs
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInteroperaPacs) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaPacsAgregar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, IIf(oTabla.cptCodigo = "", Null, oTabla.cptCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 250, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteId", adInteger, adParamInput, 0, IIf(oTabla.PacienteId = 0, Null, oTabla.PacienteId))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteNroHistoria", adInteger, adParamInput, 0, oTabla.PacienteNroHistoria)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteIdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.PacienteIdTipoDocumento = 0, Null, oTabla.PacienteIdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteDocumento", adVarChar, adParamInput, 12, IIf(oTabla.PacienteDocumento = "", Null, oTabla.PacienteDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 150, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteSexo", adVarChar, adParamInput, 1, IIf(oTabla.PacienteSexo = "", Null, oTabla.PacienteSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteFNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.PacienteFNacimiento = 0, Null, oTabla.PacienteFNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoDNI", adVarChar, adParamInput, 8, IIf(oTabla.medicoDNI = "", Null, oTabla.medicoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medico", adVarChar, adParamInput, 150, IIf(oTabla.medico = "", Null, oTabla.medico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaHoraMuestra", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaHoraMuestra = 0, Null, oTabla.fechaHoraMuestra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoCodigo", adVarChar, adParamInput, 2, IIf(oTabla.EquipoCodigo = "", Null, oTabla.EquipoCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoRuta", adVarChar, adParamInput, 200, IIf(oTabla.EquipoRuta = "", Null, oTabla.EquipoRuta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Resultado", adVarChar, adParamInput, 3000, IIf(oTabla.Resultado = "", Null, oTabla.Resultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFechaHora", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ResultadoFechaHora = 0, Null, oTabla.ResultadoFechaHora))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoDNI", adVarChar, adParamInput, 8, IIf(oTabla.ResultadoDNI = "", Null, oTabla.ResultadoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@puntoCarga", adVarChar, adParamInput, 20, IIf(oTabla.PuntoCarga = "", Null, oTabla.PuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacPaterno", adVarChar, adParamInput, 20, oTabla.PacPaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacMaterno", adVarChar, adParamInput, 20, oTabla.PacMaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacPrimerNombre", adVarChar, adParamInput, 20, oTabla.PacPrimerNombre)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacSegundoNombre", adVarChar, adParamInput, 20, oTabla.PacSegundoNombre)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedPaterno", adVarChar, adParamInput, 20, oTabla.MedPaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedMaterno", adVarChar, adParamInput, 20, oTabla.MedMaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedNombres", adVarChar, adParamInput, 40, oTabla.MedNombres)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FteFinanciamiento", adVarChar, adParamInput, 20, oTabla.FteFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoServicio", adVarChar, adParamInput, 20, oTabla.TipoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PACS", adVarChar, adParamInput, 20, oTabla.PACS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LLAVE", adVarChar, adParamInput, 14, oTabla.LLAVE)
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
'   Descripción:    modifica un registro a la tabla InteroperaPacs
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInteroperaPacs) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaPacsModificar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, IIf(oTabla.cptCodigo = "", Null, oTabla.cptCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 250, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteId", adInteger, adParamInput, 0, IIf(oTabla.PacienteId = 0, Null, oTabla.PacienteId))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteNroHistoria", adInteger, adParamInput, 0, oTabla.PacienteNroHistoria)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteIdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.PacienteIdTipoDocumento = 0, Null, oTabla.PacienteIdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteDocumento", adVarChar, adParamInput, 12, IIf(oTabla.PacienteDocumento = "", Null, oTabla.PacienteDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 150, IIf(oTabla.Paciente = "", Null, oTabla.Paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteSexo", adVarChar, adParamInput, 1, IIf(oTabla.PacienteSexo = "", Null, oTabla.PacienteSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteFNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.PacienteFNacimiento = 0, Null, oTabla.PacienteFNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoDNI", adVarChar, adParamInput, 8, IIf(oTabla.medicoDNI = "", Null, oTabla.medicoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medico", adVarChar, adParamInput, 150, IIf(oTabla.medico = "", Null, oTabla.medico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaHoraMuestra", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaHoraMuestra = 0, Null, oTabla.fechaHoraMuestra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoCodigo", adVarChar, adParamInput, 2, IIf(oTabla.EquipoCodigo = "", Null, oTabla.EquipoCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoRuta", adVarChar, adParamInput, 200, IIf(oTabla.EquipoRuta = "", Null, oTabla.EquipoRuta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Resultado", adVarChar, adParamInput, 3000, IIf(oTabla.Resultado = "", Null, oTabla.Resultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFechaHora", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ResultadoFechaHora = 0, Null, oTabla.ResultadoFechaHora))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoDNI", adVarChar, adParamInput, 8, IIf(oTabla.ResultadoDNI = "", Null, oTabla.ResultadoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@puntoCarga", adVarChar, adParamInput, 20, IIf(oTabla.PuntoCarga = "", Null, oTabla.PuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacPaterno", adVarChar, adParamInput, 20, oTabla.PacPaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacMaterno", adVarChar, adParamInput, 20, oTabla.PacMaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacPrimerNombre", adVarChar, adParamInput, 20, oTabla.PacPrimerNombre)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacSegundoNombre", adVarChar, adParamInput, 20, oTabla.PacSegundoNombre)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedPaterno", adVarChar, adParamInput, 20, oTabla.MedPaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedMaterno", adVarChar, adParamInput, 20, oTabla.MedMaterno)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MedNombres", adVarChar, adParamInput, 40, oTabla.MedNombres)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FteFinanciamiento", adVarChar, adParamInput, 20, oTabla.FteFinanciamiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoServicio", adVarChar, adParamInput, 20, oTabla.TipoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PACS", adVarChar, adParamInput, 20, oTabla.PACS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LLAVE", adVarChar, adParamInput, 14, oTabla.LLAVE)
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
'   Descripción:    elimia o borra un registro a la tabla InteroperaPacs
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInteroperaPacs) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaPacsEliminar"
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
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
'   Descripción:    Selecciona un unico registro de la tabla InteroperaPacs
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInteroperaPacs) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteroperaPacsSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.cptCodigo = IIf(IsNull(oRecordset!cptCodigo), "", oRecordset!cptCodigo)
       oTabla.cpt = IIf(IsNull(oRecordset!cpt), "", oRecordset!cpt)
       oTabla.PacienteId = IIf(IsNull(oRecordset!PacienteId), 0, oRecordset!PacienteId)
       oTabla.PacienteNroHistoria = IIf(IsNull(oRecordset!PacienteNroHistoria), 0, oRecordset!PacienteNroHistoria)
       oTabla.PacienteIdTipoDocumento = IIf(IsNull(oRecordset!PacienteIdTipoDocumento), 0, oRecordset!PacienteIdTipoDocumento)
       oTabla.PacienteDocumento = IIf(IsNull(oRecordset!PacienteDocumento), "", oRecordset!PacienteDocumento)
       oTabla.Paciente = IIf(IsNull(oRecordset!Paciente), "", oRecordset!Paciente)
       oTabla.PacienteSexo = IIf(IsNull(oRecordset!PacienteSexo), "", oRecordset!PacienteSexo)
       oTabla.PacienteFNacimiento = IIf(IsNull(oRecordset!PacienteFNacimiento), 0, oRecordset!PacienteFNacimiento)
       oTabla.medicoDNI = IIf(IsNull(oRecordset!medicoDNI), "", oRecordset!medicoDNI)
       oTabla.medico = IIf(IsNull(oRecordset!medico), "", oRecordset!medico)
       oTabla.fechaHoraMuestra = IIf(IsNull(oRecordset!fechaHoraMuestra), 0, oRecordset!fechaHoraMuestra)
       oTabla.EquipoCodigo = IIf(IsNull(oRecordset!EquipoCodigo), "", oRecordset!EquipoCodigo)
       oTabla.EquipoRuta = IIf(IsNull(oRecordset!EquipoRuta), "", oRecordset!EquipoRuta)
       oTabla.Resultado = IIf(IsNull(oRecordset!Resultado), "", oRecordset!Resultado)
       oTabla.ResultadoFechaHora = IIf(IsNull(oRecordset!ResultadoFechaHora), 0, oRecordset!ResultadoFechaHora)
       oTabla.ResultadoDNI = IIf(IsNull(oRecordset!ResultadoDNI), "", oRecordset!ResultadoDNI)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.PuntoCarga = IIf(IsNull(oRecordset!PuntoCarga), "", oRecordset!PuntoCarga)
       oTabla.PacPaterno = IIf(IsNull(oRecordset!PacPaterno), "", oRecordset!PacPaterno)
       oTabla.PacMaterno = IIf(IsNull(oRecordset!PacMaterno), "", oRecordset!PacMaterno)
       oTabla.PacPrimerNombre = IIf(IsNull(oRecordset!PacPrimerNombre), "", oRecordset!PacPrimerNombre)
       oTabla.PacSegundoNombre = IIf(IsNull(oRecordset!PacSegundoNombre), "", oRecordset!PacSegundoNombre)
       oTabla.MedPaterno = IIf(IsNull(oRecordset!MedPaterno), "", oRecordset!MedPaterno)
       oTabla.MedMaterno = IIf(IsNull(oRecordset!MedMaterno), "", oRecordset!MedMaterno)
       oTabla.MedNombres = IIf(IsNull(oRecordset!MedNombres), "", oRecordset!MedNombres)
       oTabla.FteFinanciamiento = IIf(IsNull(oRecordset!FteFinanciamiento), "", oRecordset!FteFinanciamiento)
       oTabla.TipoServicio = IIf(IsNull(oRecordset!TipoServicio), "", oRecordset!TipoServicio)
       oTabla.PACS = IIf(IsNull(oRecordset!PACS), "", oRecordset!PACS)
       oTabla.LLAVE = IIf(IsNull(oRecordset!LLAVE), "", oRecordset!LLAVE)
   
   
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOInteroperaPacs)
   oTabla.IdOrden = 0
   oTabla.cptCodigo = ""
   oTabla.cpt = ""
   oTabla.PacienteId = 0
   oTabla.PacienteNroHistoria = 0
   oTabla.PacienteIdTipoDocumento = 0
   oTabla.PacienteDocumento = ""
   oTabla.Paciente = ""
   oTabla.PacienteSexo = ""
   oTabla.PacienteFNacimiento = Now
   oTabla.medicoDNI = ""
   oTabla.medico = ""
   oTabla.fechaHoraMuestra = Now
   oTabla.EquipoCodigo = ""
   oTabla.EquipoRuta = ""
   oTabla.Resultado = ""
   oTabla.ResultadoFechaHora = Now
   oTabla.ResultadoDNI = ""
   oTabla.IdEstado = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InteroperaPacs
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InteroperaPacs
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 05:52:12 a.m.
'------------------------------------------------------------------------------------
 

