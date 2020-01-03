VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterLabCabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterLabCabecera
'        Autor: Daniel Enrique Barrantes
'        Fecha: 12/05/2018 09:32:39 a.m.
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
'   Descripción:    Inserta un registro a la tabla InterLabCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInterLabCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabCabeceraAgregar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, IIf(oTabla.identificador = 0, Null, oTabla.identificador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@num_orden", adVarChar, adParamInput, 40, IIf(oTabla.num_orden = "", Null, oTabla.num_orden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Historia", adVarChar, adParamInput, 10, IIf(oTabla.Historia = "", Null, oTabla.Historia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteEmail", adVarChar, adParamInput, 50, IIf(oTabla.PacienteEmail = "", Null, oTabla.PacienteEmail))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteFiliacion", adVarChar, adParamInput, 1, IIf(oTabla.PacienteFiliacion = "", Null, oTabla.PacienteFiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteId", adInteger, adParamInput, 0, IIf(oTabla.PacienteId = 0, Null, oTabla.PacienteId))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.PacienteTipoDocumento = 0, Null, oTabla.PacienteTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteDocumento", adVarChar, adParamInput, 12, IIf(oTabla.PacienteDocumento = "", Null, oTabla.PacienteDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@paciente", adVarChar, adParamInput, 150, IIf(oTabla.paciente = "", Null, oTabla.paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@sexo", adVarChar, adParamInput, 1, IIf(oTabla.sexo = "", Null, oTabla.sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fnacimiento = 0, Null, oTabla.fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cama", adVarChar, adParamInput, 20, IIf(oTabla.cama = "", Null, oTabla.cama))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dxCodigo", adVarChar, adParamInput, 20, IIf(oTabla.dxCodigo = "", Null, oTabla.dxCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 300, IIf(oTabla.dx = "", Null, oTabla.dx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoDNI", adVarChar, adParamInput, 8, IIf(oTabla.medicoDNI = "", Null, oTabla.medicoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medico", adVarChar, adParamInput, 150, IIf(oTabla.medico = "", Null, oTabla.medico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@servicioCodigo", adVarChar, adParamInput, 6, IIf(oTabla.servicioCodigo = "", Null, oTabla.servicioCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaHoraMuestra", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaHoraMuestra = 0, Null, oTabla.fechaHoraMuestra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoColegiatura", adVarChar, adParamInput, 10, IIf(oTabla.medicoColegiatura = "", Null, oTabla.medicoColegiatura))
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
'   Descripción:    modifica un registro a la tabla InterLabCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInterLabCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabCabeceraModificar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, IIf(oTabla.identificador = 0, Null, oTabla.identificador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@num_orden", adVarChar, adParamInput, 40, IIf(oTabla.num_orden = "", Null, oTabla.num_orden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Historia", adVarChar, adParamInput, 10, IIf(oTabla.Historia = "", Null, oTabla.Historia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteEmail", adVarChar, adParamInput, 50, IIf(oTabla.PacienteEmail = "", Null, oTabla.PacienteEmail))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteFiliacion", adVarChar, adParamInput, 1, IIf(oTabla.PacienteFiliacion = "", Null, oTabla.PacienteFiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteId", adInteger, adParamInput, 0, IIf(oTabla.PacienteId = 0, Null, oTabla.PacienteId))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.PacienteTipoDocumento = 0, Null, oTabla.PacienteTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteDocumento", adVarChar, adParamInput, 12, IIf(oTabla.PacienteDocumento = "", Null, oTabla.PacienteDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@paciente", adVarChar, adParamInput, 150, IIf(oTabla.paciente = "", Null, oTabla.paciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@sexo", adVarChar, adParamInput, 1, IIf(oTabla.sexo = "", Null, oTabla.sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fnacimiento = 0, Null, oTabla.fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cama", adVarChar, adParamInput, 20, IIf(oTabla.cama = "", Null, oTabla.cama))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dxCodigo", adVarChar, adParamInput, 20, IIf(oTabla.dxCodigo = "", Null, oTabla.dxCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 300, IIf(oTabla.dx = "", Null, oTabla.dx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoDNI", adVarChar, adParamInput, 8, IIf(oTabla.medicoDNI = "", Null, oTabla.medicoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medico", adVarChar, adParamInput, 150, IIf(oTabla.medico = "", Null, oTabla.medico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@servicioCodigo", adVarChar, adParamInput, 6, IIf(oTabla.servicioCodigo = "", Null, oTabla.servicioCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaHoraMuestra", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaHoraMuestra = 0, Null, oTabla.fechaHoraMuestra))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@medicoColegiatura", adVarChar, adParamInput, 10, IIf(oTabla.medicoColegiatura = "", Null, oTabla.medicoColegiatura))
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
'   Descripción:    elimia o borra un registro a la tabla InterLabCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInterLabCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabCabeceraEliminar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, oTabla.identificador)
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
'   Descripción:    Selecciona un unico registro de la tabla InterLabCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInterLabCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InterLabCabeceraSeleccionarPorId"
     Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, oTabla.identificador)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.identificador = IIf(IsNull(oRecordset!identificador), 0, oRecordset!identificador)
       oTabla.num_orden = IIf(IsNull(oRecordset!num_orden), "", oRecordset!num_orden)
       oTabla.Historia = IIf(IsNull(oRecordset!Historia), "", oRecordset!Historia)
       oTabla.PacienteEmail = IIf(IsNull(oRecordset!PacienteEmail), "", oRecordset!PacienteEmail)
       oTabla.PacienteFiliacion = IIf(IsNull(oRecordset!PacienteFiliacion), "", oRecordset!PacienteFiliacion)
       oTabla.PacienteId = IIf(IsNull(oRecordset!PacienteId), 0, oRecordset!PacienteId)
       oTabla.PacienteTipoDocumento = IIf(IsNull(oRecordset!PacienteTipoDocumento), 0, oRecordset!PacienteTipoDocumento)
       oTabla.PacienteDocumento = IIf(IsNull(oRecordset!PacienteDocumento), "", oRecordset!PacienteDocumento)
       oTabla.paciente = IIf(IsNull(oRecordset!paciente), "", oRecordset!paciente)
       oTabla.sexo = IIf(IsNull(oRecordset!sexo), "", oRecordset!sexo)
       oTabla.fnacimiento = IIf(IsNull(oRecordset!fnacimiento), 0, oRecordset!fnacimiento)
       oTabla.cama = IIf(IsNull(oRecordset!cama), "", oRecordset!cama)
       oTabla.dxCodigo = IIf(IsNull(oRecordset!dxCodigo), "", oRecordset!dxCodigo)
       oTabla.dx = IIf(IsNull(oRecordset!dx), "", oRecordset!dx)
       oTabla.medicoDNI = IIf(IsNull(oRecordset!medicoDNI), "", oRecordset!medicoDNI)
       oTabla.medico = IIf(IsNull(oRecordset!medico), "", oRecordset!medico)
       oTabla.servicioCodigo = IIf(IsNull(oRecordset!servicioCodigo), "", oRecordset!servicioCodigo)
       oTabla.fechaHoraMuestra = IIf(IsNull(oRecordset!fechaHoraMuestra), 0, oRecordset!fechaHoraMuestra)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.medicoColegiatura = IIf(IsNull(oRecordset!medicoColegiatura), "", oRecordset!medicoColegiatura)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOInterLabCabecera)
   oTabla.identificador = 0
   oTabla.num_orden = ""
   oTabla.Historia = ""
   oTabla.PacienteEmail = ""
   oTabla.PacienteFiliacion = ""
   oTabla.PacienteId = 0
   oTabla.PacienteTipoDocumento = 0
   oTabla.PacienteDocumento = ""
   oTabla.paciente = ""
   oTabla.sexo = ""
   oTabla.fnacimiento = Now
   oTabla.cama = ""
   oTabla.dxCodigo = ""
   oTabla.dx = ""
   oTabla.medicoDNI = ""
   oTabla.medico = ""
   oTabla.servicioCodigo = ""
   oTabla.fechaHoraMuestra = Now
   oTabla.IdEstado = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InterLabCabecera
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InterLabCabecera
'        Autor:   DAniel Barrantes B.
'        Fecha: 12/05/2018 09:32:39 a.m.
'------------------------------------------------------------------------------------
 

