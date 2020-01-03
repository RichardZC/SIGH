VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesCE
'        Programado por: Barrantes D
'        Fecha: Marzo 2011
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros

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
'   Descripción:    Inserta un registro a la tabla atencionesCE
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtencionesCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "atencionesCEAgregar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaDniMedicoJamo", adVarChar, adParamInput, 8, IIf(oTabla.CitaDniMedicoJamo = "", Null, oTabla.CitaDniMedicoJamo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.CitaFecha = 0, Null, oTabla.CitaFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaMedico", adVarChar, adParamInput, 100, IIf(oTabla.CitaMedico = "", Null, oTabla.CitaMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaServicioJamo", adVarChar, adParamInput, 100, IIf(oTabla.CitaServicioJamo = "", Null, oTabla.CitaServicioJamo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaIdServicio", adInteger, adParamInput, 0, IIf(oTabla.CitaIdServicio = 0, Null, oTabla.CitaIdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaMotivo", adVarChar, adParamInput, 1000, IIf(oTabla.CitaMotivo = "", Null, oTabla.CitaMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaExamenClinico", adVarChar, adParamInput, 1000, IIf(oTabla.CitaExamenClinico = "", Null, oTabla.CitaExamenClinico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaDiagMed", adVarChar, adParamInput, 1000, IIf(oTabla.CitaDiagMed = "", Null, oTabla.CitaDiagMed))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaExClinicos", adVarChar, adParamInput, 3000, IIf(oTabla.CitaExClinicos = "", Null, oTabla.CitaExClinicos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaTratamiento", adVarChar, adParamInput, 1000, IIf(oTabla.CitaTratamiento = "", Null, oTabla.CitaTratamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaObservaciones", adVarChar, adParamInput, 1000, IIf(oTabla.CitaObservaciones = "", Null, oTabla.CitaObservaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaFechaAtencion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.CitaFechaAtencion = 0, Null, oTabla.CitaFechaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaIdUsuario", adInteger, adParamInput, 0, IIf(oTabla.CitaIdUsuario = 0, Null, oTabla.CitaIdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeEdad", adVarChar, adParamInput, 6, IIf(oTabla.TriajeEdad = "", Null, oTabla.TriajeEdad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePresion", adVarChar, adParamInput, 13, IIf(oTabla.TriajePresion = "", Null, oTabla.TriajePresion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeTalla", adVarChar, adParamInput, 7, IIf(oTabla.TriajeTalla = "", Null, oTabla.TriajeTalla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeTemperatura", adVarChar, adParamInput, 6, IIf(oTabla.TriajeTemperatura = "", Null, oTabla.TriajeTemperatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePeso", adVarChar, adParamInput, 7, IIf(oTabla.TriajePeso = "", Null, oTabla.TriajePeso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.TriajeFecha = 0, Null, oTabla.TriajeFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeIdUsuario", adInteger, adParamInput, 0, IIf(oTabla.TriajeIdUsuario = 0, Null, oTabla.TriajeIdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePulso", adInteger, adParamInput, 0, IIf(oTabla.TriajePulso = 0, Null, oTabla.TriajePulso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFrecRespiratoria", adInteger, adParamInput, 0, IIf(oTabla.TriajeFrecRespiratoria = 0, Null, oTabla.TriajeFrecRespiratoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaAntecedente", adVarChar, adParamInput, 1000, IIf(oTabla.CitaAntecedente = "", Null, oTabla.CitaAntecedente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePerimCefalico", adCurrency, adParamInput, 0, IIf(oTabla.TriajePerimCefalico = 0, Null, oTabla.TriajePerimCefalico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFrecCardiaca", adInteger, adParamInput, 0, IIf(oTabla.TriajeFrecCardiaca = 0, Null, oTabla.TriajeFrecCardiaca))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeOrigen", adInteger, adParamInput, 0, IIf(oTabla.TriajeOrigen = 0, Null, oTabla.TriajeOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePerimAbdominal", adVarChar, adParamInput, 30, IIf(oTabla.TriajePerimAbdominal = "", Null, oTabla.TriajePerimAbdominal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeSaturacionOxigeno", adVarChar, adParamInput, 3, IIf(Val(oTabla.TriajeSaturacionOxigeno) = 0, Null, oTabla.TriajeSaturacionOxigeno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Sub GrabaHistorico(oTabla As DOAtencionesCE)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO atencionesCE" & _
"           (idAtencion" & _
"           ,NroHistoriaClinica" & _
"           ,CitaMotivo" & _
"           ,CitaExamenClinico" & _
"           ,CitaExClinicos" & _
"           ,CitaTratamiento" & _
"           ,CitaObservaciones" & _
"           ,CitaIdUsuario" & _
"           ,TriajeEdad" & _
"           ,TriajePresion" & _
"           ,TriajeTalla" & _
"           ,TriajeTemperatura" & _
"           ,TriajePeso" & _
"           ,TriajeIdUsuario" & _
"           ,TriajePulso" & _
"           ,TriajeFrecRespiratoria" & _
"           ,CitaAntecedente" & _
"           ,TriajePerimCefalico" & _
"           ,TriajeFrecCardiaca" & _
"           ,TriajePerimAbdominal" & _
"           ,TriajeSaturacionOxigeno)"
     lcSql = lcSql & "     Values (" & _
"           " & oTabla.idAtencion & _
"           ," & oTabla.NroHistoriaClinica & _
"           ,'" & Left(oTabla.CitaMotivo, 250) & _
"           ','" & Left(oTabla.CitaExamenClinico, 250) & _
"           ','" & Left(oTabla.CitaExClinicos, 250) & _
"           ','" & Left(oTabla.CitaTratamiento, 250) & _
"           ','" & Left(oTabla.CitaObservaciones, 250) & _
"           '," & oTabla.CitaIdUsuario & _
"           ,'" & oTabla.TriajeEdad & _
"           ','" & oTabla.TriajePresion & _
"           ','" & oTabla.TriajeTalla & _
"           ','" & oTabla.TriajeTemperatura & _
"           ','" & oTabla.TriajePeso & _
"           ','" & oTabla.TriajeIdUsuario & _
"           ," & oTabla.TriajePulso & _
"           ," & oTabla.TriajeFrecRespiratoria & _
"           ,'" & Left(oTabla.CitaAntecedente, 250) & _
"           '," & oTabla.TriajePerimCefalico & _
"           ," & oTabla.TriajeFrecCardiaca & _
"           ,'" & oTabla.TriajePerimAbdominal & _
"           ','" & oTabla.TriajeSaturacionOxigeno & _
"           ')"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub



'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla atencionesCE
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionesCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "atencionesCEModificar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaDniMedicoJamo", adVarChar, adParamInput, 8, IIf(oTabla.CitaDniMedicoJamo = "", Null, oTabla.CitaDniMedicoJamo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.CitaFecha = 0, Null, oTabla.CitaFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaMedico", adVarChar, adParamInput, 100, IIf(oTabla.CitaMedico = "", Null, oTabla.CitaMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaServicioJamo", adVarChar, adParamInput, 100, IIf(oTabla.CitaServicioJamo = "", Null, oTabla.CitaServicioJamo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaIdServicio", adInteger, adParamInput, 0, IIf(oTabla.CitaIdServicio = 0, Null, oTabla.CitaIdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaMotivo", adVarChar, adParamInput, 1000, IIf(oTabla.CitaMotivo = "", Null, oTabla.CitaMotivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaExamenClinico", adVarChar, adParamInput, 1000, IIf(oTabla.CitaExamenClinico = "", Null, oTabla.CitaExamenClinico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaDiagMed", adVarChar, adParamInput, 1000, IIf(oTabla.CitaDiagMed = "", Null, oTabla.CitaDiagMed))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaExClinicos", adVarChar, adParamInput, 3000, IIf(oTabla.CitaExClinicos = "", Null, oTabla.CitaExClinicos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaTratamiento", adVarChar, adParamInput, 1000, IIf(oTabla.CitaTratamiento = "", Null, oTabla.CitaTratamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaObservaciones", adVarChar, adParamInput, 1000, IIf(oTabla.CitaObservaciones = "", Null, oTabla.CitaObservaciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaFechaAtencion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.CitaFechaAtencion = 0, Null, oTabla.CitaFechaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaIdUsuario", adInteger, adParamInput, 0, IIf(oTabla.CitaIdUsuario = 0, Null, oTabla.CitaIdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeEdad", adVarChar, adParamInput, 6, IIf(oTabla.TriajeEdad = "", Null, oTabla.TriajeEdad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePresion", adVarChar, adParamInput, 13, IIf(oTabla.TriajePresion = "", Null, oTabla.TriajePresion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeTalla", adVarChar, adParamInput, 7, IIf(oTabla.TriajeTalla = "", Null, oTabla.TriajeTalla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeTemperatura", adVarChar, adParamInput, 6, IIf(oTabla.TriajeTemperatura = "", Null, oTabla.TriajeTemperatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePeso", adVarChar, adParamInput, 7, IIf(oTabla.TriajePeso = "", Null, oTabla.TriajePeso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.TriajeFecha = 0, Null, oTabla.TriajeFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeIdUsuario", adInteger, adParamInput, 0, IIf(oTabla.TriajeIdUsuario = 0, Null, oTabla.TriajeIdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePulso", adInteger, adParamInput, 0, IIf(oTabla.TriajePulso = 0, Null, oTabla.TriajePulso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFrecRespiratoria", adInteger, adParamInput, 0, IIf(oTabla.TriajeFrecRespiratoria = 0, Null, oTabla.TriajeFrecRespiratoria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CitaAntecedente", adVarChar, adParamInput, 1000, IIf(oTabla.CitaAntecedente = "", Null, oTabla.CitaAntecedente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePerimCefalico", adCurrency, adParamInput, 0, IIf(oTabla.TriajePerimCefalico = 0, Null, oTabla.TriajePerimCefalico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeFrecCardiaca", adInteger, adParamInput, 0, IIf(oTabla.TriajeFrecCardiaca = 0, Null, oTabla.TriajeFrecCardiaca))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeOrigen", adInteger, adParamInput, 0, IIf(oTabla.TriajeOrigen = 0, Null, oTabla.TriajeOrigen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajePerimAbdominal", adVarChar, adParamInput, 30, IIf(oTabla.TriajePerimAbdominal = "", Null, oTabla.TriajePerimAbdominal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TriajeSaturacionOxigeno", adVarChar, adParamInput, 3, IIf(Val(oTabla.TriajeSaturacionOxigeno) = 0, Null, oTabla.TriajeSaturacionOxigeno))
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
'   Descripción:    elimia o borra un registro a la tabla atencionesCE
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionesCE) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "atencionesCEEliminar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla atencionesCE
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionesCE) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "atencionesCESeleccionarPorId"
     Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.NroHistoriaClinica = IIf(IsNull(oRecordset!NroHistoriaClinica), 0, oRecordset!NroHistoriaClinica)
       oTabla.CitaDniMedicoJamo = IIf(IsNull(oRecordset!CitaDniMedicoJamo), "", oRecordset!CitaDniMedicoJamo)
       oTabla.CitaFecha = IIf(IsNull(oRecordset!CitaFecha), 0, oRecordset!CitaFecha)
       oTabla.CitaMedico = IIf(IsNull(oRecordset!CitaMedico), "", oRecordset!CitaMedico)
       oTabla.CitaServicioJamo = IIf(IsNull(oRecordset!CitaServicioJamo), "", oRecordset!CitaServicioJamo)
       oTabla.CitaIdServicio = IIf(IsNull(oRecordset!CitaIdServicio), 0, oRecordset!CitaIdServicio)
       oTabla.CitaMotivo = IIf(IsNull(oRecordset!CitaMotivo), "", oRecordset!CitaMotivo)
       oTabla.CitaExamenClinico = IIf(IsNull(oRecordset!CitaExamenClinico), "", oRecordset!CitaExamenClinico)
       oTabla.CitaDiagMed = IIf(IsNull(oRecordset!CitaDiagMed), "", oRecordset!CitaDiagMed)
       oTabla.CitaExClinicos = IIf(IsNull(oRecordset!CitaExClinicos), "", oRecordset!CitaExClinicos)
       oTabla.CitaTratamiento = IIf(IsNull(oRecordset!CitaTratamiento), "", oRecordset!CitaTratamiento)
       oTabla.CitaObservaciones = IIf(IsNull(oRecordset!CitaObservaciones), "", oRecordset!CitaObservaciones)
       oTabla.CitaFechaAtencion = IIf(IsNull(oRecordset!CitaFechaAtencion), 0, oRecordset!CitaFechaAtencion)
       oTabla.CitaIdUsuario = IIf(IsNull(oRecordset!CitaIdUsuario), 0, oRecordset!CitaIdUsuario)
       oTabla.TriajeEdad = IIf(IsNull(oRecordset!TriajeEdad), "", oRecordset!TriajeEdad)
       oTabla.TriajePresion = IIf(IsNull(oRecordset!TriajePresion), "", oRecordset!TriajePresion)
       oTabla.TriajeTalla = IIf(IsNull(oRecordset!TriajeTalla), "", oRecordset!TriajeTalla)
       oTabla.TriajeTemperatura = IIf(IsNull(oRecordset!TriajeTemperatura), "", oRecordset!TriajeTemperatura)
       oTabla.TriajePeso = IIf(IsNull(oRecordset!TriajePeso), "", oRecordset!TriajePeso)
       oTabla.TriajeFecha = IIf(IsNull(oRecordset!TriajeFecha), 0, oRecordset!TriajeFecha)
       oTabla.TriajeIdUsuario = IIf(IsNull(oRecordset!TriajeIdUsuario), 0, oRecordset!TriajeIdUsuario)
       oTabla.TriajeFrecRespiratoria = IIf(IsNull(oRecordset!TriajeFrecRespiratoria), 0, oRecordset!TriajeFrecRespiratoria)
       oTabla.TriajePulso = IIf(IsNull(oRecordset!TriajePulso), 0, oRecordset!TriajePulso)
       oTabla.CitaAntecedente = IIf(IsNull(oRecordset!CitaAntecedente), "", oRecordset!CitaAntecedente)
       oTabla.TriajePerimCefalico = IIf(IsNull(oRecordset!TriajePerimCefalico), 0, oRecordset!TriajePerimCefalico)
       oTabla.TriajeFrecCardiaca = IIf(IsNull(oRecordset!TriajeFrecCardiaca), 0, oRecordset!TriajeFrecCardiaca)
       oTabla.TriajeOrigen = IIf(IsNull(oRecordset!TriajeOrigen), 0, oRecordset!TriajeOrigen)
       oTabla.TriajePerimAbdominal = IIf(IsNull(oRecordset!TriajePerimAbdominal), "", oRecordset!TriajePerimAbdominal)
       oTabla.TriajeSaturacionOxigeno = IIf(IsNull(oRecordset!TriajeSaturacionOxigeno), "", oRecordset!TriajeSaturacionOxigeno)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtencionesCE)
   oTabla.idAtencion = 0
   oTabla.NroHistoriaClinica = 0
   oTabla.CitaDniMedicoJamo = ""
   oTabla.CitaFecha = Now
   oTabla.CitaMedico = ""
   oTabla.CitaServicioJamo = ""
   oTabla.CitaIdServicio = 0
   oTabla.CitaMotivo = ""
   oTabla.CitaExamenClinico = ""
   oTabla.CitaDiagMed = ""
   oTabla.CitaExClinicos = ""
   oTabla.CitaTratamiento = ""
   oTabla.CitaObservaciones = ""
   oTabla.CitaFechaAtencion = Now
   oTabla.CitaIdUsuario = 0
   oTabla.TriajeEdad = ""
   oTabla.TriajePresion = ""
   oTabla.TriajeTalla = ""
   oTabla.TriajeTemperatura = ""
   oTabla.TriajePeso = ""
   oTabla.TriajeFecha = Now
   oTabla.TriajeIdUsuario = 0
   oTabla.TriajeFrecRespiratoria = 0
   oTabla.TriajePulso = 0
   oTabla.CitaAntecedente = ""
   oTabla.TriajePerimCefalico = 0
   oTabla.TriajeFrecCardiaca = 0
   oTabla.TriajeOrigen = sightriajeorigen.Triaje
   oTabla.TriajePerimAbdominal = ""
End Sub

'MARIO 10/06
Function SeleccionarPorNroHistoria(ByVal oTabla As DOAtencionesCE) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorNroHistoria = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesCeXnrohistoriaTriaje"
     Set oParameter = .CreateParameter("@ml_NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
    oTabla.TriajeTalla = IIf(IsNull(oRecordset!TriajeTalla), "", oRecordset!TriajeTalla)
    oTabla.TriajePeso = IIf(IsNull(oRecordset!TriajePeso), "", oRecordset!TriajePeso)
    SeleccionarPorNroHistoria = True
    ms_MensajeError = ""
   Else
    oTabla.TriajeTalla = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'mgaray201411e
Public Function AtencionesCeListaTriaje(ByVal oTabla As DOAtencionesCE) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesCeListaTriaje"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesCeListaTriaje = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
