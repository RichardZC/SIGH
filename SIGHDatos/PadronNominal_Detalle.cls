VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadronNominal_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PadronNominal_Detalle
'        Programado por: Palomino Y
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla PadronNominal_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoPadronNominal_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_DetalleAgregar"
   Set oParameter = .CreateParameter("@IdPaNomDetalle", adInteger, adParamInput, 0, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDoc", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDoc = 0, Null, oTabla.IdTipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumDocumento", adInteger, adParamInput, 0, IIf(oTabla.NumDocumento = 0, Null, oTabla.NumDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HistClinica", adVarChar, adParamInput, 10, IIf(oTabla.HistClinica = "", Null, oTabla.HistClinica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Nombres", adVarChar, adParamInput, 40, IIf(oTabla.Nombres = "", Null, oTabla.Nombres))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSexo", adInteger, adParamInput, 0, IIf(oTabla.idSexo = 0, Null, oTabla.idSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FecNacimiento", adVarChar, adParamInput, 10, IIf(oTabla.FecNacimiento = "", Null, oTabla.FecNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSeguro = 0, Null, oTabla.IdTipoSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumAfiliacion", adVarChar, adParamInput, 10, IIf(oTabla.NumAfiliacion = "", Null, oTabla.NumAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FecEvaluacion", adVarChar, adParamInput, 10, IIf(oTabla.FecEvaluacion = "", Null, oTabla.FecEvaluacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adVarChar, adParamInput, 5, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adVarChar, adParamInput, 3, IIf(oTabla.Talla = "", Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagNutricional", adInteger, adParamInput, 0, IIf(oTabla.IdDiagNutricional = 0, 0, oTabla.IdDiagNutricional))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodRenaes", adInteger, adParamInput, 0, IIf(oTabla.CodRenaes = 0, Null, oTabla.CodRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagPE", adInteger, adParamInput, 0, IIf(oTabla.IdDiagPE = 0, Null, oTabla.IdDiagPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagPT", adInteger, adParamInput, 0, IIf(oTabla.IdDiagPT = 0, Null, oTabla.IdDiagPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagTE", adInteger, adParamInput, 0, IIf(oTabla.IdDiagTE = 0, Null, oTabla.IdDiagTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Hemoglobina", adInteger, adParamInput, 0, oTabla.hemoglobina)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@heces", adVarChar, adParamInput, 2, oTabla.Heces)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPaNomDetalle = .Parameters("@IdPaNomDetalle")
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
'   Descripción:    modifica un registro a la tabla PadronNominal_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPadronNominal_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_DetalleModificar"
   Set oParameter = .CreateParameter("@IdPaNomDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdPaNomDetalle = 0, Null, oTabla.IdPaNomDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDoc", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDoc = 0, Null, oTabla.IdTipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumDocumento", adInteger, adParamInput, 0, IIf(oTabla.NumDocumento = 0, Null, oTabla.NumDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HistClinica", adVarChar, adParamInput, 10, IIf(oTabla.HistClinica = "", Null, oTabla.HistClinica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Nombres", adVarChar, adParamInput, 40, IIf(oTabla.Nombres = "", Null, oTabla.Nombres))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idSexo", adInteger, adParamInput, 0, IIf(oTabla.idSexo = 0, Null, oTabla.idSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FecNacimiento", adVarChar, adParamInput, 10, IIf(oTabla.FecNacimiento = "", Null, oTabla.FecNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSeguro = 0, Null, oTabla.IdTipoSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NumAfiliacion", adVarChar, adParamInput, 10, IIf(oTabla.NumAfiliacion = "", Null, oTabla.NumAfiliacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FecEvaluacion", adVarChar, adParamInput, 10, IIf(oTabla.FecEvaluacion = "", Null, oTabla.FecEvaluacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adVarChar, adParamInput, 5, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adVarChar, adParamInput, 3, IIf(oTabla.Talla = "", Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagNutricional", adInteger, adParamInput, 0, IIf(oTabla.IdDiagNutricional = 0, Null, oTabla.IdDiagNutricional))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodRenaes", adInteger, adParamInput, 0, IIf(oTabla.CodRenaes = 0, Null, oTabla.CodRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagPE", adInteger, adParamInput, 0, IIf(oTabla.IdDiagPE = 0, Null, oTabla.IdDiagPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagPT", adInteger, adParamInput, 0, IIf(oTabla.IdDiagPT = 0, Null, oTabla.IdDiagPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagTE", adInteger, adParamInput, 0, IIf(oTabla.IdDiagTE = 0, Null, oTabla.IdDiagTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Hemoglobina", adInteger, adParamInput, 0, oTabla.hemoglobina)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@heces", adVarChar, adParamInput, 2, oTabla.Heces)
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
'   Descripción:    elimia o borra un registro a la tabla PadronNominal_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPadronNominal_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_DetalleEliminar"
   Set oParameter = .CreateParameter("@IdPaNomDetalle", adInteger, adParamInput, 0, oTabla.IdPaNomDetalle)
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
'   Descripción:    Selecciona un unico registro de la tabla PadronNominal_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPadronNominal_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PadronNominal_DetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPaNomDetalle", adInteger, adParamInput, 0, oTabla.IdPaNomDetalle)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaNomDetalle = IIf(IsNull(oRecordset!IdPaNomDetalle), 0, oRecordset!IdPaNomDetalle)
       oTabla.IdTipoDoc = IIf(IsNull(oRecordset!IdTipoDoc), 0, oRecordset!IdTipoDoc)
       oTabla.NumDocumento = IIf(IsNull(oRecordset!NumDocumento), 0, oRecordset!NumDocumento)
       oTabla.HistClinica = IIf(IsNull(oRecordset!HistClinica), "", oRecordset!HistClinica)
       oTabla.ApellidoPaterno = IIf(IsNull(oRecordset!ApellidoPaterno), "", oRecordset!ApellidoPaterno)
       oTabla.ApellidoMaterno = IIf(IsNull(oRecordset!ApellidoMaterno), "", oRecordset!ApellidoMaterno)
       oTabla.Nombres = IIf(IsNull(oRecordset!Nombres), "", oRecordset!Nombres)
       oTabla.idSexo = IIf(IsNull(oRecordset!idSexo), 0, oRecordset!idSexo)
       oTabla.FecNacimiento = IIf(IsNull(oRecordset!FecNacimiento), "", oRecordset!FecNacimiento)
       oTabla.IdTipoSeguro = IIf(IsNull(oRecordset!IdTipoSeguro), 0, oRecordset!IdTipoSeguro)
       oTabla.NumAfiliacion = IIf(IsNull(oRecordset!NumAfiliacion), "", oRecordset!NumAfiliacion)
       oTabla.FecEvaluacion = IIf(IsNull(oRecordset!FecEvaluacion), "", oRecordset!FecEvaluacion)
       oTabla.Peso = IIf(IsNull(oRecordset!Peso), "", oRecordset!Peso)
       oTabla.Talla = IIf(IsNull(oRecordset!Talla), "", oRecordset!Talla)
       oTabla.IdDiagNutricional = IIf(IsNull(oRecordset!IdDiagNutricional), 0, oRecordset!IdDiagNutricional)
       oTabla.CodRenaes = IIf(IsNull(oRecordset!CodRenaes), 0, oRecordset!CodRenaes)
       oTabla.IdDiagPE = IIf(IsNull(oRecordset!IdDiagPE), 0, oRecordset!IdDiagPE)
       oTabla.IdDiagPT = IIf(IsNull(oRecordset!IdDiagPT), 0, oRecordset!IdDiagPT)
       oTabla.IdDiagTE = IIf(IsNull(oRecordset!IdDiagTE), 0, oRecordset!IdDiagTE)
       oTabla.hemoglobina = IIf(IsNull(oRecordset!hemoglobina), 0, oRecordset!hemoglobina)
       oTabla.Heces = IIf(IsNull(oRecordset!Heces), "NO", oRecordset!Heces)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoPadronNominal_Detalle)
   oTabla.IdPaNomDetalle = 0
   oTabla.IdTipoDoc = 0
   oTabla.NumDocumento = 0
   oTabla.HistClinica = ""
   oTabla.ApellidoPaterno = ""
   oTabla.ApellidoMaterno = ""
   oTabla.Nombres = ""
   oTabla.idSexo = 0
   oTabla.FecNacimiento = ""
   oTabla.IdTipoSeguro = 0
   oTabla.NumAfiliacion = ""
   oTabla.FecEvaluacion = ""
   oTabla.Peso = ""
   oTabla.Talla = ""
   oTabla.IdDiagNutricional = 0
   oTabla.CodRenaes = 0
   oTabla.IdDiagPE = 0
   oTabla.IdDiagPT = 0
   oTabla.IdDiagTE = 0
End Sub
