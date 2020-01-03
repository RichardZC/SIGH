VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PacientesDatosAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PacientesDatosAdicionales
'        Programado por: Barrantes D
'        Fecha: Julio 2012
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
'   Descripción:    Inserta un registro a la tabla PacientesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoPacienteDatosAdd) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesDatosAdicionalesAgregar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedentes", adVarChar, adParamInput, 1000, IIf(oTabla.antecedentes = "", Null, oTabla.antecedentes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedAlergico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedAlergico = "", Null, oTabla.antecedAlergico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedObstetrico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedObstetrico = "", Null, oTabla.antecedObstetrico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedQuirurgico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedQuirurgico = "", Null, oTabla.antecedQuirurgico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedFamiliar", adVarChar, adParamInput, 1000, IIf(oTabla.antecedFamiliar = "", Null, oTabla.antecedFamiliar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedPatologico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedPatologico = "", Null, oTabla.antecedPatologico))
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
'   Descripción:    modifica un registro a la tabla PacientesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPacienteDatosAdd) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesDatosAdicionalesModificar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedentes", adVarChar, adParamInput, 1000, IIf(oTabla.antecedentes = "", Null, oTabla.antecedentes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedAlergico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedAlergico = "", Null, oTabla.antecedAlergico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedObstetrico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedObstetrico = "", Null, oTabla.antecedObstetrico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedQuirurgico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedQuirurgico = "", Null, oTabla.antecedQuirurgico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedFamiliar", adVarChar, adParamInput, 1000, IIf(oTabla.antecedFamiliar = "", Null, oTabla.antecedFamiliar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@antecedPatologico", adVarChar, adParamInput, 1000, IIf(oTabla.antecedPatologico = "", Null, oTabla.antecedPatologico))
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
'   Descripción:    elimia o borra un registro a la tabla PacientesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPacienteDatosAdd) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesDatosAdicionalesEliminar"
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
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
'   Descripción:    Selecciona un unico registro de la tabla PacientesDatosAdicionales
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPacienteDatosAdd) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PacientesDatosAdicionalesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.antecedentes = IIf(IsNull(oRecordset!antecedentes), "", oRecordset!antecedentes)
       oTabla.antecedAlergico = IIf(IsNull(oRecordset!antecedAlergico), "", oRecordset!antecedAlergico)
       oTabla.antecedObstetrico = IIf(IsNull(oRecordset!antecedObstetrico), "", oRecordset!antecedObstetrico)
       oTabla.antecedQuirurgico = IIf(IsNull(oRecordset!antecedQuirurgico), "", oRecordset!antecedQuirurgico)
       oTabla.antecedFamiliar = IIf(IsNull(oRecordset!antecedFamiliar), "", oRecordset!antecedFamiliar)
       oTabla.antecedPatologico = IIf(IsNull(oRecordset!antecedPatologico), "", oRecordset!antecedPatologico)
       oTabla.fNacimientoCalculada = IIf(IsNull(oRecordset!fNacimientoCalculada), False, oRecordset!fNacimientoCalculada)
       SeleccionarPorId = True
       ms_MensajeError = ""
   Else
       oTabla.IdPaciente = 0
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoPacienteDatosAdd)
   oTabla.IdPaciente = 0
   oTabla.antecedentes = ""
End Sub

Function DatosPersonalesAgregar(ByVal oTabla As DoPacienteDatosAdd) As Boolean
On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
     
    DatosPersonalesAgregar = False
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PacientesDatosAdicionalesPersonalesAgregar"
       Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
       .Parameters.Append oParameter
       Set oParameter = .CreateParameter("FNacimientoCalculada", adBoolean, adParamInput, 0, oTabla.fNacimientoCalculada)
       .Parameters.Append oParameter
       .Execute
    End With
     
    DatosPersonalesAgregar = True
    ms_MensajeError = ""
 
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
End Function

