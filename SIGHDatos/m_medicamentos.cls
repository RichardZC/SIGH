VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "m_medicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla m_medicamentos
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
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
'   Descripción:    Inserta un registro a la tabla m_medicamentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As Dom_medicamentos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_medicamentosAgregar"
   Set oParameter = .CreateParameter("@med_CodMed", adVarChar, adParamInput, 13, IIf(oTabla.med_CodMed = "", Null, oTabla.med_CodMed))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Nombre", adVarChar, adParamInput, 100, IIf(oTabla.med_Nombre = "", Null, oTabla.med_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FormaFarmaceutica", adVarChar, adParamInput, 15, IIf(oTabla.med_FormaFarmaceutica = "", Null, oTabla.med_FormaFarmaceutica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Presen", adVarChar, adParamInput, 20, IIf(oTabla.med_Presen = "", "", oTabla.med_Presen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Concen", adVarChar, adParamInput, 50, IIf(oTabla.med_Concen = "", Null, oTabla.med_Concen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Costo", adCurrency, adParamInput, 0, oTabla.med_Costo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio", adVarChar, adParamInput, 1, IIf(oTabla.med_Petitorio = "", "S", oTabla.med_Petitorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio2005", adVarChar, adParamInput, 6, IIf(oTabla.med_Petitorio2005 = "", Null, oTabla.med_Petitorio2005))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio2010", adVarChar, adParamInput, 6, IIf(oTabla.med_Petitorio2010 = "", Null, oTabla.med_Petitorio2010))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FecBaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.med_FecBaja = 0, Null, oTabla.med_FecBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FFDigemid", adVarChar, adParamInput, 100, IIf(oTabla.med_FFDigemid = "", Null, oTabla.med_FFDigemid))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.med_IdEstado = "", Null, oTabla.med_IdEstado))
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
'   Descripción:    modifica un registro a la tabla m_medicamentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Dom_medicamentos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_medicamentosModificar"
   Set oParameter = .CreateParameter("@med_CodMed", adVarChar, adParamInput, 13, IIf(oTabla.med_CodMed = "", Null, oTabla.med_CodMed))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Nombre", adVarChar, adParamInput, 100, IIf(oTabla.med_Nombre = "", Null, oTabla.med_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FormaFarmaceutica", adVarChar, adParamInput, 15, IIf(oTabla.med_FormaFarmaceutica = "", Null, oTabla.med_FormaFarmaceutica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Presen", adVarChar, adParamInput, 20, IIf(oTabla.med_Presen = "", "", oTabla.med_Presen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Concen", adVarChar, adParamInput, 50, IIf(oTabla.med_Concen = "", Null, oTabla.med_Concen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Costo", adCurrency, adParamInput, 0, oTabla.med_Costo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio", adVarChar, adParamInput, 1, IIf(oTabla.med_Petitorio = "", "S", oTabla.med_Petitorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio2005", adVarChar, adParamInput, 6, IIf(oTabla.med_Petitorio2005 = "", Null, oTabla.med_Petitorio2005))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_Petitorio2010", adVarChar, adParamInput, 6, IIf(oTabla.med_Petitorio2010 = "", Null, oTabla.med_Petitorio2010))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FecBaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.med_FecBaja = 0, Null, oTabla.med_FecBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_FFDigemid", adVarChar, adParamInput, 100, IIf(oTabla.med_FFDigemid = "", Null, oTabla.med_FFDigemid))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@med_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.med_IdEstado = "", Null, oTabla.med_IdEstado))
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
'   Descripción:    elimia o borra un registro a la tabla m_medicamentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Dom_medicamentos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_medicamentosEliminar"
   Set oParameter = .CreateParameter("@med_CodMed", adVarChar, adParamInput, 13, oTabla.med_CodMed)
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
'   Descripción:    Selecciona un unico registro de la tabla m_medicamentos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Dom_medicamentos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "m_medicamentosSeleccionarPorId"
     Set oParameter = .CreateParameter("@med_CodMed", adVarChar, adParamInput, 13, oTabla.med_CodMed)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.med_CodMed = IIf(IsNull(oRecordset!med_CodMed), "", oRecordset!med_CodMed)
       oTabla.med_Nombre = IIf(IsNull(oRecordset!med_Nombre), "", oRecordset!med_Nombre)
       oTabla.med_FormaFarmaceutica = IIf(IsNull(oRecordset!med_FormaFarmaceutica), "", oRecordset!med_FormaFarmaceutica)
       oTabla.med_Presen = IIf(IsNull(oRecordset!med_Presen), "", oRecordset!med_Presen)
       oTabla.med_Concen = IIf(IsNull(oRecordset!med_Concen), "", oRecordset!med_Concen)
       oTabla.med_Costo = oRecordset!med_Costo
       oTabla.med_Petitorio = IIf(IsNull(oRecordset!med_Petitorio), "", oRecordset!med_Petitorio)
       oTabla.med_Petitorio2005 = IIf(IsNull(oRecordset!med_Petitorio2005), "", oRecordset!med_Petitorio2005)
       oTabla.med_Petitorio2010 = IIf(IsNull(oRecordset!med_Petitorio2010), "", oRecordset!med_Petitorio2010)
       oTabla.med_FecBaja = IIf(IsNull(oRecordset!med_FecBaja), 0, oRecordset!med_FecBaja)
       oTabla.med_FFDigemid = IIf(IsNull(oRecordset!med_FFDigemid), "", oRecordset!med_FFDigemid)
       oTabla.med_IdEstado = IIf(IsNull(oRecordset!med_IdEstado), "", oRecordset!med_IdEstado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Dom_medicamentos)
   oTabla.med_CodMed = ""
   oTabla.med_Nombre = ""
   oTabla.med_FormaFarmaceutica = ""
   oTabla.med_Presen = ""
   oTabla.med_Concen = ""
   oTabla.med_Costo = 0
   oTabla.med_Petitorio = ""
   oTabla.med_Petitorio2005 = ""
   oTabla.med_Petitorio2010 = ""
   oTabla.med_FecBaja = Now
   oTabla.med_FFDigemid = ""
   oTabla.med_IdEstado = ""
End Sub

