VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "m_insumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: m_insumos
'        Autor: Daniel Enrique Barrantes
'        Fecha: 06/12/2013 12:51:37 p.m.
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
'   Descripción:    Inserta un registro a la tabla m_insumos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As Dom_insumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_insumosAgregar"
   Set oParameter = .CreateParameter("@ins_CodIns", adVarChar, adParamInput, 5, IIf(oTabla.ins_CodIns = "", Null, oTabla.ins_CodIns))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Nombre", adVarChar, adParamInput, 200, IIf(oTabla.ins_Nombre = "", Null, oTabla.ins_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_FormaFarmaceutica", adVarChar, adParamInput, 15, IIf(oTabla.ins_FormaFarmaceutica = "", Null, oTabla.ins_FormaFarmaceutica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Presen", adVarChar, adParamInput, 80, IIf(oTabla.ins_Presen = "", Null, oTabla.ins_Presen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Concen", adVarChar, adParamInput, 50, IIf(oTabla.ins_Concen = "", Null, oTabla.ins_Concen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Costo", adCurrency, adParamInput, 0, oTabla.ins_Costo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Observacion", adVarChar, adParamInput, 50, IIf(oTabla.ins_Observacion = "", Null, oTabla.ins_Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Petitorio", adVarChar, adParamInput, 1, IIf(oTabla.ins_Petitorio = "", "S", oTabla.ins_Petitorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_FecBaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ins_FecBaja = 0, Null, oTabla.ins_FecBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_DocBaja", adVarChar, adParamInput, 100, IIf(oTabla.ins_DocBaja = "", Null, oTabla.ins_DocBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.ins_IdEstado = "", Null, oTabla.ins_IdEstado))
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
'   Descripción:    modifica un registro a la tabla m_insumos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Dom_insumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_insumosModificar"
   Set oParameter = .CreateParameter("@ins_CodIns", adVarChar, adParamInput, 5, IIf(oTabla.ins_CodIns = "", Null, oTabla.ins_CodIns))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Nombre", adVarChar, adParamInput, 200, IIf(oTabla.ins_Nombre = "", Null, oTabla.ins_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_FormaFarmaceutica", adVarChar, adParamInput, 15, IIf(oTabla.ins_FormaFarmaceutica = "", Null, oTabla.ins_FormaFarmaceutica))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Presen", adVarChar, adParamInput, 80, IIf(oTabla.ins_Presen = "", Null, oTabla.ins_Presen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Concen", adVarChar, adParamInput, 50, IIf(oTabla.ins_Concen = "", Null, oTabla.ins_Concen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Costo", adCurrency, adParamInput, 0, oTabla.ins_Costo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Observacion", adVarChar, adParamInput, 50, IIf(oTabla.ins_Observacion = "", Null, oTabla.ins_Observacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_Petitorio", adVarChar, adParamInput, 1, IIf(oTabla.ins_Petitorio = "", "S", oTabla.ins_Petitorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_FecBaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ins_FecBaja = 0, Null, oTabla.ins_FecBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_DocBaja", adVarChar, adParamInput, 100, IIf(oTabla.ins_DocBaja = "", Null, oTabla.ins_DocBaja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ins_IdEstado", adVarChar, adParamInput, 1, IIf(oTabla.ins_IdEstado = "", Null, oTabla.ins_IdEstado))
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
'   Descripción:    elimia o borra un registro a la tabla m_insumos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Dom_insumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_insumosEliminar"
   Set oParameter = .CreateParameter("@ins_CodIns", adVarChar, adParamInput, 5, oTabla.ins_CodIns)
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
'   Descripción:    Selecciona un unico registro de la tabla m_insumos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Dom_insumos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "m_insumosSeleccionarPorId"
     Set oParameter = .CreateParameter("@ins_CodIns", adVarChar, adParamInput, 5, oTabla.ins_CodIns)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.ins_CodIns = IIf(IsNull(oRecordset!ins_CodIns), "", oRecordset!ins_CodIns)
       oTabla.ins_Nombre = IIf(IsNull(oRecordset!ins_Nombre), "", oRecordset!ins_Nombre)
       oTabla.ins_FormaFarmaceutica = IIf(IsNull(oRecordset!ins_FormaFarmaceutica), "", oRecordset!ins_FormaFarmaceutica)
       oTabla.ins_Presen = IIf(IsNull(oRecordset!ins_Presen), "", oRecordset!ins_Presen)
       oTabla.ins_Concen = IIf(IsNull(oRecordset!ins_Concen), "", oRecordset!ins_Concen)
       oTabla.ins_Costo = oRecordset!ins_Costo
       oTabla.ins_Observacion = IIf(IsNull(oRecordset!ins_Observacion), "", oRecordset!ins_Observacion)
       oTabla.ins_Petitorio = IIf(IsNull(oRecordset!ins_Petitorio), "", oRecordset!ins_Petitorio)
       oTabla.ins_FecBaja = IIf(IsNull(oRecordset!ins_FecBaja), 0, oRecordset!ins_FecBaja)
       oTabla.ins_DocBaja = IIf(IsNull(oRecordset!ins_DocBaja), "", oRecordset!ins_DocBaja)
       oTabla.ins_IdEstado = IIf(IsNull(oRecordset!ins_IdEstado), "", oRecordset!ins_IdEstado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Dom_insumos)
   oTabla.ins_CodIns = ""
   oTabla.ins_Nombre = ""
   oTabla.ins_FormaFarmaceutica = ""
   oTabla.ins_Presen = ""
   oTabla.ins_Concen = ""
   oTabla.ins_Costo = 0
   oTabla.ins_Observacion = ""
   oTabla.ins_Petitorio = ""
   oTabla.ins_FecBaja = Now
   oTabla.ins_DocBaja = ""
   oTabla.ins_IdEstado = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: m_insumos
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: m_insumos
'        Autor:   DAniel Barrantes B.
'        Fecha: 06/12/2013 12:51:37 p.m.
'------------------------------------------------------------------------------------
 

