VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InteroperaEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InteroperaEquipos
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 06:15:40 a.m.
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
'   Descripción:    Inserta un registro a la tabla InteroperaEquipos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInteroperaEquipos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaEquiposAgregar"
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Marca", adVarChar, adParamInput, 100, IIf(oTabla.Marca = "", Null, oTabla.Marca))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Modelo", adVarChar, adParamInput, 100, IIf(oTabla.Modelo = "", Null, oTabla.Modelo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Tipo", adVarChar, adParamInput, 100, IIf(oTabla.Tipo = "", Null, oTabla.Tipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ruta", adVarChar, adParamInput, 200, IIf(oTabla.Ruta = "", Null, oTabla.Ruta))
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
'   Descripción:    modifica un registro a la tabla InteroperaEquipos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInteroperaEquipos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaEquiposModificar"
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Marca", adVarChar, adParamInput, 100, IIf(oTabla.Marca = "", Null, oTabla.Marca))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Modelo", adVarChar, adParamInput, 100, IIf(oTabla.Modelo = "", Null, oTabla.Modelo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Tipo", adVarChar, adParamInput, 100, IIf(oTabla.Tipo = "", Null, oTabla.Tipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ruta", adVarChar, adParamInput, 200, IIf(oTabla.Ruta = "", Null, oTabla.Ruta))
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
'   Descripción:    elimia o borra un registro a la tabla InteroperaEquipos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInteroperaEquipos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteroperaEquiposEliminar"
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, oTabla.Codigo)
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
'   Descripción:    Selecciona un unico registro de la tabla InteroperaEquipos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInteroperaEquipos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteroperaEquiposSeleccionarPorId"
     Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, oTabla.Codigo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.Codigo = IIf(IsNull(oRecordset!Codigo), "", oRecordset!Codigo)
       oTabla.Marca = IIf(IsNull(oRecordset!Marca), "", oRecordset!Marca)
       oTabla.Modelo = IIf(IsNull(oRecordset!Modelo), "", oRecordset!Modelo)
       oTabla.Tipo = IIf(IsNull(oRecordset!Tipo), "", oRecordset!Tipo)
       oTabla.Ruta = IIf(IsNull(oRecordset!Ruta), "", oRecordset!Ruta)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOInteroperaEquipos)
   oTabla.Codigo = ""
   oTabla.Marca = ""
   oTabla.Modelo = ""
   oTabla.Tipo = ""
   oTabla.Ruta = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InteroperaEquipos
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InteroperaEquipos
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 06:15:41 a.m.
'------------------------------------------------------------------------------------
 

