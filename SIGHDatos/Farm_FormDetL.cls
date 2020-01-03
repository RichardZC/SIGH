VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Farm_FormDetL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: farm_formdetl
'        Autor: Daniel Enrique Barrantes
'        Fecha: 26/04/2019 10:38:37 a.m.
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
'   Descripción:    Inserta un registro a la tabla farm_formdetl
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarm_formdetl) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetlAgregar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, oTabla.CODIGO_EJE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, oTabla.CODIGO_PRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, oTabla.TIPSUM)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, oTabla.ANNOMES)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, oTabla.CODIGO_MED)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LOTE", adVarChar, adParamInput, 15, oTabla.Lote)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FECHVTO", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FECHVTO = 0, Null, oTabla.FECHVTO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SALDO", adInteger, adParamInput, 0, oTabla.SALDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIT", adVarChar, adParamInput, 1, oTabla.SIT)
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
'   Descripción:    modifica un registro a la tabla farm_formdetl
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarm_formdetl) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetlModificar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, oTabla.CODIGO_EJE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, oTabla.CODIGO_PRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, oTabla.TIPSUM)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, oTabla.ANNOMES)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, oTabla.CODIGO_MED)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LOTE", adVarChar, adParamInput, 15, oTabla.Lote)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FECHVTO", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FECHVTO = 0, Null, oTabla.FECHVTO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SALDO", adInteger, adParamInput, 0, oTabla.SALDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIT", adVarChar, adParamInput, 1, oTabla.SIT)
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
'   Descripción:    elimia o borra un registro a la tabla farm_formdetl
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarm_formdetl) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetlEliminar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
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
'   Descripción:    Selecciona un unico registro de la tabla farm_formdetl
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarm_formdetl) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "farm_formdetlSeleccionarPorId"
        Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, oTabla.CODIGO_MED)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.CODIGO_EJE = oRecordset!CODIGO_EJE
       oTabla.CODIGO_PRE = oRecordset!CODIGO_PRE
       oTabla.TIPSUM = oRecordset!TIPSUM
       oTabla.ANNOMES = oRecordset!ANNOMES
       oTabla.CODIGO_MED = oRecordset!CODIGO_MED
       oTabla.Lote = oRecordset!Lote
       oTabla.FECHVTO = IIf(IsNull(oRecordset!FECHVTO), 0, oRecordset!FECHVTO)
       oTabla.SALDO = oRecordset!SALDO
       oTabla.SIT = oRecordset!SIT
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarm_formdetl)
   oTabla.CODIGO_EJE = ""
   oTabla.CODIGO_PRE = ""
   oTabla.TIPSUM = ""
   oTabla.ANNOMES = ""
   oTabla.CODIGO_MED = ""
   oTabla.Lote = ""
   oTabla.FECHVTO = Now
   oTabla.SALDO = 0#
   oTabla.SIT = ""
End Sub

Function EliminarXcodigo(ByVal oTabla As DOfarm_formdetl) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarXcodigo = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetlEliminarXcodigo"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, IIf(oTabla.CODIGO_MED = "", Null, oTabla.CODIGO_MED))
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarXcodigo = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: farm_formdetl
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: farm_formdetl
'        Autor:   DAniel Barrantes B.
'        Fecha: 26/04/2019 10:38:37 a.m.
'------------------------------------------------------------------------------------
 

