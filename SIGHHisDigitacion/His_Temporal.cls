VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "His_Temporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para mantenimiento de tabla HIS_TEMPORAL
'        Programado por: Cachay F
'        Fecha: Febrero 2014
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
'   Descripción:    Inserta un registro a la tabla HIS_TEMPORAL
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoHis_Temporal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_TEMPORALAgregar"
   Set oParameter = .CreateParameter("@Codigo1", adVarChar, adParamInput, 6, IIf(oTabla.Codigo1 = "", Null, oTabla.Codigo1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo2", adVarChar, adParamInput, 6, IIf(oTabla.Codigo2 = "", Null, oTabla.Codigo2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo3", adVarChar, adParamInput, 6, IIf(oTabla.Codigo3 = "", Null, oTabla.Codigo3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo4", adVarChar, adParamInput, 6, IIf(oTabla.Codigo4 = "", Null, oTabla.Codigo4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo5", adVarChar, adParamInput, 6, IIf(oTabla.Codigo5 = "", Null, oTabla.Codigo5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo6", adVarChar, adParamInput, 6, IIf(oTabla.Codigo6 = "", Null, oTabla.Codigo6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf1", adVarChar, adParamInput, 1, IIf(oTabla.LabConf1 = "", Null, oTabla.LabConf1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf2", adVarChar, adParamInput, 1, IIf(oTabla.LabConf2 = "", Null, oTabla.LabConf2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf3", adVarChar, adParamInput, 1, IIf(oTabla.LabConf3 = "", Null, oTabla.LabConf3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf4", adVarChar, adParamInput, 1, IIf(oTabla.LabConf4 = "", Null, oTabla.LabConf4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf5", adVarChar, adParamInput, 1, IIf(oTabla.LabConf5 = "", Null, oTabla.LabConf5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf6", adVarChar, adParamInput, 1, IIf(oTabla.LabConf6 = "", Null, oTabla.LabConf6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost1", adInteger, adParamInput, 0, IIf(oTabla.Diagnost1 = 0, Null, oTabla.Diagnost1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost2", adInteger, adParamInput, 0, IIf(oTabla.Diagnost2 = 0, Null, oTabla.Diagnost2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost3", adInteger, adParamInput, 0, IIf(oTabla.Diagnost3 = 0, Null, oTabla.Diagnost3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost4", adInteger, adParamInput, 0, IIf(oTabla.Diagnost4 = 0, Null, oTabla.Diagnost4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost5", adInteger, adParamInput, 0, IIf(oTabla.Diagnost5 = 0, Null, oTabla.Diagnost5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost6", adInteger, adParamInput, 0, IIf(oTabla.Diagnost6 = 0, Null, oTabla.Diagnost6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoEdad", adInteger, adParamInput, 0, IIf(oTabla.TipoEdad = 0, Null, oTabla.TipoEdad))
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
'   Descripción:    modifica un registro a la tabla HIS_TEMPORAL
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoHis_Temporal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_TEMPORALModificar"
   Set oParameter = .CreateParameter("@Codigo1", adVarChar, adParamInput, 6, IIf(oTabla.Codigo1 = "", Null, oTabla.Codigo1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo2", adVarChar, adParamInput, 6, IIf(oTabla.Codigo2 = "", Null, oTabla.Codigo2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo3", adVarChar, adParamInput, 6, IIf(oTabla.Codigo3 = "", Null, oTabla.Codigo3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo4", adVarChar, adParamInput, 6, IIf(oTabla.Codigo4 = "", Null, oTabla.Codigo4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo5", adVarChar, adParamInput, 6, IIf(oTabla.Codigo5 = "", Null, oTabla.Codigo5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo6", adVarChar, adParamInput, 6, IIf(oTabla.Codigo6 = "", Null, oTabla.Codigo6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf1", adVarChar, adParamInput, 1, IIf(oTabla.LabConf1 = "", Null, oTabla.LabConf1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf2", adVarChar, adParamInput, 1, IIf(oTabla.LabConf2 = "", Null, oTabla.LabConf2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf3", adVarChar, adParamInput, 1, IIf(oTabla.LabConf3 = "", Null, oTabla.LabConf3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf4", adVarChar, adParamInput, 1, IIf(oTabla.LabConf4 = "", Null, oTabla.LabConf4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf5", adVarChar, adParamInput, 1, IIf(oTabla.LabConf5 = "", Null, oTabla.LabConf5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConf6", adVarChar, adParamInput, 1, IIf(oTabla.LabConf6 = "", Null, oTabla.LabConf6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost1", adInteger, adParamInput, 0, IIf(oTabla.Diagnost1 = 0, Null, oTabla.Diagnost1))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost2", adInteger, adParamInput, 0, IIf(oTabla.Diagnost2 = 0, Null, oTabla.Diagnost2))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost3", adInteger, adParamInput, 0, IIf(oTabla.Diagnost3 = 0, Null, oTabla.Diagnost3))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost4", adInteger, adParamInput, 0, IIf(oTabla.Diagnost4 = 0, Null, oTabla.Diagnost4))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost5", adInteger, adParamInput, 0, IIf(oTabla.Diagnost5 = 0, Null, oTabla.Diagnost5))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Diagnost6", adInteger, adParamInput, 0, IIf(oTabla.Diagnost6 = 0, Null, oTabla.Diagnost6))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoEdad", adInteger, adParamInput, 0, IIf(oTabla.TipoEdad = 0, Null, oTabla.TipoEdad))
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
'   Descripción:    elimia o borra un registro a la tabla HIS_TEMPORAL
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoHis_Temporal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_TEMPORALEliminar"
   Set oParameter = .CreateParameter("@Codigo1", adVarChar, adParamInput, 6, oTabla.Codigo1)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_TEMPORAL
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoHis_Temporal) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_TEMPORALSeleccionarPorId"
     Set oParameter = .CreateParameter("@Codigo1", adVarChar, adParamInput, 6, oTabla.Codigo1)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.Codigo1 = IIf(IsNull(oRecordset!Codigo1), "", oRecordset!Codigo1)
       oTabla.Codigo2 = IIf(IsNull(oRecordset!Codigo2), "", oRecordset!Codigo2)
       oTabla.Codigo3 = IIf(IsNull(oRecordset!Codigo3), "", oRecordset!Codigo3)
       oTabla.Codigo4 = IIf(IsNull(oRecordset!Codigo4), "", oRecordset!Codigo4)
       oTabla.Codigo5 = IIf(IsNull(oRecordset!Codigo5), "", oRecordset!Codigo5)
       oTabla.Codigo6 = IIf(IsNull(oRecordset!Codigo6), "", oRecordset!Codigo6)
       oTabla.LabConf1 = IIf(IsNull(oRecordset!LabConf1), "", oRecordset!LabConf1)
       oTabla.LabConf2 = IIf(IsNull(oRecordset!LabConf2), "", oRecordset!LabConf2)
       oTabla.LabConf3 = IIf(IsNull(oRecordset!LabConf3), "", oRecordset!LabConf3)
       oTabla.LabConf4 = IIf(IsNull(oRecordset!LabConf4), "", oRecordset!LabConf4)
       oTabla.LabConf5 = IIf(IsNull(oRecordset!LabConf5), "", oRecordset!LabConf5)
       oTabla.LabConf6 = IIf(IsNull(oRecordset!LabConf6), "", oRecordset!LabConf6)
       oTabla.Diagnost1 = IIf(IsNull(oRecordset!Diagnost1), 0, oRecordset!Diagnost1)
       oTabla.Diagnost2 = IIf(IsNull(oRecordset!Diagnost2), 0, oRecordset!Diagnost2)
       oTabla.Diagnost3 = IIf(IsNull(oRecordset!Diagnost3), 0, oRecordset!Diagnost3)
       oTabla.Diagnost4 = IIf(IsNull(oRecordset!Diagnost4), 0, oRecordset!Diagnost4)
       oTabla.Diagnost5 = IIf(IsNull(oRecordset!Diagnost5), 0, oRecordset!Diagnost5)
       oTabla.Diagnost6 = IIf(IsNull(oRecordset!Diagnost6), 0, oRecordset!Diagnost6)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.TipoEdad = IIf(IsNull(oRecordset!TipoEdad), 0, oRecordset!TipoEdad)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoHis_Temporal)
   oTabla.Codigo1 = ""
   oTabla.Codigo2 = ""
   oTabla.Codigo3 = ""
   oTabla.Codigo4 = ""
   oTabla.Codigo5 = ""
   oTabla.Codigo6 = ""
   oTabla.LabConf1 = ""
   oTabla.LabConf2 = ""
   oTabla.LabConf3 = ""
   oTabla.LabConf4 = ""
   oTabla.LabConf5 = ""
   oTabla.LabConf6 = ""
   oTabla.Diagnost1 = 0
   oTabla.Diagnost2 = 0
   oTabla.Diagnost3 = 0
   oTabla.Diagnost4 = 0
   oTabla.Diagnost5 = 0
   oTabla.Diagnost6 = 0
   oTabla.Edad = 0
   oTabla.TipoEdad = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: HIS_TEMPORAL
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: HIS_TEMPORAL
'        Autor:   DAniel Barrantes B.
'        Fecha: 24/11/2013 07:02:12
'------------------------------------------------------------------------------------
 
