VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerinatalAtencionDx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PerinatalAtencionesDiagnosticos
'        Programado por: Barrantes D
'        Fecha: Agosto 2011
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
'   Descripción:    Inserta un registro a la tabla PerinatalAtencionesDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoPerinatalAtencionDx) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionDiagnosticosAgregar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idLista", adInteger, adParamInput, 0, IIf(oTabla.idLista = 0, Null, oTabla.idLista))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 20, IIf(oTabla.CodigoHIS = "", Null, oTabla.CodigoHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 20, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
   .Parameters.Append oParameter
   'mgaray201411a
   Set oParameter = .CreateParameter("@IdSubclasificacionDx", adVarChar, adParamInput, 20, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdClasificacionDx", adVarChar, adParamInput, 20, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx))
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
'   Descripción:    modifica un registro a la tabla PerinatalAtencionesDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPerinatalAtencionDx) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionDiagnosticosModificar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idLista", adInteger, adParamInput, 0, IIf(oTabla.idLista = 0, Null, oTabla.idLista))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 20, IIf(oTabla.CodigoHIS = "", Null, oTabla.CodigoHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 20, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
   .Parameters.Append oParameter
   'mgaray201411a
   Set oParameter = .CreateParameter("@IdSubclasificacionDx", adVarChar, adParamInput, 20, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdClasificacionDx", adVarChar, adParamInput, 20, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx))
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
'   Descripción:    elimia o borra un registro a la tabla PerinatalAtencionesDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPerinatalAtencionDx) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionDiagnosticosEliminar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla PerinatalAtencionesDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPerinatalAtencionDx) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PerinatalAtencionDiagnosticosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idPerinatalAtencion = IIf(IsNull(oRecordset!idPerinatalAtencion), 0, oRecordset!idPerinatalAtencion)
       oTabla.idModulo = IIf(IsNull(oRecordset!idModulo), 0, oRecordset!idModulo)
       oTabla.idLista = IIf(IsNull(oRecordset!idLista), 0, oRecordset!idLista)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.CodigoHIS = IIf(IsNull(oRecordset!CodigoHIS), "", oRecordset!CodigoHIS)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       'mgaray20141003
       oTabla.labConfHIS = IIf(IsNull(oRecordset!labConfHIS), "", oRecordset!labConfHIS)
       'mgaray201410e
       oTabla.ItemDiagnostico = IIf(IsNull(oRecordset!ItemDiagnostico), 0, oRecordset!ItemDiagnostico)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoPerinatalAtencionDx)
   oTabla.idPerinatalAtencion = 0
   oTabla.idModulo = 0
   oTabla.idLista = 0
   oTabla.IdDiagnostico = 0
   oTabla.CodigoHIS = ""
   oTabla.idAtencion = 0
   'mgaray20141003
   oTabla.labConfHIS = ""
   'mgaray201410e
   oTabla.ItemDiagnostico = 0
End Sub

Function PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion(mo_idPerinatalAtencion As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion"
        Set oParameter = .CreateParameter("@mo_idPerinatalAtencion", adInteger, adParamInput, 0, mo_idPerinatalAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PerinatalAtencionDxSeleccionarPorIdPerinatalAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function EliminarXatencion(ByVal oTabla As DoPerinatalAtencionDx) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarXatencion = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionDiagnosticosEliminarXidAtencion"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
EliminarXatencion = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


