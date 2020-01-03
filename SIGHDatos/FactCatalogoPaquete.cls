VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactCatalogoPaquete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactCatalogoPaquete
'        Programado por: Barrantes D
'        Fecha: Noviembre 2013
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
'   Descripción:    Inserta un registro a la tabla FactCatalogoPaquete
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFactCatalogoPaquete) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactCatalogoPaqueteAgregar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarWChar, adParamInput, 10, oTabla.codigo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarWChar, adParamInput, 100, oTabla.Descripcion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoPaquete", adInteger, adParamInput, 0, IIf(oTabla.TipoPaquete = 0, Null, oTabla.TipoPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@esItemFarmacia", adInteger, adParamInput, 0, oTabla.EsItemFarmacia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarWChar, adParamInput, 20, oTabla.cpt)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idFactPaquete = .Parameters("@idFactPaquete")
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
'   Descripción:    modifica un registro a la tabla FactCatalogoPaquete
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFactCatalogoPaquete) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactCatalogoPaqueteModificar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, IIf(oTabla.idFactPaquete = 0, Null, oTabla.idFactPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarWChar, adParamInput, 10, oTabla.codigo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarWChar, adParamInput, 100, oTabla.Descripcion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoPaquete", adInteger, adParamInput, 0, IIf(oTabla.TipoPaquete = 0, Null, oTabla.TipoPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@esItemFarmacia", adInteger, adParamInput, 0, oTabla.EsItemFarmacia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarWChar, adParamInput, 20, oTabla.cpt)
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
'   Descripción:    elimia o borra un registro a la tabla FactCatalogoPaquete
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFactCatalogoPaquete) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactCatalogoPaqueteEliminar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, oTabla.idFactPaquete)
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
'   Descripción:    Selecciona un unico registro de la tabla FactCatalogoPaquete
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFactCatalogoPaquete) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactCatalogoPaqueteSeleccionarPorId"
     Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, oTabla.idFactPaquete)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idFactPaquete = IIf(IsNull(oRecordset!idFactPaquete), 0, oRecordset!idFactPaquete)
       oTabla.codigo = oRecordset!codigo
       oTabla.Descripcion = oRecordset!Descripcion
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.TipoPaquete = IIf(IsNull(oRecordset!TipoPaquete), 0, oRecordset!TipoPaquete)
       oTabla.EsItemFarmacia = IIf(IsNull(oRecordset!EsItemFarmacia), 0, oRecordset!EsItemFarmacia)
       oTabla.cpt = IIf(IsNull(oRecordset!cpt), "", oRecordset!cpt)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFactCatalogoPaquete)
   oTabla.idFactPaquete = 0
   oTabla.codigo = ""
   oTabla.Descripcion = ""
   oTabla.idTipoFinanciamiento = 0
   oTabla.FechaCreacion = Now
   oTabla.IdUsuario = 0
   oTabla.IdEstado = 0
   oTabla.TipoPaquete = 0
End Sub
