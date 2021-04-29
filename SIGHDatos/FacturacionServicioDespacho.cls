VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionServicioDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionServicioDespacho
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionServicioDespacho
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFacturacionServicioDespacho) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDespachoAgregar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.cantidad)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   'mgaray201411a
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, oTabla.labConfHIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@grupoHIS", adInteger, adParamInput, 0, oTabla.GrupoHIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@subGrupoHIS", adInteger, adParamInput, 0, oTabla.SubGrupoHIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
   'RHA 17/11/2020 Cambio 46 Inicio
'    Set oParameter = .CreateParameter("@IdFinan", adInteger, adParamInput, 0, oTabla.IdFinan)
'   .Parameters.Append oParameter
'    Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.CodPres = "", Null, oTabla.CodPres))
'   .Parameters.Append oParameter
   'RHA 17/11/2020 Cambio 46 Fin
   'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)
   
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

Sub GrabaHistorico(oTabla As DoFacturacionServicioDespacho)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    '
    lcSql = "select FechaCreacion from factOrdenServicio where idOrden=" & oTabla.IdOrden
    If oRsParametros.State = 1 Then oRsParametros.Close
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    '
    lcSql = "INSERT INTO FacturacionServicioDespacho" & _
"           (idOrden,IdProducto" & _
"           ,Cantidad,Precio" & _
"           ,Total,labConfHIS" & _
"           ,GrupoHIS,SubGrupoHIS,fecha0)" & _
"     Values" & _
"           (" & _
"            " & oTabla.IdOrden & "," & oTabla.idProducto & _
"           , " & oTabla.cantidad & "," & oTabla.precio & _
"           , " & oTabla.Total & ",'" & oTabla.labConfHIS & "'" & _
"           , " & oTabla.GrupoHIS & "," & oTabla.SubGrupoHIS & ",'" & oRsParametros!FechaCreacion & _
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
'   Descripción:    modifica un registro a la tabla FacturacionServicioDespacho
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFacturacionServicioDespacho) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDespachoModificar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.cantidad = 0, Null, oTabla.cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@grupoHIS", adInteger, adParamInput, 0, oTabla.GrupoHIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@subGrupoHIS", adInteger, adParamInput, 0, oTabla.SubGrupoHIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
    'RHA 17/11/2020 Cambio 46 Inicio
   ' Set oParameter = .CreateParameter("@IdFinan", adInteger, adParamInput, 0, oTabla.IdFinan)
   '.Parameters.Append oParameter
   '  Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, IIf(oTabla.CodPres = "", Null, oTabla.CodPres))
   '.Parameters.Append oParameter
   'RHA 17/11/2020 Cambio 46 Fin
   'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)
   
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionServicioDespacho
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFacturacionServicioDespacho) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionServicioDespachoEliminar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionServicioDespacho
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFacturacionServicioDespacho) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServicioDespachoSeleccionarPorId"
     Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaTabla oTabla, oRecordset
'       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
'       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
'       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
'       oTabla.PRECIO = oRecordset!PRECIO
'       oTabla.TOTAL = oRecordset!TOTAL
'       oTabla.GrupoHIS = IIf(IsNull(oRecordset!Grupo), 0, oRecordset!Grupo)
'       oTabla.subGrupoHIS = IIf(IsNull(oRecordset!subgrupo), 0, oRecordset!subgrupo)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DoFacturacionServicioDespacho)
   oTabla.IdOrden = 0
   oTabla.idProducto = 0
   oTabla.cantidad = 0
   oTabla.precio = 0
   oTabla.Total = 0
End Sub

Function SeleccionarPorIdProducto(lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServicioDespachoSeleccionarPorIdProducto"
     Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   Set SeleccionarPorIdProducto = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
   
Exit Function
End Function


Sub CargaTabla(ByRef oTabla As DoFacturacionServicioDespacho, oRecordset As Recordset)
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.cantidad = IIf(IsNull(oRecordset!cantidad), 0, oRecordset!cantidad)
       oTabla.precio = oRecordset!precio
       oTabla.Total = oRecordset!Total
       oTabla.GrupoHIS = IIf(IsNull(oRecordset!GrupoHIS), 0, oRecordset!GrupoHIS)
       oTabla.SubGrupoHIS = IIf(IsNull(oRecordset!SubGrupoHIS), 0, oRecordset!SubGrupoHIS)
       oTabla.labConfHIS = IIf(IsNull(oRecordset!labConfHIS), 0, oRecordset!labConfHIS)
End Sub


Function SeleccionarPorIdOrdenIdProducto(ByVal oTabla As DoFacturacionServicioDespacho) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdOrdenIdProducto = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionServicioDespachoPorIdOrdenIProducto"
     Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oTabla.idProducto)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaTabla oTabla, oRecordset
'       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
'       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
'       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
'       oTabla.PRECIO = oRecordset!PRECIO
'       oTabla.TOTAL = oRecordset!TOTAL
'       oTabla.GrupoHIS = IIf(IsNull(oRecordset!Grupo), 0, oRecordset!Grupo)
'       oTabla.subGrupoHIS = IIf(IsNull(oRecordset!subgrupo), 0, oRecordset!subgrupo)
   SeleccionarPorIdOrdenIdProducto = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

