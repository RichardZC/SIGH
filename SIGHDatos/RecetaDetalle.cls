VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla RecetaDetalle
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
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
'   Descripción:    Inserta un registro a la tabla RecetaDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoRecetaDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleAgregar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadPedida", adInteger, adParamInput, 0, IIf(oTabla.CantidadPedida = 0, Null, oTabla.CantidadPedida))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoEnRegistroReceta", adInteger, adParamInput, 0, oTabla.SaldoEnRegistroReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoEnDespachoReceta", adInteger, adParamInput, 0, IIf(oTabla.SaldoEnDespachoReceta = 0, Null, oTabla.SaldoEnDespachoReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDosisRecetada", adInteger, adParamInput, 0, oTabla.idDosisRecetada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoDetalle", adInteger, adParamInput, 0, IIf(oTabla.idEstadoDetalle = 0, Null, oTabla.idEstadoDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoAnulacionMedico", adVarChar, adParamInput, 300, IIf(oTabla.MotivoAnulacionMedico = "", Null, oTabla.MotivoAnulacionMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 300, oTabla.Observaciones)
   .Parameters.Append oParameter
   
   'Actualizado 26092014
   Set oParameter = .CreateParameter("@IdViaAdministracion", adInteger, adParamInput, 0, IIf(oTabla.IdViaAdministracion = 0, Null, oTabla.IdViaAdministracion))
   .Parameters.Append oParameter
   '''''''''''''''''''''
   Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 20, IIf(oTabla.dx = "", Null, oTabla.dx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
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

Sub GrabaHistorico(oTabla As DoRecetaDetalle)
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
    lcSql = "select FechaReceta from recetaCabecera where idReceta=" & oTabla.IdReceta
    If oRsParametros.State = 1 Then oRsParametros.Close
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    '
    
    lcSql = "INSERT INTO RecetaDetalle" & _
"           (idReceta" & _
"           ,idItem" & _
"           ,CantidadPedida" & _
"           ,Precio" & _
"           ,CantidadDespachada" & _
"           ,idDosisRecetada" & _
"           ,idEstadoDetalle" & _
"           ,observaciones" & _
"           ,IdViaAdministracion,fecha0" & _
"           )" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdReceta & _
"           ," & oTabla.idItem & _
"           ," & oTabla.CantidadPedida & _
"           ," & oTabla.precio & _
"           ," & oTabla.CantidadDespachada & _
"           ," & oTabla.idDosisRecetada & _
"           ," & oTabla.idEstadoDetalle & _
"           ,'" & oTabla.Observaciones & _
"           '," & oTabla.IdViaAdministracion & ",'" & Format(oRsParametros!fechaReceta, "dd/mm/yyyy") & "'" & _
"           )"
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub




'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla RecetaDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoRecetaDetalle, Optional lbLaModificacionEsAnivelDeRecetaEitem As Boolean) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleModificar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadPedida", adInteger, adParamInput, 0, IIf(oTabla.CantidadPedida = 0, Null, oTabla.CantidadPedida))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoEnRegistroReceta", adInteger, adParamInput, 0, oTabla.SaldoEnRegistroReceta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoEnDespachoReceta", adInteger, adParamInput, 0, IIf(oTabla.SaldoEnDespachoReceta = 0, Null, oTabla.SaldoEnDespachoReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDosisRecetada", adInteger, adParamInput, 0, oTabla.idDosisRecetada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoDetalle", adInteger, adParamInput, 0, IIf(oTabla.idEstadoDetalle = 0, Null, oTabla.idEstadoDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoAnulacionMedico", adVarChar, adParamInput, 300, IIf(oTabla.MotivoAnulacionMedico = "", Null, oTabla.MotivoAnulacionMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Observaciones", adVarChar, adParamInput, 300, oTabla.Observaciones)
   .Parameters.Append oParameter
   
      'Actualizado 26092014
   Set oParameter = .CreateParameter("@IdViaAdministracion", adInteger, adParamInput, 0, IIf(oTabla.IdViaAdministracion = 0, Null, oTabla.IdViaAdministracion))
   .Parameters.Append oParameter
   '''''''''''''''''''''
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 20, IIf(oTabla.dx = "", Null, oTabla.dx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@WhereIdRecetaItem", adInteger, adParamInput, 0, IIf(lbLaModificacionEsAnivelDeRecetaEitem = True, 1, 0))
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
'   Descripción:    elimia o borra un registro a la tabla RecetaDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoRecetaDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleEliminar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
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
'   Descripción:    Selecciona un unico registro de la tabla RecetaDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoRecetaDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaCampos oTabla, oRecordset
   End If
   SeleccionarPorId = True
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoRecetaDetalle)
   oTabla.IdReceta = 0
   oTabla.idItem = 0
   oTabla.CantidadPedida = 0
   oTabla.precio = 0
   oTabla.Total = 0
   oTabla.SaldoEnRegistroReceta = 0
   oTabla.SaldoEnDespachoReceta = 0
   oTabla.CantidadDespachada = 0
End Sub


Function SeleccionarPorIdRecetaIditem(ByVal oTabla As DoRecetaDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdRecetaIditem = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaDetalleSeleccionarPorIdRecetaIditem"
     Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, oTabla.idItem)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaCampos oTabla, oRecordset
   End If
   SeleccionarPorIdRecetaIditem = True
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub CargaCampos(ByRef oTabla As DoRecetaDetalle, oRecordset As Recordset)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.idItem = IIf(IsNull(oRecordset!idItem), 0, oRecordset!idItem)
       oTabla.CantidadPedida = IIf(IsNull(oRecordset!CantidadPedida), 0, oRecordset!CantidadPedida)
       oTabla.precio = oRecordset!precio
       oTabla.Total = oRecordset!Total
       oTabla.SaldoEnRegistroReceta = IIf(IsNull(oRecordset!SaldoEnRegistroReceta), 0, oRecordset!SaldoEnRegistroReceta)
       oTabla.SaldoEnDespachoReceta = IIf(IsNull(oRecordset!SaldoEnDespachoReceta), 0, oRecordset!SaldoEnDespachoReceta)
       oTabla.CantidadDespachada = IIf(IsNull(oRecordset!CantidadDespachada), 0, oRecordset!CantidadDespachada)
       oTabla.idDosisRecetada = IIf(IsNull(oRecordset!idDosisRecetada), 0, oRecordset!idDosisRecetada)
       oTabla.idEstadoDetalle = IIf(IsNull(oRecordset!idEstadoDetalle), 0, oRecordset!idEstadoDetalle)
       oTabla.MotivoAnulacionMedico = IIf(IsNull(oRecordset!MotivoAnulacionMedico), "", oRecordset!MotivoAnulacionMedico)
       oTabla.Observaciones = IIf(IsNull(oRecordset!Observaciones), "", oRecordset!Observaciones)
       
       'Actualizado 26092014
       oTabla.IdViaAdministracion = IIf(IsNull(oRecordset!IdViaAdministracion), 0, oRecordset!IdViaAdministracion)
       '''''''''''''''''''''
       oTabla.dx = IIf(IsNull(oRecordset!dx), "", oRecordset!dx)
       oTabla.idPuntoCarga = IIf(IsNull(oRecordset!idPuntoCarga), 0, oRecordset!idPuntoCarga)
End Sub



Function RecetaDetalleActualizaCantDespachada(ByVal oTabla As DoRecetaDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
RecetaDetalleActualizaCantDespachada = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleActualizaCantDespachada"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoDetalle", adInteger, adParamInput, 0, IIf(oTabla.idEstadoDetalle = 0, Null, oTabla.idEstadoDetalle))
   .Parameters.Append oParameter
   .Execute
End With
 
RecetaDetalleActualizaCantDespachada = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function SeleccionarPorIdReceta(lnIdReceta As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, lnIdReceta)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdReceta = oRecordset
   
Exit Function
ManejadorDeError:
   MsgBox Err.Description
   
End Function

